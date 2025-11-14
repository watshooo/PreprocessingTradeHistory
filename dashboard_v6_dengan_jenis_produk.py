import pandas as pd
import numpy as np
import glob
import os
from datetime import datetime, timedelta
from openpyxl.utils import get_column_letter

# === KONSTANTA === #
MONTH_MAP = {
    'JAN':1, 'FEB':2, 'MAR':3, 'APR':4, 'MAY':5, 'JUN':6,
    'JUL':7, 'AUG':8, 'SEP':9, 'OCT':10, 'NOV':11, 'DEC':12
}
MONTH_REV = {v:k for k, v in MONTH_MAP.items()}

MONTH_NAME_ID = {
    1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April',
    5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus',
    9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
}

CONTRACT_SIZE_PER_LOT = 25000  # kg

# === FUNGSI TAMBAHAN: Ekstrak Jenis Produk dari Contract === #
def ekstrak_jenis_produk(contract_name):
    """
    Ekstrak jenis produk dari nama contract.
    Contoh: 'CPOID-JAN25' → 'CPOID'
    """
    if pd.isna(contract_name):
        return 'Unknown'
    try:
        jenis = str(contract_name).split('-')[0].strip().upper()
        return jenis if jenis else 'Unknown'
    except:
        return 'Unknown'


# === 1️⃣ Fungsi Bantu Perhitungan === #
def hitung_NV(row):
    price = row['Price']
    lot = row['Vol(LOT)']
    if pd.isna(price) or pd.isna(lot):
        return np.nan
    return float(lot) * CONTRACT_SIZE_PER_LOT * float(price)

def hitung_contract_size(row):
    lot = row['Vol(LOT)']
    if pd.isna(lot):
        return np.nan
    return float(lot) * CONTRACT_SIZE_PER_LOT

def hitung_margin(row, rate_spot, rate_remote):
    lot = row['Vol(LOT)']
    date_trade = row['DateTrade']
    contract_suffix = str(row['Contract']).split('-')[-1]

    try:
        bulan_kontrak = MONTH_MAP[contract_suffix[:3].upper()]
        tahun_kontrak = 2000 + int(contract_suffix[3:])
        
        if bulan_kontrak == 1:
            bulan_sebelum = 12
            tahun_sebelum = tahun_kontrak - 1
        else:
            bulan_sebelum = bulan_kontrak - 1
            tahun_sebelum = tahun_kontrak
        
        start_spot = datetime(tahun_sebelum, bulan_sebelum, 16)
        end_spot = datetime(tahun_kontrak, bulan_kontrak, 15)
        
        if start_spot <= date_trade <= end_spot:
            margin_per_sisi = lot * rate_spot
        else:
            margin_per_sisi = lot * rate_remote
        
        return margin_per_sisi * 2
        
    except Exception:
        return lot * rate_remote * 2

def cari_kolom(nama_kolom, df, return_letter=False):
    if nama_kolom in df.columns:
        idx = df.columns.get_loc(nama_kolom)
        if return_letter:
            col_letter = get_column_letter(idx + 1)
            return f"{col_letter}:{col_letter}"
        return idx
    else:
        raise ValueError(f"Kolom '{nama_kolom}' tidak ada di DataFrame")

# === 2️⃣ Fungsi untuk Baca & Siapkan Data Kurs JISDOR === #
def load_jisdor(file_path):
    kurs_df = pd.read_excel(file_path, skiprows=4, header=0)
    kurs_df = kurs_df[[c for c in kurs_df.columns if not c.startswith("Unnamed")]]
    kurs_df.columns = [col.strip().capitalize() for col in kurs_df.columns]

    if 'Tanggal' not in kurs_df.columns or 'Kurs' not in kurs_df.columns:
        raise ValueError("File kurs JISDOR harus memiliki kolom 'Tanggal' dan 'Kurs'.")

    kurs_df['Tanggal'] = pd.to_datetime(kurs_df['Tanggal'])
    kurs_df = kurs_df.sort_values('Tanggal').reset_index(drop=True)
    kurs_df['Kurs'] = pd.to_numeric(kurs_df['Kurs'], errors='coerce')
    kurs_df['Kurs'] = kurs_df['Kurs'].ffill()

    return kurs_df

def padankan_kurs(df_trade, kurs_df):
    kurs_df = kurs_df.sort_values('Tanggal')
    df_trade = df_trade.sort_values('DateTrade')

    df_merged = pd.merge_asof(
        df_trade,
        kurs_df,
        left_on='DateTrade',
        right_on='Tanggal',
        direction='backward'
    )

    df_merged.rename(columns={
        'Kurs': 'Kurs_Jisdor',
        'Tanggal': 'Tanggal_Kurs'
    }, inplace=True)
    
    if 'No' in df_merged.columns:
        df_merged.drop(columns=['No'], inplace=True)
    
    df_merged['Notional_Value_USD'] = df_merged['Notional_Value'] / df_merged['Kurs_Jisdor']

    return df_merged

# === 3️⃣ Fungsi Proses File === #
def process_file(file_path, kurs_df, rate_spot=5_000_000, rate_remote=3_500_000):
    print(f"Membaca file: {os.path.basename(file_path)}")
    df = pd.read_excel(file_path, header=0)
    df = df[1:].reset_index(drop=True)

    df.columns = [
        'DateTrade', 'Trade ID', 'Contract', 'Acc.Buy', 'Mbr.Buy',
        'Acc.Sell', 'Mbr.Sell', 'Currency', 'Price', 'Unit',
        'Vol(LOT)', 'ClosePosition'
    ]

    df['DateTrade'] = pd.to_datetime(df['DateTrade'])
    df['Price'] = pd.to_numeric(df['Price'], errors='coerce')
    df['Vol(LOT)'] = pd.to_numeric(df['Vol(LOT)'], errors='coerce')
    
    # ✨ TAMBAHAN: Ekstrak Jenis_Produk dari Contract
    df['Jenis_Produk'] = df['Contract'].apply(ekstrak_jenis_produk)
    
    df['Contract_Size_KG'] = df.apply(hitung_contract_size, axis=1)
    df['Notional_Value'] = df.apply(hitung_NV, axis=1)
    df['Margin'] = df.apply(
        hitung_margin,
        axis=1,
        rate_spot=rate_spot,
        rate_remote=rate_remote
    )

    df = padankan_kurs(df, kurs_df)

    if not df.empty:
        sample_date = df['DateTrade'].iloc[0]
        sheet_name = f"{MONTH_REV[sample_date.month]}{str(sample_date.year)[-2:]}"
    else:
        sheet_name = None

    return df, sheet_name

# === 4️⃣ Fungsi Proses Folder === #
def process_folder(input_folder, kurs_df, pattern='*.xlsx', **kwargs):
    files = glob.glob(os.path.join(input_folder, pattern))
    if not files:
        raise FileNotFoundError(f"Tidak ada file dengan pola {pattern} di folder {input_folder}")

    all_data = []
    sheet_map = {}

    for file_path in files:
        df, sheet_name = process_file(file_path, kurs_df, **kwargs)
        if sheet_name:
            sheet_map[sheet_name] = df
            all_data.append(df)

    dashboard_df = pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    return dashboard_df, sheet_map

# === 5️⃣ Fungsi Buat Rekap Volume === #
def buat_rekap_volume(dashboard_df):
    """Buat rekap volume per bulan."""
    if dashboard_df.empty or 'DateTrade' not in dashboard_df.columns:
        return pd.DataFrame({'Bulan': [], 'Volume_Lot': []}), ""
    
    dashboard_df['Bulan_Num'] = dashboard_df['DateTrade'].dt.month
    dashboard_df['Tahun'] = dashboard_df['DateTrade'].dt.year
    
    rekap = dashboard_df.groupby('Bulan_Num')['Vol(LOT)'].sum().reset_index()
    rekap['Bulan'] = rekap['Bulan_Num'].map(MONTH_NAME_ID)
    rekap = rekap[['Bulan', 'Vol(LOT)']].rename(columns={'Vol(LOT)': 'Volume_Lot'})
    
    total_volume = rekap['Volume_Lot'].sum()
    total_row = pd.DataFrame({'Bulan': ['Total'], 'Volume_Lot': [total_volume]})
    rekap = pd.concat([rekap, total_row], ignore_index=True)
    
    min_year = dashboard_df['Tahun'].min()
    max_year = dashboard_df['Tahun'].max()
    tahun_str = str(min_year) if min_year == max_year else f"{min_year}-{max_year}"
    
    return rekap, tahun_str

# === 6️⃣ Fungsi Buat Breakdown Volume === #
def buat_breakdown_volume(dashboard_df):
    """
    Buat breakdown volume transaksi per jenis produk dan tahun.
    """
    if dashboard_df.empty or 'Contract' not in dashboard_df.columns:
        return pd.DataFrame(), "", []
    
    dashboard_df['Jenis_Produk'] = dashboard_df['Contract'].apply(ekstrak_jenis_produk)
    dashboard_df['Tahun'] = dashboard_df['DateTrade'].dt.year
    
    breakdown = dashboard_df.groupby(['Jenis_Produk', 'Tahun'])['Vol(LOT)'].sum().reset_index()
    
    pivot = breakdown.pivot(index='Jenis_Produk', columns='Tahun', values='Vol(LOT)').fillna(0)
    pivot = pivot.sort_index(axis=1)
    list_tahun = sorted(pivot.columns.tolist())
    
    pivot = pivot.reset_index()
    
    if len(list_tahun) >= 2:
        tahun_awal = list_tahun[0]
        tahun_akhir = list_tahun[-1]
        
        pivot['Perubahan (%)'] = pivot.apply(
            lambda row: ((row[tahun_akhir] - row[tahun_awal]) / row[tahun_awal] * 100)
            if row[tahun_awal] > 0 else 0,
            axis=1
        )
    else:
        pivot['Perubahan (%)'] = 0
    
    total_row = {'Jenis_Produk': 'Total Lot'}
    for tahun in list_tahun:
        total_row[tahun] = pivot[tahun].sum()
    
    if len(list_tahun) >= 2:
        tahun_awal = list_tahun[0]
        tahun_akhir = list_tahun[-1]
        total_row['Perubahan (%)'] = (
            (total_row[tahun_akhir] - total_row[tahun_awal]) / total_row[tahun_awal] * 100
            if total_row[tahun_awal] > 0 else 0
        )
    else:
        total_row['Perubahan (%)'] = 0
    
    pivot = pd.concat([pivot, pd.DataFrame([total_row])], ignore_index=True)
    
    min_year = min(list_tahun)
    max_year = max(list_tahun)
    tahun_str = str(min_year) if min_year == max_year else f"{min_year}-{max_year}"
    
    return pivot, tahun_str, list_tahun

# === 7️⃣ Fungsi Buat Nilai Transaksi RP === #
def buat_nilai_transaksi_rp(dashboard_df):
    """Buat sheet Nilai_Transaksi_RP dengan total notional value per bulan dalam Rupiah."""
    if dashboard_df.empty or 'Notional_Value' not in dashboard_df.columns:
        return pd.DataFrame({'Bulan': [], 'Nilai Transaksi RP': []}), ""
    
    if 'Bulan_Num' not in dashboard_df.columns:
        dashboard_df['Bulan_Num'] = dashboard_df['DateTrade'].dt.month
    if 'Tahun' not in dashboard_df.columns:
        dashboard_df['Tahun'] = dashboard_df['DateTrade'].dt.year
    
    nilai_rp = dashboard_df.groupby('Bulan_Num')['Notional_Value'].sum().reset_index()
    nilai_rp['Bulan'] = nilai_rp['Bulan_Num'].map(MONTH_NAME_ID)
    nilai_rp = nilai_rp[['Bulan', 'Notional_Value']].rename(
        columns={'Notional_Value': 'Nilai Transaksi RP'}
    )
    
    total_nilai = nilai_rp['Nilai Transaksi RP'].sum()
    total_row = pd.DataFrame({'Bulan': ['Total'], 'Nilai Transaksi RP': [total_nilai]})
    nilai_rp = pd.concat([nilai_rp, total_row], ignore_index=True)
    
    min_year = dashboard_df['Tahun'].min()
    max_year = dashboard_df['Tahun'].max()
    tahun_str = str(min_year) if min_year == max_year else f"{min_year}-{max_year}"
    
    return nilai_rp, tahun_str

# === 8️⃣ Fungsi Buat Nilai Transaksi USD === #
def buat_nilai_transaksi_usd(dashboard_df):
    """Buat sheet Nilai_transaksi_USD dengan total notional value per bulan dalam USD."""
    if dashboard_df.empty or 'Notional_Value_USD' not in dashboard_df.columns:
        return pd.DataFrame({'Bulan': [], 'Nilai Transaksi (USD)': []}), ""
    
    if 'Bulan_Num' not in dashboard_df.columns:
        dashboard_df['Bulan_Num'] = dashboard_df['DateTrade'].dt.month
    if 'Tahun' not in dashboard_df.columns:
        dashboard_df['Tahun'] = dashboard_df['DateTrade'].dt.year
    
    nilai_usd = dashboard_df.groupby('Bulan_Num')['Notional_Value_USD'].sum().reset_index()
    nilai_usd['Bulan'] = nilai_usd['Bulan_Num'].map(MONTH_NAME_ID)
    nilai_usd = nilai_usd[['Bulan', 'Notional_Value_USD']].rename(
        columns={'Notional_Value_USD': 'Nilai Transaksi (USD)'}
    )
    
    total_nilai = nilai_usd['Nilai Transaksi (USD)'].sum()
    total_row = pd.DataFrame({'Bulan': ['Total'], 'Nilai Transaksi (USD)': [total_nilai]})
    nilai_usd = pd.concat([nilai_usd, total_row], ignore_index=True)
    
    min_year = dashboard_df['Tahun'].min()
    max_year = dashboard_df['Tahun'].max()
    tahun_str = str(min_year) if min_year == max_year else f"{min_year}-{max_year}"
    
    return nilai_usd, tahun_str

# === 9️⃣ Fungsi Buat Margin Transaksi === #
def buat_margin_transaksi(dashboard_df):
    """Buat sheet Margin_Transaksi dengan total margin per bulan dalam Rupiah."""
    if dashboard_df.empty or 'Margin' not in dashboard_df.columns:
        return pd.DataFrame({'Bulan': [], 'Margin Transaksi (Rp)': []}), ""
    
    if 'Bulan_Num' not in dashboard_df.columns:
        dashboard_df['Bulan_Num'] = dashboard_df['DateTrade'].dt.month
    if 'Tahun' not in dashboard_df.columns:
        dashboard_df['Tahun'] = dashboard_df['DateTrade'].dt.year
    
    margin = dashboard_df.groupby('Bulan_Num')['Margin'].sum().reset_index()
    margin['Bulan'] = margin['Bulan_Num'].map(MONTH_NAME_ID)
    margin = margin[['Bulan', 'Margin']].rename(
        columns={'Margin': 'Margin Transaksi (Rp)'}
    )
    
    total_margin = margin['Margin Transaksi (Rp)'].sum()
    total_row = pd.DataFrame({'Bulan': ['Total'], 'Margin Transaksi (Rp)': [total_margin]})
    margin = pd.concat([margin, total_row], ignore_index=True)
    
    min_year = dashboard_df['Tahun'].min()
    max_year = dashboard_df['Tahun'].max()
    tahun_str = str(min_year) if min_year == max_year else f"{min_year}-{max_year}"
    
    return margin, tahun_str

# === 🔟 Fungsi Output ke Excel === #
def write_output(dashboard_df, sheet_map, output_file):
    """
    Tulis Excel dengan urutan sheet:
    1. Rekap_Volume_Transaksi
    2. Breakdown_Volume_Transaksi
    3. Nilai_Transaksi_RP
    4. Nilai_transaksi_USD
    5. Margin_Transaksi
    6. Dashboard (dengan Jenis_Produk)
    7. Sheet bulanan (JAN25, FEB25, dst dengan Jenis_Produk)
    """
    def parse_sheet_order(name):
        month_str = name[:3].upper()
        year_str = name[3:]
        month_num = MONTH_MAP[month_str]
        year_num = 2000 + int(year_str)
        return (year_num, month_num)

    sorted_sheets = sorted(sheet_map.items(), key=lambda x: parse_sheet_order(x[0]))

    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        # 1️⃣ Sheet Rekap Volume
        rekap_df, tahun_str_rekap = buat_rekap_volume(dashboard_df)
        rekap_df.to_excel(writer, index=False, sheet_name='Rekap_Volume_Transaksi', startrow=2)
        
        # 2️⃣ Sheet Breakdown Volume
        breakdown_df, tahun_str_breakdown, list_tahun = buat_breakdown_volume(dashboard_df)
        breakdown_df.to_excel(writer, index=False, sheet_name='Breakdown_Volume_Transaksi', startrow=3)
        
        # 3️⃣ Sheet Nilai Transaksi RP
        nilai_rp_df, tahun_str_rp = buat_nilai_transaksi_rp(dashboard_df)
        nilai_rp_df.to_excel(writer, index=False, sheet_name='Nilai_Transaksi_RP', startrow=2)
        
        # 4️⃣ Sheet Nilai Transaksi USD
        nilai_usd_df, tahun_str_usd = buat_nilai_transaksi_usd(dashboard_df)
        nilai_usd_df.to_excel(writer, index=False, sheet_name='Nilai_transaksi_USD', startrow=2)
        
        # 5️⃣ Sheet Margin Transaksi
        margin_df, tahun_str_margin = buat_margin_transaksi(dashboard_df)
        margin_df.to_excel(writer, index=False, sheet_name='Margin_Transaksi', startrow=2)
        
        # 6️⃣ Sheet Dashboard (dengan Jenis_Produk)
        dashboard_df.to_excel(writer, index=False, sheet_name='Dashboard')

        # 7️⃣ Sheet bulanan (dengan Jenis_Produk sudah ada dari process_file)
        for sheet_name, df_month in sorted_sheets:
            df_month.to_excel(writer, index=False, sheet_name=sheet_name)

        # === Format Excel === #
        workbook = writer.book
        fmt_decimal = workbook.add_format({'align':'right', 'num_format':'#,##0.00'})
        fmt_integer = workbook.add_format({'align':'right', 'num_format':'#,##0'})
        fmt_percent = workbook.add_format({'align':'right', 'num_format':'0.00%'})
        fmt_bold = workbook.add_format({'bold': True, 'align':'right', 'num_format':'#,##0'})
        fmt_bold_decimal = workbook.add_format({'bold': True, 'align':'right', 'num_format':'#,##0.00'})
        fmt_title = workbook.add_format({
            'bold': True, 'font_size': 14, 'align': 'center', 'valign': 'vcenter'
        })
        fmt_header_center = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#D9E1F2'
        })
        
        # === Format Sheet Rekap_Volume_Transaksi === #
        ws_rekap = writer.sheets['Rekap_Volume_Transaksi']
        judul_rekap = f"VOLUME TRANSAKSI PERIODE TAHUN {tahun_str_rekap}"
        ws_rekap.merge_range('A1:B1', judul_rekap, fmt_title)
        ws_rekap.set_column('A:A', 20)
        ws_rekap.set_column('B:B', 15, fmt_integer)
        last_row_rekap = len(rekap_df) + 2
        ws_rekap.write(last_row_rekap, 1, rekap_df.iloc[-1]['Volume_Lot'], fmt_bold)
        
        # === Format Sheet Breakdown_Volume_Transaksi === #
        ws_breakdown = writer.sheets['Breakdown_Volume_Transaksi']
        
        judul_breakdown = f"VOLUME TRANSAKSI PERIODE {tahun_str_breakdown}"
        num_cols = len(list_tahun) + 2
        ws_breakdown.merge_range(0, 0, 0, num_cols - 1, judul_breakdown, fmt_title)
        
        ws_breakdown.write(2, 0, 'Jenis Produk', fmt_header_center)
        for idx, tahun in enumerate(list_tahun, start=1):
            ws_breakdown.write(2, idx, 'Lot', fmt_header_center)
        ws_breakdown.write(2, num_cols - 1, '', fmt_header_center)
        
        ws_breakdown.set_column(0, 0, 18)
        for idx in range(1, len(list_tahun) + 1):
            ws_breakdown.set_column(idx, idx, 12, fmt_integer)
        ws_breakdown.set_column(num_cols - 1, num_cols - 1, 15, fmt_decimal)
        
        last_row_breakdown = len(breakdown_df) + 3
        for col_idx in range(1, num_cols):
            cell_value = breakdown_df.iloc[-1, col_idx]
            if col_idx == num_cols - 1:
                ws_breakdown.write(last_row_breakdown, col_idx, cell_value / 100, fmt_bold)
            else:
                ws_breakdown.write(last_row_breakdown, col_idx, cell_value, fmt_bold)
        
        for row_idx in range(4, last_row_breakdown):
            cell_value = breakdown_df.iloc[row_idx - 4, num_cols - 1]
            ws_breakdown.write(row_idx, num_cols - 1, cell_value / 100, fmt_percent)
        
        # === Format Sheet Nilai_Transaksi_RP === #
        ws_nilai_rp = writer.sheets['Nilai_Transaksi_RP']
        judul_nilai_rp = f"Notional Value Rupiah Transaksi Periode {tahun_str_rp}"
        ws_nilai_rp.merge_range('A1:B1', judul_nilai_rp, fmt_title)
        ws_nilai_rp.set_column('A:A', 20)
        ws_nilai_rp.set_column('B:B', 25, fmt_decimal)
        last_row_nilai_rp = len(nilai_rp_df) + 2
        ws_nilai_rp.write(last_row_nilai_rp, 1, nilai_rp_df.iloc[-1]['Nilai Transaksi RP'], fmt_bold_decimal)
        
        # === Format Sheet Nilai_transaksi_USD === #
        ws_nilai_usd = writer.sheets['Nilai_transaksi_USD']
        judul_nilai_usd = f"Notional Value (USD) Transaksi Periode {tahun_str_usd}"
        ws_nilai_usd.merge_range('A1:B1', judul_nilai_usd, fmt_title)
        ws_nilai_usd.set_column('A:A', 20)
        ws_nilai_usd.set_column('B:B', 25, fmt_decimal)
        last_row_nilai_usd = len(nilai_usd_df) + 2
        ws_nilai_usd.write(last_row_nilai_usd, 1, nilai_usd_df.iloc[-1]['Nilai Transaksi (USD)'], fmt_bold_decimal)
        
        # === Format Sheet Margin_Transaksi === #
        ws_margin = writer.sheets['Margin_Transaksi']
        judul_margin = f"Margin Transaksi Rupiah Periode {tahun_str_margin}"
        ws_margin.merge_range('A1:B1', judul_margin, fmt_title)
        ws_margin.set_column('A:A', 20)
        ws_margin.set_column('B:B', 25, fmt_decimal)
        last_row_margin = len(margin_df) + 2
        ws_margin.write(last_row_margin, 1, margin_df.iloc[-1]['Margin Transaksi (Rp)'], fmt_bold_decimal)
        
        # === Format Sheet Dashboard dan Bulanan === #
        for sheet_name in writer.sheets:
            if sheet_name in ['Rekap_Volume_Transaksi', 'Breakdown_Volume_Transaksi', 
                             'Nilai_Transaksi_RP', 'Nilai_transaksi_USD', 'Margin_Transaksi']:
                continue
                
            worksheet = writer.sheets[sheet_name]
            
            if 'Notional_Value_USD' in dashboard_df.columns:
                col_range = cari_kolom('Notional_Value_USD', dashboard_df, True)
                worksheet.set_column(col_range, 20, fmt_decimal)
            
            if 'Contract_Size_KG' in dashboard_df.columns:
                col_range = cari_kolom('Contract_Size_KG', dashboard_df, True)
                worksheet.set_column(col_range, 18, fmt_integer)
            
            # ✨ TAMBAHAN: Format kolom Jenis_Produk jika ada
            if 'Jenis_Produk' in dashboard_df.columns:
                col_range = cari_kolom('Jenis_Produk', dashboard_df, True)
                worksheet.set_column(col_range, 15)

    print(f"✅ Selesai. File output: {output_file}")
    print(f"📊 Total sheet yang dibuat: {len(writer.sheets)}")

# === 1️⃣1️⃣ Main Routine === #
def main():
    input_folder = 'D:/cod/testDat/trade_history'
    kurs_file = 'D:/cod/testDat/Informasi_Kurs_Jisdor.xlsx'
    output_file = 'dashboard_v6_with_jenis_produk.xlsx'

    print("=" * 60)
    print("🚀 MEMULAI PROSES PENGOLAHAN DATA TRADE HISTORY")
    print("   [UPGRADE V6] Dengan Kolom Jenis_Produk di Setiap Sheet")
    print("=" * 60)
    
    kurs_df = load_jisdor(kurs_file)
    print(f"✅ Kurs JISDOR berhasil dimuat: {len(kurs_df)} baris")

    dashboard_df, sheet_map = process_folder(
        input_folder,
        kurs_df,
        rate_spot=5_000_000,
        rate_remote=3_500_000
    )
    
    print(f"✅ Data dashboard berhasil dikompilasi: {len(dashboard_df)} transaksi")
    print(f"✅ Sheet bulanan yang dibuat: {len(sheet_map)} sheet")
    print(f"✅ Kolom 'Jenis_Produk' ditambahkan ke semua sheet bulanan")

    write_output(dashboard_df, sheet_map, output_file)
    
    print("=" * 60)
    print("🎉 PROSES SELESAI!")
    print(f"📁 File output tersimpan: {output_file}")
    print("=" * 60)

if __name__ == '__main__':
    main()
