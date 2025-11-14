#!/usr/bin/env python3
"""
Trade History Dashboard - Python Processor
Wrapper untuk memanggil fungsi-fungsi dari dashboard_v6_dengan_jenis_produk.py
FIX: Handle encoding issue di Windows (UTF-8)
"""

import argparse
import sys
import os
import pandas as pd
import io

# FIX: Set UTF-8 encoding untuk stdout di Windows
if sys.platform == 'win32':
    # Redirect output ke UTF-8
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

# PENTING: Setup path untuk import module Python original
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from dashboard_v6_dengan_jenis_produk import (
        load_jisdor,
        process_file,
        write_output,
        buat_rekap_volume,
        buat_breakdown_volume,
        buat_nilai_transaksi_rp,
        buat_nilai_transaksi_usd,
        buat_margin_transaksi,
        MONTH_MAP,
        MONTH_REV,
        MONTH_NAME_ID,
        CONTRACT_SIZE_PER_LOT
    )
    print("[OK] Successfully imported functions from dashboard_v6_dengan_jenis_produk.py")
except ImportError as e:
    print(f"[ERROR] Error importing functions: {e}")
    print(f"        Current directory: {os.getcwd()}")
    print(f"        Script directory: {os.path.dirname(os.path.abspath(__file__))}")
    sys.exit(1)


def main():
    """
    Main function - Process trade history files
    """
    parser = argparse.ArgumentParser(
        description='Process trade history data and generate dashboard Excel'
    )
    parser.add_argument('--jisdor', required=True, help='Path to JISDOR Excel file')
    parser.add_argument('--output', required=True, help='Output Excel file path')
    parser.add_argument('--rate-spot', type=float, default=5000000, 
                       help='Spot rate for margin calculation (default: 5,000,000)')
    parser.add_argument('--rate-remote', type=float, default=3500000, 
                       help='Remote rate for margin calculation (default: 3,500,000)')
    parser.add_argument('--trade-file', action='append', required=True, 
                       help='Trade history Excel file(s) - can be multiple')
    
    args = parser.parse_args()
    
    print("=" * 70)
    print("[START] TRADE HISTORY DASHBOARD - PYTHON PROCESSOR")
    print("=" * 70)
    
    try:
        # Validasi file JISDOR ada
        if not os.path.exists(args.jisdor):
            raise FileNotFoundError(f"JISDOR file tidak ditemukan: {args.jisdor}")
        
        # Validasi semua trade files ada
        for trade_file in args.trade_file:
            if not os.path.exists(trade_file):
                raise FileNotFoundError(f"Trade history file tidak ditemukan: {trade_file}")
        
        print(f"[INFO] JISDOR file: {os.path.basename(args.jisdor)}")
        print(f"[INFO] Trade history files: {len(args.trade_file)} file(s)")
        print(f"[INFO] Rate Spot: {args.rate_spot:,.0f} Rp")
        print(f"[INFO] Rate Remote: {args.rate_remote:,.0f} Rp")
        print(f"[INFO] Output file: {os.path.basename(args.output)}")
        print("-" * 70)
        
        # 1. Load JISDOR exchange rate data
        print("[STEP 1] Loading JISDOR exchange rate data...")
        kurs_df = load_jisdor(args.jisdor)
        print(f"[OK] Loaded {len(kurs_df)} rows of JISDOR data")
        print(f"[OK] Date range: {kurs_df['Tanggal'].min().date()} to {kurs_df['Tanggal'].max().date()}")
        
        # 2. Process all trade history files
        print(f"\n[STEP 2] Processing {len(args.trade_file)} trade history file(s)...")
        all_data = []
        sheet_map = {}
        
        for i, trade_file in enumerate(args.trade_file, 1):
            filename = os.path.basename(trade_file)
            print(f"\n[FILE {i}/{len(args.trade_file)}] Processing: {filename}")
            
            try:
                df, sheet_name = process_file(
                    trade_file,
                    kurs_df,
                    rate_spot=args.rate_spot,
                    rate_remote=args.rate_remote
                )
                
                if df is None or df.empty:
                    print(f"[WARN] No valid data in {filename}")
                    continue
                
                print(f"[OK] Processed {len(df)} transactions")
                print(f"[OK] Sheet name: {sheet_name}")
                
                if sheet_name:
                    sheet_map[sheet_name] = df
                all_data.append(df)
                
            except Exception as e:
                print(f"[ERROR] Error processing {filename}: {str(e)}")
                continue
        
        # 3. Combine all data
        if not all_data:
            print("\n[ERROR] No valid data to process")
            sys.exit(1)
        
        dashboard_df = pd.concat(all_data, ignore_index=True)
        print(f"\n[OK] Combined {len(all_data)} file(s) into dashboard")
        print(f"[OK] Total transactions: {len(dashboard_df)}")
        
        # 4. Generate Excel output
        print(f"\n[STEP 3] Generating Excel output...")
        write_output(dashboard_df, sheet_map, args.output)
        
        print(f"[OK] Output saved: {args.output}")
        print("=" * 70)
        print("[SUCCESS] Processing completed successfully!")
        print("=" * 70)
        
        return 0
        
    except FileNotFoundError as e:
        print(f"\n[ERROR] File not found: {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"\n[ERROR] {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    sys.exit(main())
