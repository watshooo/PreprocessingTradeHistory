// API Configuration
const API_BASE = window.location.origin;

// State Management
let uploadedFiles = {
    jisdor: null,
    tradeHistory: []
};

let selectedFiles = {
    jisdor: null,
    tradeHistory: []
};

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    console.log("✅ Page loaded, initializing...");
    initializeEventListeners();
    loadUploadedFiles();
    loadOutputFiles();
});

function initializeEventListeners() {
    // File inputs
    const jisdorInput = document.getElementById('jisdorFile');
    const tradeInput = document.getElementById('tradeFiles');
    
    if (jisdorInput) jisdorInput.addEventListener('change', handleJisdorFile);
    if (tradeInput) tradeInput.addEventListener('change', handleTradeFiles);
    
    // Buttons
    const uploadBtn = document.getElementById('uploadBtn');
    const processBtn = document.getElementById('processBtn');
    const refreshFilesBtn = document.getElementById('refreshFilesBtn');
    const refreshOutputsBtn = document.getElementById('refreshOutputsBtn');
    const cleanupFilesBtn = document.getElementById('cleanupFilesBtn');
    const cleanupOutputsBtn = document.getElementById('cleanupOutputsBtn');
    
    if (uploadBtn) uploadBtn.addEventListener('click', uploadFiles);
    if (processBtn) processBtn.addEventListener('click', processData);
    if (refreshFilesBtn) refreshFilesBtn.addEventListener('click', loadUploadedFiles);
    if (refreshOutputsBtn) refreshOutputsBtn.addEventListener('click', loadOutputFiles);
    if (cleanupFilesBtn) cleanupFilesBtn.addEventListener('click', cleanupAllFiles);
    if (cleanupOutputsBtn) cleanupOutputsBtn.addEventListener('click', cleanupAllOutputs);
    
    console.log("✅ Event listeners initialized");
}

// Handle file selection
function handleJisdorFile(e) {
    const file = e.target.files[0];
    if (file) {
        uploadedFiles.jisdor = file;
        displayFileInfo('jisdorFileInfo', file);
    }
}

function handleTradeFiles(e) {
    const files = Array.from(e.target.files);
    uploadedFiles.tradeHistory = files;
    displayMultipleFilesInfo('tradeFilesInfo', files);
}

function displayFileInfo(elementId, file) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    element.classList.add('active');
    element.innerHTML = `
        <p><strong>File:</strong> ${file.name}</p>
        <p><strong>Size:</strong> ${formatFileSize(file.size)}</p>
    `;
}

function displayMultipleFilesInfo(elementId, files) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    element.classList.add('active');
    element.innerHTML = `
        <p><strong>Selected:</strong> ${files.length} file(s)</p>
        ${files.map(f => `<p>• ${f.name} (${formatFileSize(f.size)})</p>`).join('')}
    `;
}

// Upload files to server
async function uploadFiles() {
    if (!uploadedFiles.jisdor) {
        showNotification('Please select JISDOR file', 'error');
        return;
    }

    if (uploadedFiles.tradeHistory.length === 0) {
        showNotification('Please select at least one trade history file', 'error');
        return;
    }

    showLoading(true);

    try {
        // Upload JISDOR file
        console.log("📁 Uploading JISDOR file...");
        const jisdorResult = await uploadFile(uploadedFiles.jisdor);
        selectedFiles.jisdor = jisdorResult.output_file;

        // Upload trade history files
        selectedFiles.tradeHistory = [];
        for (const file of uploadedFiles.tradeHistory) {
            console.log("📁 Uploading trade file:", file.name);
            const result = await uploadFile(file);
            selectedFiles.tradeHistory.push(result.output_file);
        }

        showNotification('Files uploaded successfully!', 'success');
        loadUploadedFiles();
        updateSelectedFilesInfo();
        
        const processBtn = document.getElementById('processBtn');
        if (processBtn) processBtn.disabled = false;

    } catch (error) {
        console.error('Upload error:', error);
        showNotification('Upload failed: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function uploadFile(file) {
    const formData = new FormData();
    formData.append('file', file);

    const response = await fetch(`${API_BASE}/api/upload`, {
        method: 'POST',
        body: formData
    });

    if (!response.ok) {
        throw new Error('Upload failed: ' + response.statusText);
    }

    return await response.json();
}

// Load uploaded files from server
async function loadUploadedFiles() {
    try {
        console.log("📂 Loading uploaded files...");
        const response = await fetch(`${API_BASE}/api/files`);
        const data = await response.json();

        if (data.success) {
            displayUploadedFiles(data.files || []);
        } else {
            console.error('Error loading files:', data.error);
        }
    } catch (error) {
        console.error('Failed to load files:', error);
    }
}

function displayUploadedFiles(files) {
    const container = document.getElementById('uploadedFilesList');
    if (!container) return;

    if (files.length === 0) {
        container.innerHTML = '<p class="empty-state">No files uploaded yet</p>';
        return;
    }

    container.innerHTML = files.map(file => `
        <div class="file-item" data-filename="${file.name}">
            <div class="file-item-info">
                <div class="file-item-name">${file.name}</div>
                <div class="file-item-meta">
                    Size: ${formatFileSize(file.size)} | 
                    Uploaded: ${formatDate(file.uploaded_at)}
                </div>
            </div>
            <div class="file-item-actions">
                <button class="btn btn-secondary btn-sm" onclick="selectFileForProcess('${file.name}')">
                    Select
                </button>
            </div>
        </div>
    `).join('');
}

function selectFileForProcess(filename) {
    console.log("✅ Selected file:", filename);
    
    // Simple selection logic
    if (filename.toLowerCase().includes('jisdor')) {
        selectedFiles.jisdor = filename;
    } else {
        if (!selectedFiles.tradeHistory.includes(filename)) {
            selectedFiles.tradeHistory.push(filename);
        }
    }
    updateSelectedFilesInfo();
    
    const processBtn = document.getElementById('processBtn');
    if (processBtn && selectedFiles.jisdor && selectedFiles.tradeHistory.length > 0) {
        processBtn.disabled = false;
    }
}

function updateSelectedFilesInfo() {
    const infoElement = document.getElementById('selectedFilesInfo');
    if (!infoElement) return;
    
    infoElement.classList.add('active');
    infoElement.innerHTML = `
        <p><strong>JISDOR File:</strong> ${selectedFiles.jisdor || 'Not selected'}</p>
        <p><strong>Trade History Files:</strong> ${selectedFiles.tradeHistory.length} file(s)</p>
        ${selectedFiles.tradeHistory.map(f => `<p>• ${f}</p>`).join('')}
    `;
}

// Process data
async function processData() {
    if (!selectedFiles.jisdor || selectedFiles.tradeHistory.length === 0) {
        showNotification('Please select JISDOR and trade history files', 'error');
        return;
    }

    showLoading(true);

    const requestData = {
        jisdor_file: selectedFiles.jisdor,
        trade_history_files: selectedFiles.tradeHistory,
        config: {
            rate_spot: parseFloat(document.getElementById('rateSpot').value),
            rate_remote: parseFloat(document.getElementById('rateRemote').value)
        }
    };

    console.log("⚙️  Processing with data:", requestData);

    try {
        const response = await fetch(`${API_BASE}/api/process`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(requestData)
        });

        const result = await response.json();

        if (result.success) {
            showNotification('Data processed successfully!', 'success');
            displayProcessLog(result.logs);
            loadOutputFiles();
        } else {
            showNotification('Processing failed: ' + result.error, 'error');
            displayProcessLog(result.logs);
        }

    } catch (error) {
        console.error('Processing error:', error);
        showNotification('Processing error: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

function displayProcessLog(logs) {
    if (!logs || logs.length === 0) return;

    const logElement = document.getElementById('processLog');
    if (!logElement) return;
    
    logElement.classList.add('active');
    logElement.innerHTML = logs.map(log => `<p>${log}</p>`).join('');
}

// Load output files
async function loadOutputFiles() {
    try {
        console.log("📂 Loading output files...");
        const response = await fetch(`${API_BASE}/api/outputs`);
        const data = await response.json();

        if (data.success) {
            displayOutputFiles(data.files || []);
        } else {
            console.error('Error loading outputs:', data.error);
        }
    } catch (error) {
        console.error('Failed to load output files:', error);
    }
}

function displayOutputFiles(files) {
    const container = document.getElementById('outputFilesList');
    if (!container) return;

    if (files.length === 0) {
        container.innerHTML = '<p class="empty-state">No output files yet</p>';
        return;
    }

    container.innerHTML = files.map(file => `
        <div class="file-item">
            <div class="file-item-info">
                <div class="file-item-name">📊 ${file.name}</div>
                <div class="file-item-meta">
                    Size: ${formatFileSize(file.size)} | 
                    Created: ${formatDate(file.created_at)}
                </div>
            </div>
            <div class="file-item-actions">
                <a href="${API_BASE}/api/download/${file.name}" 
                   class="btn btn-success btn-sm" 
                   download>
                    ⬇️ Download
                </a>
            </div>
        </div>
    `).join('');
}

// NEW: Cleanup functions
async function cleanupAllFiles() {
    if (!confirm('Are you sure you want to delete ALL uploaded files? This cannot be undone!')) {
        return;
    }

    showLoading(true);

    try {
        console.log("🗑️  Cleaning up all files...");
        const response = await fetch(`${API_BASE}/api/cleanup`, {
            method: 'DELETE'
        });

        const result = await response.json();

        if (result.success) {
            showNotification(`Deleted ${result.deleted} files`, 'success');
            console.log(`✅ Cleanup completed: ${result.deleted} files deleted`);
            
            // Clear selected files
            selectedFiles.jisdor = null;
            selectedFiles.tradeHistory = [];
            updateSelectedFilesInfo();
            
            // Reload lists
            loadUploadedFiles();
            loadOutputFiles();
        } else {
            showNotification('Cleanup failed', 'error');
        }

    } catch (error) {
        console.error('Cleanup error:', error);
        showNotification('Cleanup failed: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

async function cleanupAllOutputs() {
    if (!confirm('Are you sure you want to delete ALL output files? This cannot be undone!')) {
        return;
    }

    showLoading(true);

    try {
        console.log("🗑️  Cleaning up all outputs...");
        const response = await fetch(`${API_BASE}/api/cleanup`, {
            method: 'DELETE'
        });

        const result = await response.json();

        if (result.success) {
            showNotification(`Deleted ${result.deleted} files`, 'success');
            console.log(`✅ Cleanup completed: ${result.deleted} files deleted`);
            loadOutputFiles();
        } else {
            showNotification('Cleanup failed', 'error');
        }

    } catch (error) {
        console.error('Cleanup error:', error);
        showNotification('Cleanup failed: ' + error.message, 'error');
    } finally {
        showLoading(false);
    }
}

// Utility functions
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

function formatDate(dateString) {
    if (!dateString) return 'N/A';
    const date = new Date(dateString);
    return date.toLocaleString('id-ID', {
        year: 'numeric',
        month: 'short',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
    });
}

function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (!overlay) return;
    
    if (show) {
        overlay.classList.add('active');
    } else {
        overlay.classList.remove('active');
    }
}

function showNotification(message, type = 'info') {
    // Using browser alert for now
    const emoji = type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️';
    console.log(`${emoji} ${message}`);
    alert(`${emoji} ${message}`);
}
