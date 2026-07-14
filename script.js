// ============================================
// KARAD DIVISION REPORT GENERATOR
// Complete JavaScript Logic
// ============================================

// ----------------------------
// 1. STATE MANAGEMENT
// ----------------------------
const state = {
    files: {
        delRange: null,
        delSingle: null,
        dssRange: null,
        dssSingle: null,
        codRange: null,
        codSingle: null,
        lbRange: null,
        lbSingle: null,
        master: null
    },
    data: {
        delRange: null,
        delSingle: null,
        dssRange: null,
        dssSingle: null,
        codRange: null,
        codSingle: null,
        lbRange: null,
        lbSingle: null,
        master: null
    },
    processed: false,
    workbookData: null
};

// File input IDs mapping
const FILE_IDS = {
    delRange: 'delRange',
    delSingle: 'delSingle',
    dssRange: 'dssRange',
    dssSingle: 'dssSingle',
    codRange: 'codRange',
    codSingle: 'codSingle',
    lbRange: 'lbRange',
    lbSingle: 'lbSingle',
    master: 'masterFile'
};

const FILE_LABELS = {
    delRange: 'Delivery (Range)',
    delSingle: 'Delivery (Latest)',
    dssRange: 'DSS (Range)',
    dssSingle: 'DSS (Latest)',
    codRange: 'COD Digital (Range)',
    codSingle: 'COD Digital (Latest)',
    lbRange: 'LB Clearance (Range)',
    lbSingle: 'LB Clearance (Latest)',
    master: 'Office Master'
};

// ----------------------------
// 2. FILE PARSING FUNCTIONS
// ----------------------------

/**
 * Parse CSV file content to array of objects
 */
function parseCSV(content, fileName) {
    const lines = content.split('\n').filter(line => line.trim() !== '');
    if (lines.length === 0) return [];
    
    // Get headers from first line
    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    
    // Parse each data row
    const result = [];
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.trim().replace(/^"|"$/g, ''));
        if (values.length === headers.length) {
            const row = {};
            headers.forEach((header, index) => {
                // Try to convert numeric values
                const val = values[index];
                if (!isNaN(val) && val !== '') {
                    row[header] = parseFloat(val);
                } else {
                    row[header] = val;
                }
            });
            result.push(row);
        }
    }
    
    return result;
}

/**
 * Parse Excel file (XLSX) using SheetJS library
 */
function parseExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet);
                resolve(jsonData);
            } catch (error) {
                reject(error);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Read a file and parse it based on extension
 */
async function readAndParseFile(file, fileType) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                let data;
                if (file.name.endsWith('.csv')) {
                    data = parseCSV(e.target.result, file.name);
                } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
                    // For Excel files, we need to parse differently
                    // This is handled separately in the main flow
                    data = 'excel';
                }
                resolve(data);
            } catch (error) {
                reject(error);
            }
        };
        
        reader.onerror = reject;
        reader.readAsText(file);
    });
}

/**
 * Process all uploaded files
 */
async function processAllFiles() {
    try {
        updateStatus('Starting file processing...', 'info');
        updateProgress(10);
        
        // Process Master File (Excel)
        updateStatus('Processing Office Master file...', 'info');
        const masterFile = state.files.master;
        if (masterFile) {
            const masterData = await parseExcel(masterFile);
            state.data.master = masterData;
            updateStatus(`✅ Master file loaded: ${masterData.length} offices`, 'success');
        }
        updateProgress(20);
        
        // Process each CSV file
        const csvFiles = ['delRange', 'delSingle', 'dssRange', 'dssSingle', 'codRange', 'codSingle', 'lbRange', 'lbSingle'];
        let processed = 0;
        
        for (const key of csvFiles) {
            const file = state.files[key];
            if (file) {
                updateStatus(`Processing ${FILE_LABELS[key]}...`, 'info');
                const data = await readAndParseFile(file, key);
                state.data[key] = data;
                updateStatus(`✅ ${FILE_LABELS[key]}: ${data.length} records loaded`, 'success');
            }
            processed++;
            const progress = 20 + (processed / csvFiles.length) * 60;
            updateProgress(Math.min(progress, 80));
        }
        
        updateStatus('✅ All files processed successfully!', 'success');
        updateProgress(80);
        
        return true;
    } catch (error) {
        updateStatus(`❌ Error processing files: ${error.message}`, 'error');
        console.error('Processing error:', error);
        return false;
    }
}

// ----------------------------
// 3. UI UPDATE FUNCTIONS
// ----------------------------

/**
 * Update file status in UI
 */
function updateFileStatus(fileId, status) {
    const statusElement = document.getElementById(`${fileId}-status`);
    if (statusElement) {
        statusElement.textContent = status === 'uploaded' ? '✅ Uploaded' : '❌ Not uploaded';
        statusElement.className = `file-status ${status === 'uploaded' ? 'uploaded' : 'not-uploaded'}`;
    }
}

/**
 * Check if all files are uploaded
 */
function checkAllFilesUploaded() {
    const allUploaded = Object.values(state.files).every(file => file !== null);
    const generateBtn = document.getElementById('generateBtn');
    if (generateBtn) {
        generateBtn.disabled = !allUploaded;
    }
    return allUploaded;
}

/**
 * Update status messages
 */
function updateStatus(message, type = 'info') {
    const container = document.getElementById('statusMessages');
    if (!container) return;
    
    const statusArea = document.getElementById('statusArea');
    if (statusArea) {
        statusArea.style.display = 'block';
    }
    
    const messageEl = document.createElement('div');
    messageEl.className = `message ${type}`;
    const timestamp = new Date().toLocaleTimeString();
    messageEl.textContent = `[${timestamp}] ${message}`;
    container.appendChild(messageEl);
    
    // Auto-scroll to bottom
    container.scrollTop = container.scrollHeight;
}

/**
 * Update progress bar
 */
function updateProgress(percent) {
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    if (progressFill) {
        progressFill.style.width = `${percent}%`;
    }
    if (progressText) {
        progressText.textContent = `${Math.round(percent)}%`;
    }
}

/**
 * Clear all status messages
 */
function clearStatus() {
    const container = document.getElementById('statusMessages');
    if (container) {
        container.innerHTML = '';
    }
    const statusArea = document.getElementById('statusArea');
    if (statusArea) {
        statusArea.style.display = 'none';
    }
    updateProgress(0);
}

/**
 * Show download area
 */
function showDownloadArea(summary) {
    const downloadArea = document.getElementById('downloadArea');
    if (downloadArea) {
        downloadArea.style.display = 'block';
        const summaryEl = document.getElementById('downloadSummary');
        if (summaryEl && summary) {
            summaryEl.textContent = summary;
        }
    }
}

// ----------------------------
// 4. FILE INPUT HANDLERS
// ----------------------------

/**
 * Set up file input handlers
 */
function setupFileInputs() {
    Object.keys(FILE_IDS).forEach(key => {
        const elementId = FILE_IDS[key];
        const input = document.getElementById(elementId);
        if (input) {
            input.addEventListener('change', function(e) {
                const file = this.files[0];
                if (file) {
                    state.files[key] = file;
                    updateFileStatus(elementId, 'uploaded');
                    updateStatus(`📎 ${FILE_LABELS[key]}: ${file.name} uploaded`, 'info');
                } else {
                    state.files[key] = null;
                    updateFileStatus(elementId, 'not-uploaded');
                }
                checkAllFilesUploaded();
            });
        }
    });
}

/**
 * Clear all files
 */
function clearAllFiles() {
    Object.keys(FILE_IDS).forEach(key => {
        const elementId = FILE_IDS[key];
        const input = document.getElementById(elementId);
        if (input) {
            input.value = '';
        }
        state.files[key] = null;
        state.data[key] = null;
        updateFileStatus(elementId, 'not-uploaded');
    });
    
    clearStatus();
    
    const downloadArea = document.getElementById('downloadArea');
    if (downloadArea) {
        downloadArea.style.display = 'none';
    }
    
    state.processed = false;
    state.workbookData = null;
    
    checkAllFilesUploaded();
    updateStatus('🗑️ All files cleared', 'info');
}

// ----------------------------
// 5. GENERATE REPORT - MAIN FUNCTION
// ----------------------------

/**
 * Main function to generate the report
 */
async function generateReport() {
    const generateBtn = document.getElementById('generateBtn');
    generateBtn.disabled = true;
    generateBtn.textContent = '⏳ Processing...';
    
    try {
        // Clear previous status
        clearStatus();
        
        // Process all files
        const success = await processAllFiles();
        if (!success) {
            throw new Error('File processing failed');
        }
        
        // Validate data
        updateStatus('Validating data...', 'info');
        validateData();
        updateProgress(90);
        
        // Generate workbook
        updateStatus('Generating Excel workbook...', 'info');
        const workbook = await generateWorkbook();
        state.workbookData = workbook;
        updateProgress(100);
        
        updateStatus('✅ Report generation complete!', 'success');
        showDownloadArea(`Generated report with ${state.data.master ? state.data.master.length : 0} offices`);
        
        state.processed = true;
        
    } catch (error) {
        updateStatus(`❌ Error generating report: ${error.message}`, 'error');
        console.error('Generation error:', error);
    } finally {
        generateBtn.disabled = false;
        generateBtn.textContent = '🚀 Generate Report';
    }
}

/**
 * Validate loaded data
 */
function validateData() {
    const errors = [];
    
    if (!state.data.master || state.data.master.length === 0) {
        errors.push('Master file is empty or invalid');
    }
    
    // Check each CSV
    const csvFiles = ['delRange', 'delSingle', 'dssRange', 'dssSingle', 'codRange', 'codSingle', 'lbRange', 'lbSingle'];
    for (const key of csvFiles) {
        if (!state.data[key] || state.data[key].length === 0) {
            // Not all offices will have data in all files - this is expected
            // Just log it as a warning
            console.warn(`${FILE_LABELS[key]} has no data or is empty`);
        }
    }
    
    if (errors.length > 0) {
        throw new Error(errors.join('; '));
    }
    
    updateStatus(`✅ Data validation passed: ${state.data.master.length} offices in master file`, 'success');
}

// ----------------------------
// 6. EXCEL GENERATION (PLACEHOLDER - WILL BE EXPANDED IN STEP 6)
// ----------------------------

/**
 * Generate Excel workbook - placeholder for now
 * Will be fully implemented in Step 6
 */
async function generateWorkbook() {
    // This is a placeholder - will be expanded in Step 6
    // For now, we'll create a simple workbook
    const wb = XLSX.utils.book_new();
    
    // Create a simple summary sheet
    const summaryData = [
        ['Karad Division Report'],
        ['Generated:', new Date().toLocaleString()],
        [''],
        ['Total Offices:', state.data.master ? state.data.master.length : 0],
        ['Source Files:', 'All 9 files uploaded']
    ];
    
    const ws = XLSX.utils.aoa_to_sheet(summaryData);
    XLSX.utils.book_append_sheet(wb, ws, 'Summary');
    
    return wb;
}

// ----------------------------
// 7. DOWNLOAD HANDLER
// ----------------------------

/**
 * Download the generated workbook
 */
function downloadWorkbook() {
    if (!state.workbookData) {
        alert('Please generate the report first!');
        return;
    }
    
    try {
        const wb = state.workbookData;
        const filename = `Karad_Division_Report_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, filename);
        updateStatus(`📥 Downloaded: ${filename}`, 'success');
    } catch (error) {
        updateStatus(`❌ Download error: ${error.message}`, 'error');
        console.error('Download error:', error);
    }
}

// ----------------------------
// 8. INITIALIZATION
// ----------------------------

/**
 * Initialize the application
 */
function init() {
    console.log('🏤 Karad Division Report Generator initialized');
    
    // Set up file inputs
    setupFileInputs();
    
    // Set up generate button
    const generateBtn = document.getElementById('generateBtn');
    if (generateBtn) {
        generateBtn.addEventListener('click', generateReport);
    }
    
    // Set up clear button
    const clearBtn = document.getElementById('clearBtn');
    if (clearBtn) {
        clearBtn.addEventListener('click', clearAllFiles);
    }
    
    // Set up download button
    const downloadBtn = document.getElementById('downloadBtn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', downloadWorkbook);
    }
    
    // Initial status
    updateStatus('🚀 Ready. Upload all 9 files to begin.', 'info');
    checkAllFilesUploaded();
}

// Run initialization when DOM is ready
document.addEventListener('DOMContentLoaded', init);
