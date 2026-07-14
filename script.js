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

function parseCSV(content, fileName) {
    const lines = content.split('\n').filter(line => line.trim() !== '');
    if (lines.length === 0) return [];
    
    const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
    
    const result = [];
    for (let i = 1; i < lines.length; i++) {
        const values = lines[i].split(',').map(v => v.trim().replace(/^"|"$/g, ''));
        if (values.length === headers.length) {
            const row = {};
            headers.forEach((header, index) => {
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

async function readAndParseFile(file, fileType) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = function(e) {
            try {
                let data;
                if (file.name.endsWith('.csv')) {
                    data = parseCSV(e.target.result, file.name);
                } else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
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

async function processAllFiles() {
    try {
        updateStatus('Starting file processing...', 'info');
        updateProgress(10);
        
        updateStatus('Processing Office Master file...', 'info');
        const masterFile = state.files.master;
        if (masterFile) {
            const masterData = await parseExcel(masterFile);
            state.data.master = masterData;
            updateStatus(`✅ Master file loaded: ${masterData.length} offices`, 'success');
        }
        updateProgress(20);
        
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

function updateFileStatus(fileId, status) {
    const statusElement = document.getElementById(`${fileId}-status`);
    if (statusElement) {
        statusElement.textContent = status === 'uploaded' ? '✅ Uploaded' : '❌ Not uploaded';
        statusElement.className = `file-status ${status === 'uploaded' ? 'uploaded' : 'not-uploaded'}`;
    }
}

function checkAllFilesUploaded() {
    const allUploaded = Object.values(state.files).every(file => file !== null);
    const generateBtn = document.getElementById('generateBtn');
    if (generateBtn) {
        generateBtn.disabled = !allUploaded;
    }
    return allUploaded;
}

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
    
    container.scrollTop = container.scrollHeight;
}

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

async function generateReport() {
    const generateBtn = document.getElementById('generateBtn');
    generateBtn.disabled = true;
    generateBtn.textContent = '⏳ Processing...';
    
    try {
        clearStatus();
        
        const success = await processAllFiles();
        if (!success) {
            throw new Error('File processing failed');
        }
        
        updateStatus('Validating data...', 'info');
        validateData();
        updateProgress(90);
        
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

function validateData() {
    const errors = [];
    
    if (!state.data.master || state.data.master.length === 0) {
        errors.push('Master file is empty or invalid');
    }
    
    const csvFiles = ['delRange', 'delSingle', 'dssRange', 'dssSingle', 'codRange', 'codSingle', 'lbRange', 'lbSingle'];
    for (const key of csvFiles) {
        if (!state.data[key] || state.data[key].length === 0) {
            console.warn(`${FILE_LABELS[key]} has no data or is empty`);
        }
    }
    
    if (errors.length > 0) {
        throw new Error(errors.join('; '));
    }
    
    updateStatus(`✅ Data validation passed: ${state.data.master.length} offices in master file`, 'success');
}

// ----------------------------
// 6. EXCEL GENERATION - FULL IMPLEMENTATION
// ----------------------------

function getOfficeData(data, officeId) {
    if (!data || !Array.isArray(data)) return null;
    return data.find(row => {
        const id = row['office-id'] || row['office_id'] || row['Office ID'];
        return id == officeId;
    }) || null;
}

function getOfficeDataGrouped(data, officeId) {
    if (!data || !Array.isArray(data)) return null;
    const rows = data.filter(row => {
        const id = row['Office ID'] || row['office_id'] || row['office-id'];
        return id == officeId;
    });
    if (rows.length === 0) return null;
    
    const totalCount = rows.reduce((sum, row) => sum + (row['Total COD Count'] || 0), 0);
    const digitalCount = rows.reduce((sum, row) => sum + (row['COD Digital Count'] || 0), 0);
    return { totalCount, digitalCount };
}

function calculateDeliveryPct(row) {
    if (!row) return null;
    const invoice = row['invoice-count'] || 0;
    const deposit = row['deposit-count'] || 0;
    if (invoice === 0) return null;
    return ((invoice - deposit) / invoice) * 100;
}

function calculateDSSPct(row) {
    if (!row) return null;
    const total = row['total_pdm_art_count'] || 0;
    const dss = row['total_dss_art_count'] || 0;
    if (total === 0) return null;
    return (dss / total) * 100;
}

function calculateCODPct(row) {
    if (!row) return null;
    const total = row.totalCount || 0;
    const digital = row.digitalCount || 0;
    if (total === 0) return null;
    return (digital / total) * 100;
}

function calculateAlertLevel(metrics) {
    const KPIS = {
        del: 90,
        dss: 90,
        cod: 80,
        lb: 100
    };
    
    let breaches = { period: false, single: false };
    
    const checks = [
        { key: 'delPctRange', kpi: KPIS.del, period: true },
        { key: 'delPctSingle', kpi: KPIS.del, period: false },
        { key: 'dssPctRange', kpi: KPIS.dss, period: true },
        { key: 'dssPctSingle', kpi: KPIS.dss, period: false },
        { key: 'codPctRange', kpi: KPIS.cod, period: true },
        { key: 'codPctSingle', kpi: KPIS.cod, period: false },
        { key: 'lbPctRange', kpi: KPIS.lb, period: true },
        { key: 'lbPctSingle', kpi: KPIS.lb, period: false }
    ];
    
    for (const check of checks) {
        const value = metrics[check.key];
        if (value !== null && value < check.kpi) {
            if (check.period) breaches.period = true;
            else breaches.single = true;
        }
    }
    
    if (breaches.period && breaches.single) return '🔴';
    if (breaches.period || breaches.single) return '🟡';
    return '🟢';
}

function buildMasterSorted() {
    const master = state.data.master;
    if (!master) return [];
    
    const subDivOrder = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj'];
    
    const sorted = [...master].sort((a, b) => {
        const idxA = subDivOrder.indexOf(a['Sub Division Name']);
        const idxB = subDivOrder.indexOf(b['Sub Division Name']);
        if (idxA !== idxB) return idxA - idxB;
        
        if (a['Sub Office Name'] < b['Sub Office Name']) return -1;
        if (a['Sub Office Name'] > b['Sub Office Name']) return 1;
        
        const typeOrder = { 'HPO': 0, 'SPO': 1, 'BPO': 2 };
        const typeA = typeOrder[a['Office Type Code']] || 3;
        const typeB = typeOrder[b['Office Type Code']] || 3;
        if (typeA !== typeB) return typeA - typeB;
        
        if (a['Office Name'] < b['Office Name']) return -1;
        if (a['Office Name'] > b['Office Name']) return 1;
        return 0;
    });
    
    return sorted;
}

function buildSheet1(masterSorted) {
    const headers = [
        'Alert Level', 'Sr. No.', 'Sub Division', 'Sub Office', 'Office Name', 'Office Type',
        'Postman', 'Invoice', 'Deposit', 'Delivery % (Period)',
        'Postman', 'Invoice', 'Deposit', 'Delivery % (Latest)',
        'PDM Articles', 'DSS Articles', 'DSS % (Period)',
        'PDM Articles', 'DSS Articles', 'DSS % (Latest)',
        'COD Total', 'COD Digital', 'COD Digital % (Period)',
        'COD Total', 'COD Digital', 'COD Digital % (Latest)',
        'Total LB', 'LB Clearance % (Period)', 'LB Clearance % (Latest)',
        'Office ID', 'Universe', 'Avg Cleared (Period)', 'Avg Cleared (Latest)'
    ];
    
    const rows = [headers];
    let srNo = 1;
    
    for (const office of masterSorted) {
        const officeId = office['Office ID'];
        const officeType = office['Office Type Code'];
        const universe = officeType === 'BPO' ? 'B' : 'A';
        
        const delRange = getOfficeData(state.data.delRange, officeId);
        const delSingle = getOfficeData(state.data.delSingle, officeId);
        const dssRange = getOfficeData(state.data.dssRange, officeId);
        const dssSingle = getOfficeData(state.data.dssSingle, officeId);
        const codRange = getOfficeDataGrouped(state.data.codRange, officeId);
        const codSingle = getOfficeDataGrouped(state.data.codSingle, officeId);
        const lbRange = getOfficeData(state.data.lbRange, officeId);
        const lbSingle = getOfficeData(state.data.lbSingle, officeId);
        
        const delPctRange = calculateDeliveryPct(delRange);
        const delPctSingle = calculateDeliveryPct(delSingle);
        const dssPctRange = calculateDSSPct(dssRange);
        const dssPctSingle = calculateDSSPct(dssSingle);
        const codPctRange = calculateCODPct(codRange);
        const codPctSingle = calculateCODPct(codSingle);
        const lbPctRange = lbRange ? lbRange['clearance-percentage'] : null;
        const lbPctSingle = lbSingle ? lbSingle['clearance-percentage'] : null;
        
        const alertLevel = calculateAlertLevel({
            delPctRange, delPctSingle,
            dssPctRange, dssPctSingle,
            codPctRange, codPctSingle,
            lbPctRange, lbPctSingle
        });
        
        const row = [
            alertLevel,
            srNo++,
            office['Sub Division Name'] || '',
            office['Sub Office Name'] || '',
            office['Office Name'] || '',
            office['Office Type Code'] || '',
            delRange ? delRange['postman-count'] || 0 : '',
            delRange ? delRange['invoice-count'] || 0 : '',
            delRange ? delRange['deposit-count'] || 0 : '',
            delPctRange !== null ? delPctRange : '',
            delSingle ? delSingle['postman-count'] || 0 : '',
            delSingle ? delSingle['invoice-count'] || 0 : '',
            delSingle ? delSingle['deposit-count'] || 0 : '',
            delPctSingle !== null ? delPctSingle : '',
            dssRange ? dssRange['total_pdm_art_count'] || 0 : '',
            dssRange ? dssRange['total_dss_art_count'] || 0 : '',
            dssPctRange !== null ? dssPctRange : '',
            dssSingle ? dssSingle['total_pdm_art_count'] || 0 : '',
            dssSingle ? dssSingle['total_dss_art_count'] || 0 : '',
            dssPctSingle !== null ? dssPctSingle : '',
            codRange ? codRange.totalCount || 0 : '',
            codRange ? codRange.digitalCount || 0 : '',
            codPctRange !== null ? codPctRange : '',
            codSingle ? codSingle.totalCount || 0 : '',
            codSingle ? codSingle.digitalCount || 0 : '',
            codPctSingle !== null ? codPctSingle : '',
            lbRange ? lbRange['total-letterboxes'] || 0 : '',
            lbPctRange !== null ? lbPctRange : '',
            lbPctSingle !== null ? lbPctSingle : '',
            officeId,
            universe,
            lbRange ? lbRange['avg-cleared'] || 0 : '',
            lbSingle ? lbSingle['avg-cleared'] || 0 : ''
        ];
        
        rows.push(row);
    }
    
    return rows;
}

function calculateSubDivisionSummary(offices) {
    const result = {
        delP: { defaulterCount: 0, sumInv: 0, sumDep: 0, pct: null },
        delL: { defaulterCount: 0, sumInv: 0, sumDep: 0, pct: null },
        dssP: { defaulterCount: 0, sumPdm: 0, sumDss: 0, pct: null },
        dssL: { defaulterCount: 0, sumPdm: 0, sumDss: 0, pct: null },
        codP: { defaulterCount: 0, sumTotal: 0, sumDigital: 0, pct: null },
        codL: { defaulterCount: 0, sumTotal: 0, sumDigital: 0, pct: null },
        lbP: { defaulterCount: 0, sumAvg: 0, sumBoxes: 0, pct: null },
        lbL: { defaulterCount: 0, sumAvg: 0, sumBoxes: 0, pct: null }
    };
    
    for (const office of offices) {
        const officeId = office['Office ID'];
        
        const delRange = getOfficeData(state.data.delRange, officeId);
        if (delRange) {
            const inv = delRange['invoice-count'] || 0;
            const dep = delRange['deposit-count'] || 0;
            result.delP.sumInv += inv;
            result.delP.sumDep += dep;
            const pct = inv > 0 ? ((inv - dep) / inv) * 100 : null;
            if (pct !== null && pct < 90) result.delP.defaulterCount++;
        }
        
        const delSingle = getOfficeData(state.data.delSingle, officeId);
        if (delSingle) {
            const inv = delSingle['invoice-count'] || 0;
            const dep = delSingle['deposit-count'] || 0;
            result.delL.sumInv += inv;
            result.delL.sumDep += dep;
            const pct = inv > 0 ? ((inv - dep) / inv) * 100 : null;
            if (pct !== null && pct < 90) result.delL.defaulterCount++;
        }
        
        const dssRange = getOfficeData(state.data.dssRange, officeId);
        if (dssRange) {
            const pdm = dssRange['total_pdm_art_count'] || 0;
            const dss = dssRange['total_dss_art_count'] || 0;
            result.dssP.sumPdm += pdm;
            result.dssP.sumDss += dss;
            const pct = pdm > 0 ? (dss / pdm) * 100 : null;
            if (pct !== null && pct < 90) result.dssP.defaulterCount++;
        }
        
        const dssSingle = getOfficeData(state.data.dssSingle, officeId);
        if (dssSingle) {
            const pdm = dssSingle['total_pdm_art_count'] || 0;
            const dss = dssSingle['total_dss_art_count'] || 0;
            result.dssL.sumPdm += pdm;
            result.dssL.sumDss += dss;
            const pct = pdm > 0 ? (dss / pdm) * 100 : null;
            if (pct !== null && pct < 90) result.dssL.defaulterCount++;
        }
        
        const codRange = getOfficeDataGrouped(state.data.codRange, officeId);
        if (codRange) {
            const total = codRange.totalCount || 0;
            const digital = codRange.digitalCount || 0;
            result.codP.sumTotal += total;
            result.codP.sumDigital += digital;
            const pct = total > 0 ? (digital / total) * 100 : null;
            if (pct !== null && pct < 80) result.codP.defaulterCount++;
        }
        
        const codSingle = getOfficeDataGrouped(state.data.codSingle, officeId);
        if (codSingle) {
            const total = codSingle.totalCount || 0;
            const digital = codSingle.digitalCount || 0;
            result.codL.sumTotal += total;
            result.codL.sumDigital += digital;
            const pct = total > 0 ? (digital / total) * 100 : null;
            if (pct !== null && pct < 80) result.codL.defaulterCount++;
        }
        
        const lbRange = getOfficeData(state.data.lbRange, officeId);
        if (lbRange) {
            const avg = lbRange['avg-cleared'] || 0;
            const boxes = lbRange['total-letterboxes'] || 0;
            result.lbP.sumAvg += avg;
            result.lbP.sumBoxes += boxes;
            const pct = boxes > 0 ? (avg / boxes) * 100 : null;
            if (pct !== null && pct < 100) result.lbP.defaulterCount++;
        }
        
        const lbSingle = getOfficeData(state.data.lbSingle, officeId);
        if (lbSingle) {
            const avg = lbSingle['avg-cleared'] || 0;
            const boxes = lbSingle['total-letterboxes'] || 0;
            result.lbL.sumAvg += avg;
            result.lbL.sumBoxes += boxes;
            const pct = boxes > 0 ? (avg / boxes) * 100 : null;
            if (pct !== null && pct < 100) result.lbL.defaulterCount++;
        }
    }
    
    if (result.delP.sumInv > 0) {
        result.delP.pct = ((result.delP.sumInv - result.delP.sumDep) / result.delP.sumInv) * 100;
    }
    if (result.delL.sumInv > 0) {
        result.delL.pct = ((result.delL.sumInv - result.delL.sumDep) / result.delL.sumInv) * 100;
    }
    if (result.dssP.sumPdm > 0) {
        result.dssP.pct = (result.dssP.sumDss / result.dssP.sumPdm) * 100;
    }
    if (result.dssL.sumPdm > 0) {
        result.dssL.pct = (result.dssL.sumDss / result.dssL.sumPdm) * 100;
    }
    if (result.codP.sumTotal > 0) {
        result.codP.pct = (result.codP.sumDigital / result.codP.sumTotal) * 100;
    }
    if (result.codL.sumTotal > 0) {
        result.codL.pct = (result.codL.sumDigital / result.codL.sumTotal) * 100;
    }
    if (result.lbP.sumBoxes > 0) {
        result.lbP.pct = (result.lbP.sumAvg / result.lbP.sumBoxes) * 100;
    }
    if (result.lbL.sumBoxes > 0) {
        result.lbL.pct = (result.lbL.sumAvg / result.lbL.sumBoxes) * 100;
    }
    
    return result;
}

function getPrevailingIssue(summary) {
    const issues = [];
    
    if (summary.delP.defaulterCount > 0 || summary.delL.defaulterCount > 0) {
        issues.push({ name: 'Delivery Returns', count: summary.delP.defaulterCount + summary.delL.defaulterCount });
    }
    if (summary.dssP.defaulterCount > 0 || summary.dssL.defaulterCount > 0) {
        issues.push({ name: 'DSS Adoption', count: summary.dssP.defaulterCount + summary.dssL.defaulterCount });
    }
    if (summary.codP.defaulterCount > 0 || summary.codL.defaulterCount > 0) {
        issues.push({ name: 'COD Cash', count: summary.codP.defaulterCount + summary.codL.defaulterCount });
    }
    if (summary.lbP.defaulterCount > 0 || summary.lbL.defaulterCount > 0) {
        issues.push({ name: 'LB Clearance', count: summary.lbP.defaulterCount + summary.lbL.defaulterCount });
    }
    
    if (issues.length === 0) return 'All KPIs Met ✅';
    
    issues.sort((a, b) => b.count - a.count);
    return issues[0].name;
}

function accumulateTotals(totals, summary) {
    totals.delP.count += summary.delP.defaulterCount;
    totals.delP.sumInv += summary.delP.sumInv;
    totals.delP.sumDep += summary.delP.sumDep;
    totals.delL.count += summary.delL.defaulterCount;
    totals.delL.sumInv += summary.delL.sumInv;
    totals.delL.sumDep += summary.delL.sumDep;
    
    totals.dssP.count += summary.dssP.defaulterCount;
    totals.dssP.sumPdm += summary.dssP.sumPdm;
    totals.dssP.sumDss += summary.dssP.sumDss;
    totals.dssL.count += summary.dssL.defaulterCount;
    totals.dssL.sumPdm += summary.dssL.sumPdm;
    totals.dssL.sumDss += summary.dssL.sumDss;
    
    totals.codP.count += summary.codP.defaulterCount;
    totals.codP.sumTotal += summary.codP.sumTotal;
    totals.codP.sumDigital += summary.codP.sumDigital;
    totals.codL.count += summary.codL.defaulterCount;
    totals.codL.sumTotal += summary.codL.sumTotal;
    totals.codL.sumDigital += summary.codL.sumDigital;
    
    totals.lbP.count += summary.lbP.defaulterCount;
    totals.lbP.sumAvg += summary.lbP.sumAvg;
    totals.lbP.sumBoxes += summary.lbP.sumBoxes;
    totals.lbL.count += summary.lbL.defaulterCount;
    totals.lbL.sumAvg += summary.lbL.sumAvg;
    totals.lbL.sumBoxes += summary.lbL.sumBoxes;
    
    if (totals.delP.sumInv > 0) {
        totals.delP.pct = ((totals.delP.sumInv - totals.delP.sumDep) / totals.delP.sumInv) * 100;
    }
    if (totals.delL.sumInv > 0) {
        totals.delL.pct = ((totals.delL.sumInv - totals.delL.sumDep) / totals.delL.sumInv) * 100;
    }
    if (totals.dssP.sumPdm > 0) {
        totals.dssP.pct = (totals.dssP.sumDss / totals.dssP.sumPdm) * 100;
    }
    if (totals.dssL.sumPdm > 0) {
        totals.dssL.pct = (totals.dssL.sumDss / totals.dssL.sumPdm) * 100;
    }
    if (totals.codP.sumTotal > 0) {
        totals.codP.pct = (totals.codP.sumDigital / totals.codP.sumTotal) * 100;
    }
    if (totals.codL.sumTotal > 0) {
        totals.codL.pct = (totals.codL.sumDigital / totals.codL.sumTotal) * 100;
    }
    if (totals.lbP.sumBoxes > 0) {
        totals.lbP.pct = (totals.lbP.sumAvg / totals.lbP.sumBoxes) * 100;
    }
    if (totals.lbL.sumBoxes > 0) {
        totals.lbL.pct = (totals.lbL.sumAvg / totals.lbL.sumBoxes) * 100;
    }
}

function buildSheet2(masterSorted) {
    const rows = [];
    
    rows.push(['Table A: Excluding B.Os (Universe A)']);
    rows.push(['Sr. No.', 'Sub Division', 'Offices <90% Del (P)', 'Offices <90% Del (L)', 'Delivery % (P)', 'Delivery % (L)', 
               'Prevailing Issue', 'Offices <90% DSS (P)', 'Offices <90% DSS (L)', 'DSS % (P)', 'DSS % (L)',
               'COD Digital % (P)', 'COD Digital % (L)', 'LB Clearance % (P)', 'LB Clearance % (L)']);
    
    const subDivs = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj'];
    let srNo = 1;
    let totalsA = { offices: 0, delP: { count: 0, sumInv: 0, sumDep: 0 }, delL: { count: 0, sumInv: 0, sumDep: 0 }, 
                   dssP: { count: 0, sumPdm: 0, sumDss: 0 }, dssL: { count: 0, sumPdm: 0, sumDss: 0 },
                   codP: { count: 0, sumTotal: 0, sumDigital: 0 }, codL: { count: 0, sumTotal: 0, sumDigital: 0 },
                   lbP: { count: 0, sumAvg: 0, sumBoxes: 0 }, lbL: { count: 0, sumAvg: 0, sumBoxes: 0 } };
    
    for (const subDiv of subDivs) {
        const offices = masterSorted.filter(o => o['Sub Division Name'] === subDiv && o['Office Type Code'] !== 'BPO');
        const summary = calculateSubDivisionSummary(offices);
        const prevailingIssue = getPrevailingIssue(summary);
        
        rows.push([
            srNo++,
            subDiv,
            summary.delP.defaulterCount,
            summary.delL.defaulterCount,
            summary.delP.pct !== null ? summary.delP.pct : '',
            summary.delL.pct !== null ? summary.delL.pct : '',
            prevailingIssue,
            summary.dssP.defaulterCount,
            summary.dssL.defaulterCount,
            summary.dssP.pct !== null ? summary.dssP.pct : '',
            summary.dssL.pct !== null ? summary.dssL.pct : '',
            summary.codP.pct !== null ? summary.codP.pct : '',
            summary.codL.pct !== null ? summary.codL.pct : '',
            summary.lbP.pct !== null ? summary.lbP.pct : '',
            summary.lbL.pct !== null ? summary.lbL.pct : ''
        ]);
        
        accumulateTotals(totalsA, summary);
    }
    
    rows.push(['', 'Division Total', totalsA.delP.count, totalsA.delL.count, 
               totalsA.delP.pct !== null ? totalsA.delP.pct : '', totalsA.delL.pct !== null ? totalsA.delL.pct : '',
               'All KPIs Met ✅', totalsA.dssP.count, totalsA.dssL.count,
               totalsA.dssP.pct !== null ? totalsA.dssP.pct : '', totalsA.dssL.pct !== null ? totalsA.dssL.pct : '',
               totalsA.codP.pct !== null ? totalsA.codP.pct : '', totalsA.codL.pct !== null ? totalsA.codL.pct : '',
               totalsA.lbP.pct !== null ? totalsA.lbP.pct : '', totalsA.lbL.pct !== null ? totalsA.lbL.pct : '']);
    
    rows.push([]);
    rows.push([]);
    
    rows.push(['Table B: Only B.Os (Universe B)']);
    rows.push(['Sr. No.', 'Sub Division', 'Offices <90% Del (P)', 'Offices <90% Del (L)', 'Delivery % (P)', 'Delivery % (L)', 
               'Prevailing Issue', 'Offices <90% DSS (P)', 'Offices <90% DSS (L)', 'DSS % (P)', 'DSS % (L)',
               'COD Digital % (P)', 'COD Digital % (L)', 'LB Clearance % (P)', 'LB Clearance % (L)']);
    
    srNo = 1;
    let totalsB = { offices: 0, delP: { count: 0, sumInv: 0, sumDep: 0 }, delL: { count: 0, sumInv: 0, sumDep: 0 },
                   dssP: { count: 0, sumPdm: 0, sumDss: 0 }, dssL: { count: 0, sumPdm: 0, sumDss: 0 },
                   codP: { count: 0, sumTotal: 0, sumDigital: 0 }, codL: { count: 0, sumTotal: 0, sumDigital: 0 },
                   lbP: { count: 0, sumAvg: 0, sumBoxes: 0 }, lbL: { count: 0, sumAvg: 0, sumBoxes: 0 } };
    
    for (const subDiv of subDivs) {
        const offices = masterSorted.filter(o => o['Sub Division Name'] === subDiv && o['Office Type Code'] === 'BPO');
        const summary = calculateSubDivisionSummary(offices);
        const prevailingIssue = getPrevailingIssue(summary);
        
        rows.push([
            srNo++,
            subDiv,
            summary.delP.defaulterCount,
            summary.delL.defaulterCount,
            summary.delP.pct !== null ? summary.delP.pct : '',
            summary.delL.pct !== null ? summary.delL.pct : '',
            prevailingIssue,
            summary.dssP.defaulterCount,
            summary.dssL.defaulterCount,
            summary.dssP.pct !== null ? summary.dssP.pct : '',
            summary.dssL.pct !== null ? summary.dssL.pct : '',
            summary.codP.pct !== null ? summary.codP.pct : '',
            summary.codL.pct !== null ? summary.codL.pct : '',
            '', '',
            ''
        ]);
        
        accumulateTotals(totalsB, summary);
    }
    
    rows.push(['', 'Division Total', totalsB.delP.count, totalsB.delL.count, 
               totalsB.delP.pct !== null ? totalsB.delP.pct : '', totalsB.delL.pct !== null ? totalsB.delL.pct : '',
               'All KPIs Met ✅', totalsB.dssP.count, totalsB.dssL.count,
               totalsB.dssP.pct !== null ? totalsB.dssP.pct : '', totalsB.dssL.pct !== null ? totalsB.dssL.pct : '',
               totalsB.codP.pct !== null ? totalsB.codP.pct : '', totalsB.codL.pct !== null ? totalsB.codL.pct : '',
               '', '']);
    
    return rows;
}

function countOfficesBelowAnyKPI(offices) {
    let count = 0;
    for (const office of offices) {
        const officeId = office['Office ID'];
        
        const delRange = getOfficeData(state.data.delRange, officeId);
        if (delRange) {
            const inv = delRange['invoice-count'] || 0;
            const dep = delRange['deposit-count'] || 0;
            const pct = inv > 0 ? ((inv - dep) / inv) * 100 : null;
            if (pct !== null && pct < 90) { count++; continue; }
        }
        
        const dssRange = getOfficeData(state.data.dssRange, officeId);
        if (dssRange) {
            const pdm = dssRange['total_pdm_art_count'] || 0;
            const dss = dssRange['total_dss_art_count'] || 0;
            const pct = pdm > 0 ? (dss / pdm) * 100 : null;
            if (pct !== null && pct < 90) { count++; continue; }
        }
        
        const codRange = getOfficeDataGrouped(state.data.codRange, officeId);
        if (codRange) {
            const total = codRange.totalCount || 0;
            const digital = codRange.digitalCount || 0;
            const pct = total > 0 ? (digital / total) * 100 : null;
            if (pct !== null && pct < 80) { count++; continue; }
        }
        
        const lbRange = getOfficeData(state.data.lbRange, officeId);
        if (lbRange) {
            const avg = lbRange['avg-cleared'] || 0;
            const boxes = lbRange['total-letterboxes'] || 0;
            const pct = boxes > 0 ? (avg / boxes) * 100 : null;
            if (pct !== null && pct < 100) { count++; continue; }
        }
    }
    return count;
}

function buildSheet3(masterSorted) {
    const rows = [];
    
    rows.push(['Sub Office Drill-Down: Performance by Scope']);
    rows.push(['Sr. No.', 'Sub Division', 'Sub Office', 'Scope', 
               'Offices <90% Del (P)', 'Offices <90% Del (L)', 'Delivery % (P)', 'Delivery % (L)',
               'Offices <90% DSS (P)', 'Offices <90% DSS (L)', 'DSS % (P)', 'DSS % (L)',
               'COD Digital % (P)', 'COD Digital % (L)', 'LB Clearance % (P)', 'LB Clearance % (L)',
               'Offices Below Any KPI']);
    
    let srNo = 1;
    const subDivs = ['ASP Karad West', 'SDIP Karad East', 'SDIP Vaduj'];
    
    const subOffices = {};
    for (const office of masterSorted) {
        const key = office['Sub Office Name'];
        if (!subOffices[key]) {
            subOffices[key] = [];
        }
        subOffices[key].push(office);
    }
    
    const subOfficeNames = Object.keys(subOffices).sort();
    for (const subOfficeName of subOfficeNames) {
        const offices = subOffices[subOfficeName];
        const subDiv = offices[0]['Sub Division Name'] || '';
        
        const exclOffices = offices.filter(o => o['Office Type Code'] !== 'BPO');
        const exclSummary = calculateSubDivisionSummary(exclOffices);
        const exclBelowAny = countOfficesBelowAnyKPI(exclOffices);
        
        rows.push([
            srNo,
            subDiv,
            subOfficeName,
            'Excl. B.Os',
            exclSummary.delP.defaulterCount,
            exclSummary.delL.defaulterCount,
            exclSummary.delP.pct !== null ? exclSummary.delP.pct : '',
            exclSummary.delL.pct !== null ? exclSummary.delL.pct : '',
            exclSummary.dssP.defaulterCount,
            exclSummary.dssL.defaulterCount,
            exclSummary.dssP.pct !== null ? exclSummary.dssP.pct : '',
            exclSummary.dssL.pct !== null ? exclSummary.dssL.pct : '',
            exclSummary.codP.pct !== null ? exclSummary.codP.pct : '',
            exclSummary.codL.pct !== null ? exclSummary.codL.pct : '',
            exclSummary.lbP.pct !== null ? exclSummary.lbP.pct : '',
            exclSummary.lbL.pct !== null ? exclSummary.lbL.pct : '',
            exclBelowAny
        ]);
        
        const onlyOffices = offices.filter(o => o['Office Type Code'] === 'BPO');
        const onlySummary = calculateSubDivisionSummary(onlyOffices);
        const onlyBelowAny = countOfficesBelowAnyKPI(onlyOffices);
        
        rows.push([
            srNo,
            subDiv,
            subOfficeName,
            'Only B.Os',
            onlySummary.delP.defaulterCount,
            onlySummary.delL.defaulterCount,
            onlySummary.delP.pct !== null ? onlySummary.delP.pct : '',
            onlySummary.delL.pct !== null ? onlySummary.delL.pct : '',
            onlySummary.dssP.defaulterCount,
            onlySummary.dssL.defaulterCount,
            onlySummary.dssP.pct !== null ? onlySummary.dssP.pct : '',
            onlySummary.dssL.pct !== null ? onlySummary.dssL.pct : '',
            onlySummary.codP.pct !== null ? onlySummary.codP.pct : '',
            onlySummary.codL.pct !== null ? onlySummary.codL.pct : '',
            '', '',
            onlyBelowAny
        ]);
        
        const inclSummary = calculateSubDivisionSummary(offices);
        const inclBelowAny = countOfficesBelowAnyKPI(offices);
        
        rows.push([
            srNo,
            subDiv,
            subOfficeName,
            'Incl. B.Os',
            inclSummary.delP.defaulterCount,
            inclSummary.delL.defaulterCount,
            inclSummary.delP.pct !== null ? inclSummary.delP.pct : '',
            inclSummary.delL.pct !== null ? inclSummary.delL.pct : '',
            inclSummary.dssP.defaulterCount,
            inclSummary.dssL.defaulterCount,
            inclSummary.dssP.pct !== null ? inclSummary.dssP.pct : '',
            inclSummary.dssL.pct !== null ? inclSummary.dssL.pct : '',
            inclSummary.codP.pct !== null ? inclSummary.codP.pct : '',
            inclSummary.codL.pct !== null ? inclSummary.codL.pct : '',
            inclSummary.lbP.pct !== null ? inclSummary.lbP.pct : '',
            inclSummary.lbL.pct !== null ? inclSummary.lbL.pct : '',
            inclBelowAny
        ]);
        
        srNo++;
        rows.push([]);
    }
    
    return rows;
}

function applySheet1Formatting(ws) {
    const colWidths = [
        { wch: 12 }, { wch: 8 }, { wch: 22 }, { wch: 22 }, { wch: 25 }, { wch: 10 },
        { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 16 },
        { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 16 },
        { wch: 12 }, { wch: 12 }, { wch: 16 },
        { wch: 12 }, { wch: 12 }, { wch: 16 },
        { wch: 12 }, { wch: 12 }, { wch: 16 },
        { wch: 12 }, { wch: 12 }, { wch: 16 },
        { wch: 12 }, { wch: 16 }, { wch: 16 },
        { wch: 12 }, { wch: 10 }, { wch: 14 }, { wch: 14 }
    ];
    ws['!cols'] = colWidths;
}

function applySheet2Formatting(ws) {
    ws['!cols'] = [
        { wch: 8 }, { wch: 25 }, { wch: 18 }, { wch: 18 }, 
        { wch: 16 }, { wch: 16 }, { wch: 20 }, 
        { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 },
        { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }
    ];
}

function applySheet3Formatting(ws) {
    ws['!cols'] = [
        { wch: 8 }, { wch: 22 }, { wch: 22 }, { wch: 14 },
        { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 },
        { wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 16 },
        { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 },
        { wch: 18 }
    ];
}

async function generateWorkbook() {
    const wb = XLSX.utils.book_new();
    
    const masterSorted = buildMasterSorted();
    
    const sheet1Data = buildSheet1(masterSorted);
    const ws1 = XLSX.utils.aoa_to_sheet(sheet1Data);
    applySheet1Formatting(ws1);
    XLSX.utils.book_append_sheet(wb, ws1, 'Performance Register');
    
    const sheet2Data = buildSheet2(masterSorted);
    const ws2 = XLSX.utils.aoa_to_sheet(sheet2Data);
    applySheet2Formatting(ws2);
    XLSX.utils.book_append_sheet(wb, ws2, 'Sub Division Summary');
    
    const sheet3Data = buildSheet3(masterSorted);
    const ws3 = XLSX.utils.aoa_to_sheet(sheet3Data);
    applySheet3Formatting(ws3);
    XLSX.utils.book_append_sheet(wb, ws3, 'Sub Office Details');
    
    return wb;
}

// ----------------------------
// 7. DOWNLOAD HANDLER
// ----------------------------

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

function init() {
    console.log('🏤 Karad Division Report Generator initialized');
    
    setupFileInputs();
    
    const generateBtn = document.getElementById('generateBtn');
    if (generateBtn) {
        generateBtn.addEventListener('click', generateReport);
    }
    
    const clearBtn = document.getElementById('clearBtn');
    if (clearBtn) {
        clearBtn.addEventListener('click', clearAllFiles);
    }
    
    const downloadBtn = document.getElementById('downloadBtn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', downloadWorkbook);
    }
    
    updateStatus('🚀 Ready. Upload all 9 files to begin.', 'info');
    checkAllFilesUploaded();
}

document.addEventListener('DOMContentLoaded', init);
