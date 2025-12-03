// app.js

let sheetNames = [];
let selectedSheets = [];
let allMatches = []; // Stores the results from the last successful match run
let sheetDataCache = {}; // Cache to store full sheet data fetched from /api/data

/**
 * Utility to fetch data from the Go API.
 */
async function fetchData(url, options = {}) {
    const response = await fetch(url, options);
    const result = await response.json();

    if (response.status !== 200) {
        throw new Error(result.error || 'API Error: ' + url);
    }
    return result;
}

// ---------------------------------------------------------------------
// --- File Upload and Sheet Selection ---
// ---------------------------------------------------------------------

window.handleFileUpload = async function() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) return;

    document.getElementById('fileNameDisplay').textContent = 'Uploading ' + file.name + '...';

    const formData = new FormData();
    formData.append('excelFile', file);

    try {
        const result = await fetchData('/api/upload', { method: 'POST', body: formData });
        
        sheetNames = result.sheetNames;
        sheetDataCache = {}; // Clear cache on new file
        allMatches = []; // Clear matches
        
        document.getElementById('fileNameDisplay').textContent = 'File uploaded successfully.';
        initializeSheetSelector(sheetNames);
        document.getElementById('exportControls').style.display = 'none';
        document.getElementById('results').innerHTML = '';
        document.getElementById('visualizationSection').style.display = 'none';

    } catch (error) {
        alert('Upload failed: ' + error.message);
        document.getElementById('fileNameDisplay').textContent = 'Upload failed.';
    }
}

function initializeSheetSelector(names) {
    const selectorDiv = document.getElementById('sheetSelector');
    selectorDiv.innerHTML = '';
    selectedSheets = []; 

    names.forEach(name => {
        const btn = document.createElement('div');
        btn.className = 'sheet-option';
        btn.textContent = name;
        btn.setAttribute('data-sheet-name', name); 
        btn.onclick = (e) => toggleSheetSelection(name, e.target);
        selectorDiv.appendChild(btn);
    });
    document.getElementById('matchControls').style.display = 'block';
    updateSelectionSlots();
}

window.toggleSheetSelection = function(sheetName, element) {
    const index = selectedSheets.indexOf(sheetName);
    
    if (index > -1) {
        selectedSheets.splice(index, 1);
        element.classList.remove('selected');
    } else { 
        if (selectedSheets.length < 2) {
            selectedSheets.push(sheetName); 
            element.classList.add('selected');
        } else {
            const oldSheet = selectedSheets.pop();
            document.querySelector('.sheet-option[data-sheet-name="' + oldSheet + '"]').classList.remove('selected'); 
            selectedSheets.push(sheetName);
            element.classList.add('selected');
        }
    }
    updateSelectionSlots();
}

function updateSelectionSlots() {
    document.getElementById('slot1Name').textContent = selectedSheets[0] || 'Click a sheet name...';
    document.getElementById('slot2Name').textContent = selectedSheets[1] || 'Click a second sheet name...';

    const matchBtn = document.getElementById('matchBtn');
    const isReady = selectedSheets.length === 2;
    
    matchBtn.disabled = !isReady;
    matchBtn.onclick = isReady ? runMatch : null;

    if (isReady) {
        // Fetch column headers for export configuration
        populateExportColumns(selectedSheets[0], 1);
        populateExportColumns(selectedSheets[1], 2);
    } else {
        document.getElementById('exportIgnoreCols1Container').innerHTML = '<h3>Sheet 1 Columns</h3>';
        document.getElementById('exportIgnoreCols2Container').innerHTML = '<h3>Sheet 2 Columns</h3>';
        document.getElementById('exportControls').style.display = 'none';
        document.getElementById('exportAllBtn').disabled = true;
    }
}

// ---------------------------------------------------------------------
// --- Column Selection for Export ---
// ---------------------------------------------------------------------

/**
 * Fetches column headers and populates the "Ignore Columns" selector.
 */
async function populateExportColumns(sheetName, slotIndex) {
    const containerId = 'exportIgnoreCols' + slotIndex + 'Container';
    const container = document.getElementById(containerId);

    try {
        if (!sheetDataCache[sheetName]) {
            const data = await fetchData('/api/data/' + sheetName);
            sheetDataCache[sheetName] = data;
        }

        const headers = sheetDataCache[sheetName].headers;
        let html = '<h3>' + sheetName + ' Columns (Ignore on Export)</h3>';

        headers.forEach((header, index) => {
            html += '<label><input type="checkbox" data-col-index="' + index + '">' + header + '</label>';
        });
        container.innerHTML = html;
        document.getElementById('exportControls').style.display = 'flex';

    } catch (error) {
        container.innerHTML = '<p style="color:red;">Failed to load columns: ' + error.message + '</p>';
    }
}

/**
 * Gets a list of 0-based indices for ignored columns.
 */
function getCheckedColumnIndices(containerId) {
    const container = document.getElementById(containerId);
    const checkboxes = container.querySelectorAll('input[type="checkbox"]:checked');
    const indices = [];
    checkboxes.forEach(cb => {
        indices.push(parseInt(cb.getAttribute('data-col-index')));
    });
    return indices;
}


// ---------------------------------------------------------------------
// --- Matching and Results (Updated Output) ---
// ---------------------------------------------------------------------

window.runMatch = async function() {
    if (selectedSheets.length !== 2) return;
    
    document.getElementById('matchBtn').disabled = true;
    document.getElementById('results').innerHTML = '<p>Running comparison...</p>';
    document.getElementById('exportAllBtn').disabled = true;
    document.getElementById('visualizationSection').style.display = 'none';

    const payload = {
        sheet1: selectedSheets[0],
        sheet2: selectedSheets[1],
        useFuzzy: document.getElementById('fuzzyMatch').checked,
        fuzzyThreshold: parseInt(document.getElementById('fuzzyThreshold').value) || 20,
        isTargeted: false
    };

    try {
        allMatches = await fetchData('/api/match', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(payload)
        });

        const totalMatches = allMatches.reduce((sum, group) => sum + group.matches.length, 0);
        
        let html = '<h2>Comparison Results</h2>';
        html += '<p>Total Matches: <strong>' + totalMatches + '</strong> across ' + allMatches.length + ' unique column pairs.</p>';

        allMatches.forEach((group, index) => {
            // Group Header
            html += '<h3>Group ' + (index + 1) + ': ' + group.header1 + ' (Sheet 1) ↔ ' + group.header2 + ' (Sheet 2)</h3>'; 
            html += '<p>Matches Found: <strong>' + group.matches.length + '</strong></p>';
            
            // Export Group button
            html += '<button class="export-group-btn" onclick="handleExportMatchGroup(' + index + ')">Export Group Rows</button>';
            
            // Display a few example matches (Improved formatting)
            html += '<ul>';
            group.matches.slice(0, 3).forEach(match => {
                const type = match.isFuzzy ? 'Fuzzy Match' : 'Exact Match';
                html += '<li><strong>' + type + '</strong>: Row ' + match.originalRow1 + ' (Sheet 1) = "' + match.val1 + '" matched Row ' + match.originalRow2 + ' (Sheet 2) = "' + match.val2 + '"</li>';
            });
            if (group.matches.length > 3) html += '<li>... and ' + (group.matches.length - 3) + ' more matches.</li>';
            html += '</ul>';
        });
        
        document.getElementById('results').innerHTML = html;
        document.getElementById('exportAllBtn').disabled = totalMatches === 0;
        document.getElementById('visualizationSection').style.display = 'block';

    } catch (error) {
        document.getElementById('results').innerHTML = '<p style="color: red;">Matching failed: ' + error.message + '</p>';
    } finally {
        document.getElementById('matchBtn').disabled = false;
    }
}

// ---------------------------------------------------------------------
// --- Export Logic (Framework - Requires XLSX.js) ---
// ---------------------------------------------------------------------

/**
 * Prepares data for export by joining sheets and ignoring selected columns.
 * NOTE: This relies on the global XLSX object from the xlsx.js library.
 */
function prepareExportData(group, ignoreCols1, ignoreCols2) {
    if (!sheetDataCache[group.Tab1] || !sheetDataCache[group.Tab2]) {
        alert("Error: Missing sheet data cache. Please re-run the match.");
        return [];
    }

    const data1 = sheetDataCache[group.Tab1];
    const data2 = sheetDataCache[group.Tab2];
    
    // 1. Build Header
    const headers1Filtered = data1.headers.filter((_, i) => !ignoreCols1.includes(i));
    const headers2Filtered = data2.headers.filter((_, i) => !ignoreCols2.includes(i));
    const header = ['Row1', ...headers1Filtered, 'Row2', ...headers2Filtered, 'Match_Col1', 'Match_Col2', 'Match_Type'];

    // 2. Build Rows
    const combinedRows = [];
    group.matches.forEach(m => {
        // Row index in Go's Rows array is 2 less than originalRowX
        const rowData1 = data1.rows[m.originalRow1 - 2];
        const rowData2 = data2.rows[m.originalRow2 - 2];

        // Filter out ignored columns
        const rowData1Filtered = rowData1.filter((_, i) => !ignoreCols1.includes(i));
        const rowData2Filtered = rowData2.filter((_, i) => !ignoreCols2.includes(i));

        const row = [
            m.originalRow1,
            ...rowData1Filtered,
            m.originalRow2,
            ...rowData2Filtered,
            group.header1, 
            group.header2,
            m.isFuzzy ? 'Fuzzy' : 'Exact'
        ];
        combinedRows.push(row);
    });

    return [header, ...combinedRows];
}

/**
 * Exports all matches from all groups into a single Excel file.
 */
window.handleExportAllMatches = function() {
    if (allMatches.length === 0) return alert("No matches found to export.");
    if (typeof XLSX === 'undefined') return alert("Error: XLSX library (xlsx.full.min.js) is required for export.");

    const ignoreCols1 = getCheckedColumnIndices('exportIgnoreCols1Container');
    const ignoreCols2 = getCheckedColumnIndices('exportIgnoreCols2Container');
    
    // We need to re-process all groups to combine them
    let combinedData = [];
    
    allMatches.forEach(g => {
        const [header, ...rows] = prepareExportData(g, ignoreCols1, ignoreCols2);
        // Add header only once
        if (combinedData.length === 0) {
            combinedData.push(header);
        }
        combinedData.push(...rows);
    });

    if (combinedData.length <= 1) return alert("No rows selected for export after filtering.");

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(combinedData);
    XLSX.utils.book_append_sheet(wb, ws, "All Matches Combined");
    XLSX.writeFile(wb, "EDM_All_Matches_Combined.xlsx");
}

/**
 * Exports a single match group to an Excel file.
 */
window.handleExportMatchGroup = function(groupIndex) {
    if (typeof XLSX === 'undefined') return alert("Error: XLSX library (xlsx.full.min.js) is required for export.");

    const g = allMatches[groupIndex];
    const ignoreCols1 = getCheckedColumnIndices('exportIgnoreCols1Container');
    const ignoreCols2 = getCheckedColumnIndices('exportIgnoreCols2Container');
    
    const data = prepareExportData(g, ignoreCols1, ignoreCols2);
    
    if (data.length <= 1) return alert("No rows selected in this group.");
    
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(data), "Results");
    XLSX.writeFile(wb, "Match_Group_Export_" + g.Header1 + "-" + g.Header2 + ".xlsx");
}

// Initial state setup
document.addEventListener('DOMContentLoaded', () => {
    updateSelectionSlots();
});