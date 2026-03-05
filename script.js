// ════════════════════════════════════════════════════════════
//  SPEC DECODER DATA
// ════════════════════════════════════════════════════════════
const specData = {
    product: {
        'C': 'Cover',
        'W': 'Cover with Flaps',
        'J': 'Jacket',
        'T': 'Tip-In',
        'F': 'Cover For Case'
    },

    colours: {
        '0': 'No Colour Print',
        '1': '1 Spot Colour',
        '2': '2 Spot Colours',
        '3': '3 Spot Colours',
        '4': '4 Process Colours',
        '5': '4 Process Colours + 1 Spot Colour',
        '6': '4 Process Colours + 2 Spot Colours',
        '7': '4 Spot Colours',
        '8': '4 Process Colours + 3 Spot Colours',
        '9': '4 Process Colours + 4 Spot Colours'
    },

    finish: {
        '0': 'No Finish',
        '1': 'Gloss Varnish In Line',
        '2': 'Gloss Varnish In Line + Matt Varnish Offline',
        '3': 'Gloss Varnish Off Line',
        '4': 'Matt Varnish Off Line',
        '5': 'Gloss Laminate (Standard)',
        '6': 'Matt Laminate (Standard)',
        '7': 'Matt Laminate (Standard) / Gloss Spot Varnish',
        '8': 'Silk Laminate',
        '9': 'Anti-Scuff Laminate',
        'A': 'Gloss Laminate (Standard) / Matt Spot UV',
        'B': 'Silk Laminate / Matt Spot UV',
        'C': 'Anti-Scuff Laminate / Gloss Spot UV',
        'D': 'Gloss Varnish Off Line + Matt Spot UV',
        'E': 'Matt Varnish In Line + Gloss Spot UV',
        'F': 'Matt Varnish In Line',
        'G': 'Matt Varnish Off Line + Gloss Spot UV',
        'H': 'Outwork Lamination',
        'J': 'Outwork Lamination / Gloss Spot UV',
        'K': 'Outwork Lamination / Matt Spot UV',
        'L': 'Gloss Spot UV',
        'M': 'Matt Spot UV',
        'N': 'Gloss Varnish In Line + Matt Spot UV',
        'Q': 'Soft Matt Lam',
        'R': 'Soft Matt Lam / Gloss Spot Varnish',
        'V': 'Recycled Matt Laminate',
        'W': 'Recycled Matt Laminate / Gloss Spot Varnish',
        'Y': 'Recycled Gloss Laminate',
        'Z': 'Recycled Gloss Laminate / Matt Spot UV'
    },

    texture: {
        'P': 'Plain',
        'G': 'Grained'
    },

    weight: {
        '1': '220 gsm',
        '2': '220 gsm',
        '3': '260 gsm',
        '4': '150 gsm',
        '5': '135 gsm',
        '6': '130 gsm',
        '7': '220 gsm'
    },

    specialInks: {
        'F':  'Fluorescent',
        'S':  'Spot Colour',
        'M':  'Non-Conventional Metallic',
        'K':  'Conventional Metallic (used with M)',
        'B':  'Blocked (after print, before laminate)',
        'E':  'Embossed',
        'D':  'Debossed',
        'C':  'Die-Cutting',
        'P':  'Print Over Foil',
        'L':  'Block Over Laminate',
        'U':  'Uncoated Printing',
        'PB': 'Print Black Over Foil',
        'BE': 'Block & Emboss (same pass)',
        'DE': 'Deboss & Emboss (same pass)',
        'BD': 'Block & Deboss (same pass)',
        'S1': 'Other Spot UV',
        'S2': 'Pile Spot UV',
        'S3': 'Glitter Spot UV',
        'V1': 'Glow Varnish',
        'H1': 'Holographic Lam'
    }
};

// ════════════════════════════════════════════════════════════
//  DECODE HELPERS
// ════════════════════════════════════════════════════════════

/**
 * Decode a spec code into a plain-English summary string.
 * Returns null if the code is too short or the first char is unknown.
 */
function decodeSpecCode(code) {
    code = (code || '').toUpperCase().trim();
    if (code.length < 6) return null;

    const parts = [];

    const product = specData.product[code[0]];
    if (product) parts.push(`Product: ${product}`);

    const outside = specData.colours[code[1]];
    if (outside) parts.push(`Outside: ${outside}`);

    const inside = specData.colours[code[2]];
    if (inside) parts.push(`Inside: ${inside}`);

    const finish = specData.finish[code[3]];
    if (finish) parts.push(`Finish: ${finish}`);

    if (code.length > 6) {
        const special = [];
        let i = 0;
        const section = code.substring(6);
        while (i < section.length) {
            if (i < section.length - 1 && specData.specialInks[section.substring(i, i + 2)]) {
                special.push(specData.specialInks[section.substring(i, i + 2)]);
                i += 2;
            } else if (specData.specialInks[section[i]]) {
                special.push(specData.specialInks[section[i]]);
                i++;
            } else {
                i++;
            }
        }
        if (special.length) parts.push(`Special: ${special.join(', ')}`);
    }

    return parts.length ? parts.join(' | ') : null;
}

// ════════════════════════════════════════════════════════════
//  INIT
// ════════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('excelFile').addEventListener('change', e => {
        if (e.target.files[0]) processFile(e.target.files[0]);
    });

    document.querySelectorAll('input[name="splitMode"]').forEach(el => {
        el.addEventListener('change', renderTable);
    });
});


// ════════════════════════════════════════════════════════════
//  ORDER SPLITTER
// ════════════════════════════════════════════════════════════
let workbookData = [];   // parsed row objects (for preview table)
let sourceWb     = null; // raw SheetJS workbook (for format-preserving exports)
let fileName     = '';

function processFile(file) {
    fileName = file.name;
    showStatus('Processing file...', 'info');

    const reader = new FileReader();
    reader.onload = e => {
        try {
            // Keep the raw workbook so we can copy cells with original formatting
            sourceWb = XLSX.read(new Uint8Array(e.target.result), {
                type: 'array',
                cellStyles: true,
                cellDates:  false,
                cellNF:     true
            });

            const ws  = sourceWb.Sheets[sourceWb.SheetNames[0]];
            const raw = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });

            if (!raw.length) { showStatus('No data found in file', 'danger'); return; }

            const cols = Object.keys(raw[0]);
            if (!cols.includes('Cover Spec') || !cols.includes('Paper')) {
                showStatus('Could not find "Cover Spec" or "Paper" columns', 'danger');
                return;
            }

            workbookData = raw;

            document.getElementById('modeRow').style.display = '';
            renderTable();
            showStatus(`File loaded successfully! ${raw.length} rows found.`, 'success');
            enableSplitterButtons(true);

        } catch (err) {
            showStatus('Error processing Excel file. Please check the file format.', 'danger');
            enableSplitterButtons(false);
        }
    };
    reader.onerror = () => showStatus('Error reading file. Please try again.', 'danger');
    reader.readAsArrayBuffer(file);
}

function clearAll() {
    workbookData = [];
    sourceWb     = null;
    fileName     = '';

    document.getElementById('excelFile').value        = '';
    document.getElementById('modeRow').style.display  = 'none';
    document.getElementById('comboInfo').textContent  = '';
    document.getElementById('previewBody').innerHTML  =
        '<tr><td colspan="7" class="text-center text-muted">No data loaded</td></tr>';
    document.getElementById('modeBoth').checked       = true;

    enableSplitterButtons(false);
    showStatus('All data cleared.', 'info');
    setTimeout(() => { document.getElementById('status').style.display = 'none'; }, 3000);
}

function getMode() {
    return document.querySelector('input[name="splitMode"]:checked').value;
}

// Numeric helper — parse a cell value to float, treating blanks/non-numeric as -Infinity
// so that rows with missing values always sort to the bottom.
function numVal(v) {
    const n = parseFloat(v);
    return isNaN(n) ? -Infinity : n;
}

// Sort an array of row indices (into workbookData) by:
//   Trim Width desc → Trim Height desc → Extent desc → Quantity desc
function sortIndices(indices) {
    return [...indices].sort((a, b) => {
        const ra = workbookData[a];
        const rb = workbookData[b];

        const wDiff = numVal(rb['Trim Width'])  - numVal(ra['Trim Width']);
        if (wDiff !== 0) return wDiff;

        const hDiff = numVal(rb['Trim Height']) - numVal(ra['Trim Height']);
        if (hDiff !== 0) return hDiff;

        const eDiff = numVal(rb['Extent'])      - numVal(ra['Extent']);
        if (eDiff !== 0) return eDiff;

        return numVal(rb['Quantity']) - numVal(ra['Quantity']);
    });
}

// Returns true if the Cover Spec code for this row decodes to "No Finish" (position 3 = '0')
function isNoFinish(row) {
    const code = (row['Cover Spec'] || '').toUpperCase().trim();
    return code.length >= 4 && code[3] === '0';
}

// Returns a map of groupKey -> array of 0-based row indices (excluding header row 0)
function getGroupIndices(mode) {
    const groups = {};
    workbookData.forEach((row, idx) => {
        let key;
        if (mode === 'both')       key = `${row['Cover Spec']} - ${row['Paper']}`;
        else if (mode === 'cover') key = row['Cover Spec'] || 'Unknown';
        else                       key = row['Paper']      || 'Unknown';

        // When the finish is "No Finish", further split by Bleed (Yes / No)
        if (isNoFinish(row)) {
            const bleed = (String(row['Bleeds'] || '').trim().toLowerCase() === 'yes') ? 'Yes' : 'No';
            key += ` - Bleed ${bleed}`;
        }

        if (!groups[key]) groups[key] = [];
        groups[key].push(idx); // idx into workbookData = data row index (0-based)
    });

    // Sort each group's rows: Trim Width → Trim Height → Extent → Quantity (all desc)
    Object.keys(groups).forEach(key => {
        groups[key] = sortIndices(groups[key]);
    });

    return groups;
}

function safeSheetName(key) {
    return key.replace(/[\/\\?*\[\]]/g, '-').substring(0, 31);
}

function renderTable() {
    const mode   = getMode();
    const groups = getGroupIndices(mode);
    const keys   = Object.keys(groups).sort();
    const tbody  = document.getElementById('previewBody');
    tbody.innerHTML = '';

    document.getElementById('comboInfo').textContent =
        `${keys.length} sheet${keys.length !== 1 ? 's' : ''} will be created from ${workbookData.length} rows`;

    keys.forEach(key => {
        const indices   = groups[key];
        const firstRow  = workbookData[indices[0]];
        const sheetName = safeSheetName(key);
        const coverSpec = firstRow['Cover Spec'] || '';
        const decoded   = mode !== 'paper' ? (decodeSpecCode(coverSpec) || coverSpec || '—') : '—';

        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>
                <button class="btn btn-primary btn-sm" onclick="downloadSheet('${encodeURIComponent(key)}')">
                    <i class="fas fa-download"></i>
                </button>
            </td>
            <td>${sheetName}</td>
            <td><span class="badge bg-secondary">${indices.length}</span></td>
            <td class="cover-spec-cell">${decoded}</td>
            <td>${mode !== 'cover' ? (firstRow['Paper']  || '—') : '—'}</td>
            <td>${firstRow['GSM']    || '—'}</td>
            <td>${firstRow['Micron'] || '—'}</td>
        `;
        tbody.appendChild(tr);
    });
}

// ── Core: build a new worksheet from the source sheet, copying only the
//   header row plus the specified data row indices, preserving all cell
//   objects (values, styles, number formats) and column widths exactly.
function buildFilteredSheet(rowIndices) {
    const srcWs  = sourceWb.Sheets[sourceWb.SheetNames[0]];
    const srcRef = XLSX.utils.decode_range(srcWs['!ref']);
    const numCols = srcRef.e.c + 1;

    const newWs = {};

    // Copy header row (source row 0) → destination row 0
    for (let c = 0; c < numCols; c++) {
        const srcAddr = XLSX.utils.encode_cell({ r: 0, c });
        if (srcWs[srcAddr]) newWs[XLSX.utils.encode_cell({ r: 0, c })] = { ...srcWs[srcAddr] };
    }

    // Copy each selected data row; source data rows are 1-based (row 0 = header)
    rowIndices.forEach((dataIdx, destIdx) => {
        const srcRow = dataIdx + 1; // +1 because row 0 is the header
        for (let c = 0; c < numCols; c++) {
            const srcAddr  = XLSX.utils.encode_cell({ r: srcRow, c });
            const destAddr = XLSX.utils.encode_cell({ r: destIdx + 1, c });
            if (srcWs[srcAddr]) {
                newWs[destAddr] = { ...srcWs[srcAddr] };
            }
        }
    });

    // Set the sheet range
    newWs['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: rowIndices.length, c: srcRef.e.c }
    });

    // Copy column widths exactly (sliced to only the columns present)
    if (srcWs['!cols']) newWs['!cols'] = srcWs['!cols'].slice(0, numCols).map(c => c ? { ...c } : {});

    // Copy row heights only for rows in the new sheet to avoid phantom blank rows
    if (srcWs['!rows']) {
        const newRows = [];
        newRows[0] = srcWs['!rows'][0] ? { ...srcWs['!rows'][0] } : undefined;
        rowIndices.forEach((srcDataIdx, destIdx) => {
            const srcRowIdx = srcDataIdx + 1;
            newRows[destIdx + 1] = srcWs['!rows'][srcRowIdx] ? { ...srcWs['!rows'][srcRowIdx] } : undefined;
        });
        newWs['!rows'] = newRows;
    }

    return newWs;
}

function downloadSheet(encodedKey) {
    const key     = decodeURIComponent(encodedKey);
    const groups  = getGroupIndices(getMode());
    const indices = groups[key];
    if (!indices) return;

    const wb        = XLSX.utils.book_new();
    const ws        = buildFilteredSheet(indices);
    const sheetName = safeSheetName(key);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    XLSX.writeFile(wb, `${fileName.replace(/\.[^.]+$/, '')}_${sheetName}.xlsx`);
    showStatus(`Downloaded: ${sheetName}.xlsx`, 'success');
}

function downloadAll() {
    const groups   = getGroupIndices(getMode());
    const baseName = fileName.replace(/\.[^.]+$/, '');

    Object.entries(groups).forEach(([key, indices], i) => {
        setTimeout(() => {
            const wb        = XLSX.utils.book_new();
            const ws        = buildFilteredSheet(indices);
            const sheetName = safeSheetName(key);
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            XLSX.writeFile(wb, `${baseName}_${sheetName}.xlsx`);
        }, i * 300);
    });

    showStatus(`Downloading ${Object.keys(groups).length} files...`, 'success');
}

function downloadCombined() {
    const groups = getGroupIndices(getMode());
    const wb     = XLSX.utils.book_new();

    // Summary sheet — plain data, no special formatting needed
    const summaryRows = [['Sheet', 'Orders', 'Cover Spec', 'Decoded', 'Paper', 'GSM', 'Micron']];
    Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0])).forEach(([key, indices]) => {
        const row = workbookData[indices[0]];
        summaryRows.push([
            safeSheetName(key),
            indices.length,
            row['Cover Spec'] || '',
            decodeSpecCode(row['Cover Spec']) || '',
            row['Paper']  || '',
            row['GSM']    || '',
            row['Micron'] || ''
        ]);
    });
    const summaryWs = XLSX.utils.aoa_to_sheet(summaryRows);
    XLSX.utils.book_append_sheet(wb, summaryWs, 'Summary');

    // One format-preserving sheet per group
    Object.entries(groups).sort((a, b) => a[0].localeCompare(b[0])).forEach(([key, indices]) => {
        const ws = buildFilteredSheet(indices);
        XLSX.utils.book_append_sheet(wb, ws, safeSheetName(key));
    });

    XLSX.writeFile(wb, `${fileName.replace(/\.[^.]+$/, '')}_Split.xlsx`);
    showStatus(`Combined workbook downloaded (${Object.keys(groups).length} sheets + Summary)`, 'success');
}

function enableSplitterButtons(enabled) {
    document.getElementById('clearBtn').disabled            = !enabled;
    document.getElementById('downloadAllBtn').disabled      = !enabled;
    document.getElementById('downloadCombinedBtn').disabled = !enabled;
}

function showStatus(message, type) {
    const el = document.getElementById('status');
    el.className     = `alert alert-${type}`;
    el.textContent   = message;
    el.style.display = 'block';
    if (type === 'success') setTimeout(() => { el.style.display = 'none'; }, 3000);
}