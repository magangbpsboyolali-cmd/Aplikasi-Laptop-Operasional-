// ==========================================
// GOOGLE APPS SCRIPT - INVENTRACK BPS BOYOLALI
// Copy seluruh kode ini ke Google Apps Script Editor
// ==========================================

// GANTI DENGAN ID SPREADSHEET ANDA
const SPREADSHEET_ID = '12QxeGdy0sVTR2lqh0zxerzFM6HDXP_nuR-a4MOURj88';
var MAX_BORROW_DAYS = 5;

function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
    try {
        var params = e.parameter || {};

        // Handle JSON body for POST requests
        if (e.postData && e.postData.contents) {
            try {
                var body = JSON.parse(e.postData.contents);
                // Merge body into params
                for (var key in body) {
                    params[key] = body[key];
                }
            } catch (jsonErr) {
                // If parsing fails, ignore custom body
            }
        }

        var action = params.action;
        var sheetName = params.sheet;
        var result = { success: true };

        var allowedSheets = ['Data Laptop', 'DATA PEGAWAI', 'Data Peminjaman', 'Data Pengembalian', 'Kode Akses', 'Keperluan'];
        if (sheetName && allowedSheets.indexOf(sheetName) === -1) {
            return jsonResponse({ success: false, error: 'Invalid sheet name' });
        }

        if (action === 'getAllData') {
            result = getAllSheets();
        } else if (action === 'getSheet' && sheetName) {
            result = getSheetData(sheetName);
        } else if (action === 'appendRow' && sheetName && params.row) {
            var row = typeof params.row === 'string' ? JSON.parse(params.row) : params.row;
            result = appendData(sheetName, row);
        } else if (action === 'updateRow' && sheetName && params.row && params.matchCol && params.matchVal) {
            var row = typeof params.row === 'string' ? JSON.parse(params.row) : params.row;
            result = updateRow(sheetName, params.matchCol, params.matchVal, row);
        } else if (action === 'clearAndInsert' && sheetName && params.rows) {
            var rows = typeof params.rows === 'string' ? JSON.parse(params.rows) : params.rows;
            result = saveData(sheetName, rows);
        } else if (action === 'processReturn') {
            result = processReturnTransaction(params);
        } else {
            result.message = 'InvenTrack API - BPS Kab. Boyolali';
            result.actions = ['getAllData', 'getSheet', 'appendRow', 'updateRow', 'clearAndInsert', 'processReturn'];
        }

        return jsonResponse(result);
    } catch (error) {
        return jsonResponse({ success: false, error: error.toString() });
    }
}

// ==========================================
// PROCESS RETURN TRANSACTION (Atomic Operation)
// ==========================================
function processReturnTransaction(params) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Append to Data Pengembalian
    var returnRow = typeof params.returnRow === 'string' ? JSON.parse(params.returnRow) : params.returnRow;
    var resAppend = appendData('Data Pengembalian', returnRow);
    if (!resAppend.success) return resAppend;

    // 2. Update Data Peminjaman (STATUS=Selesai, TGL_REALISASI...)
    var updatePinjam = updateRow('Data Peminjaman', 'ID', params.peminjamanId, {
        STATUS: 'Selesai',
        TGL_REALISASI_PENGEMBALIAN: params.tglRealisasi,
        KONDISI_PENGEMBALIAN: params.kondisi
    });

    // 3. Update Data Laptop (STATUS)
    var updateLaptop = updateRow('Data Laptop', 'ID', params.laptopId, {
        STATUS: params.newLaptopStatus
    });

    return {
        success: true,
        message: 'Transaction completed',
        details: {
            append: resAppend.success,
            updatePinjam: updatePinjam.success,
            updateLaptop: updateLaptop.success
        }
    };
}

function jsonResponse(obj) {
    return ContentService
        .createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// GET ALL DATA
// ==========================================
function getAllSheets() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheets = ['Data Laptop', 'DATA PEGAWAI', 'Data Peminjaman', 'Data Pengembalian', 'Kode Akses', 'Keperluan'];
    var result = {
        success: true,
        timestamp: new Date().toISOString(),
        data: {}
    };

    sheets.forEach(function (name) {
        var sheet = ss.getSheetByName(name);
        var keyName = name.toLowerCase().replace(/\s+/g, '_');

        var lastRow = sheet ? sheet.getLastRow() : 0;
        var lastCol = sheet ? sheet.getLastColumn() : 0;

        if (sheet && lastRow > 1 && lastCol > 0) {
            // Get exact range to avoid over-fetching phantom rows
            var data = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
            var headers = data[0];
            var rows = [];

            for (var i = 1; i < data.length; i++) {
                var row = data[i];
                // Fast empty check
                if (row.join('').length === 0) continue;

                var obj = {};
                for (var j = 0; j < headers.length; j++) {
                    obj[headers[j]] = row[j];
                }
                rows.push(obj);
            }
            result.data[keyName] = rows;
        } else {
            result.data[keyName] = [];
        }
    });

    return result;
}

// ==========================================
// GET SINGLE SHEET
// ==========================================
function getSheetData(sheetName) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) {
        return { success: true, data: [], count: 0 };
    }

    // Optimization: Use getDisplayValues to avoid date formatting loop overhead
    var data = sheet.getDataRange().getDisplayValues();
    var headers = data[0];
    var rows = [];

    for (var i = 1; i < data.length; i++) {
        var obj = {};
        for (var j = 0; j < headers.length; j++) {
            obj[headers[j]] = data[i][j];
        }
        rows.push(obj);
    }
    return { success: true, data: rows, count: rows.length };
}

// ==========================================
// APPEND ROW
// ==========================================
function appendData(sheetName, rowObj) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        // Auto-create sheet if not found
        sheet = ss.insertSheet(sheetName);
    }

    var lastCol = sheet.getLastColumn();
    var headers;

    if (lastCol === 0) {
        // Sheet has no headers — create them from rowObj keys
        headers = Object.keys(rowObj);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        sheet.getRange(1, 1, 1, headers.length)
            .setBackground('#1e3a5f')
            .setFontColor('#FFFFFF')
            .setFontWeight('bold');
        sheet.setFrozenRows(1);
    } else {
        headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        var rowKeys = Object.keys(rowObj);
        var missingHeaders = rowKeys.filter(function (key) { return headers.indexOf(key) === -1; });
        if (missingHeaders.length > 0) {
            sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
            headers = headers.concat(missingHeaders);
            lastCol = headers.length;
        }
    }

    var rowData = headers.map(function (h) { return rowObj[h] !== undefined ? rowObj[h] : ''; });

    if (sheetName === 'Data Peminjaman') {
        var validation = validateBorrowDuration(rowObj);
        if (!validation.success) return validation;
    }

    // Insert latest records at the top for history sheets
    if (sheetName === 'Data Peminjaman' || sheetName === 'Data Pengembalian') {
        sheet.insertRowsAfter(1, 1);
        var targetRange = sheet.getRange(2, 1, 1, rowData.length);
        targetRange.setValues([rowData]);
        targetRange.clearFormat();
    } else {
        sheet.appendRow(rowData);
    }

    return { success: true, message: 'Row appended to ' + sheetName };
}

function validateBorrowDuration(rowObj) {
    var startDate = parseDateOnly(rowObj.TGL_PINJAM);
    var endDate = parseDateOnly(rowObj.TGL_KEMBALI_RENCANA);

    if (!startDate || !endDate) {
        return { success: false, error: 'Tanggal pinjam/pengembalian tidak valid.' };
    }

    var diffDays = Math.round((endDate.getTime() - startDate.getTime()) / (24 * 60 * 60 * 1000));
    if (diffDays < 0) {
        return { success: false, error: 'Tanggal kembali tidak boleh sebelum tanggal pinjam.' };
    }

    if (diffDays >= MAX_BORROW_DAYS) {
        return {
            success: false,
            error: 'Maksimal peminjaman per user adalah 5 hari dari tanggal pinjam.'
        };
    }

    return { success: true };
}

function parseDateOnly(value) {
    if (value === null || value === undefined) return null;
    var dateText = String(value).trim();
    if (!dateText) return null;

    var slashMatch = dateText.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (slashMatch) {
        var day = ('0' + slashMatch[1]).slice(-2);
        var month = ('0' + slashMatch[2]).slice(-2);
        var year = slashMatch[3];
        var parsedSlash = new Date(year + '-' + month + '-' + day + 'T00:00:00');
        if (!isNaN(parsedSlash.getTime())) return parsedSlash;
    }

    var onlyDate = dateText.split('T')[0];
    var parsed = new Date(onlyDate + 'T00:00:00');
    if (isNaN(parsed.getTime())) return null;
    return parsed;
}

// ==========================================
// UPDATE ROW (match by column value)
// ==========================================
function updateRow(sheetName, matchCol, matchVal, rowObj) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'Sheet not found' };

    var lastCol = sheet.getLastColumn();
    if (lastCol === 0) return { success: false, error: 'Empty sheet' };

    // 1. Find the column index
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colIndex = headers.indexOf(matchCol);
    if (colIndex === -1) return { success: false, error: 'Column not found: ' + matchCol };

    // 2. Use TextFinder for fast lookup
    // Search only in the specific column to avoid false positives
    var columnRange = sheet.getRange(2, colIndex + 1, sheet.getLastRow() - 1, 1);
    var finder = columnRange.createTextFinder(matchVal).matchEntireCell(true);
    var result = finder.findNext();

    if (!result) return { success: false, error: 'Row not found with ' + matchCol + '=' + matchVal };

    var rowIndex = result.getRow();

    // 3. Update fields
    for (var key in rowObj) {
        var hIdx = headers.indexOf(key);
        if (hIdx !== -1) {
            sheet.getRange(rowIndex, hIdx + 1).setValue(rowObj[key]);
        }
    }

    return { success: true, message: 'Row updated' };
}

// ==========================================
// CLEAR AND INSERT (full replace)
// ==========================================
function saveData(sheetName, rows) {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { sheet = ss.insertSheet(sheetName); }

    if (!rows || rows.length === 0) {
        var lastRow = sheet.getLastRow();
        if (lastRow > 1) {
            sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
        }
        return { success: true, count: 0 };
    }

    var headers = Object.keys(rows[0]);
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#1e3a5f')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    sheet.setFrozenRows(1);

    var data = rows.map(function (row) {
        return headers.map(function (h) { return row[h] !== undefined ? row[h] : ''; });
    });
    if (data.length > 0) {
        sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    }
    return { success: true, count: rows.length };
}

// ==========================================
// SETUP SHEETS (Jalankan sekali!)
// ==========================================
function setupSheets() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    // 1. Data Laptop
    createSheet(ss, 'Data Laptop', ['ID', 'MERK', 'TYPE', 'NUP', 'STATUS']);

    // 2. DATA PEGAWAI
    createSheet(ss, 'DATA PEGAWAI', ['NO', 'NIP', 'NAMA', 'TIM DIVISI']);

    // 3. Data Peminjaman
    createSheet(ss, 'Data Peminjaman', [
        'ID', 'LAPTOP_ID', 'NAMA_PEMINJAM', 'NIP', 'DIVISI',
        'KEPERLUAN', 'DESKRIPSI_KEPERLUAN',
        'TGL_PINJAM', 'TGL_KEMBALI_RENCANA', 'STATUS',
        'Timestamp'
    ]);

    // 4. Data Pengembalian
    createSheet(ss, 'Data Pengembalian', [
        'ID', 'PEMINJAMAN_ID', 'LAPTOP_ID', 'NAMA_PEMINJAM',
        'TGL_PINJAM', 'TGL_KEMBALI_RENCANA',
        'TGL_REALISASI_PENGEMBALIAN', 'KONDISI_PENGEMBALIAN', 'CATATAN_PENGEMBALIAN', 'STATUS',
        'Timestamp'
    ]);

    // 5. Kode Akses untuk login
    createSheet(ss, 'Kode Akses', ['ID', 'KODE']);

    // 6. Keperluan
    createSheet(ss, 'Keperluan', ['KEPERLUAN']);

    return 'Setup complete!';
}

function createSheet(ss, name, headers) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }

    var lastCol = sheet.getLastColumn();

    if (sheet.getLastRow() === 0 || lastCol === 0) {
        // Sheet is empty, set all headers
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } else {
        // Check for missing headers and append them
        var currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
        var newHeaders = [];
        var addedCount = 0;

        for (var i = 0; i < headers.length; i++) {
            if (currentHeaders.indexOf(headers[i]) === -1) {
                newHeaders.push(headers[i]);
                addedCount++;
            }
        }

        if (addedCount > 0) {
            sheet.getRange(1, lastCol + 1, 1, addedCount).setValues([newHeaders]);
        }
    }

    // Style header row
    var finalLastCol = sheet.getLastColumn();
    var headerRange = sheet.getRange(1, 1, 1, finalLastCol);
    headerRange.setBackground('#1e3a5f')
        .setFontColor('#FFFFFF')
        .setFontWeight('bold');
    sheet.setFrozenRows(1);
}
