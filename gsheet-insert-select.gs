/**
 * ============================================================================
 * GOOGLE APPS SCRIPT: DYNAMIC SHEET API (NESTED RECORD & DYNAMIC HEADERS)
 * ============================================================================
 * 
 * --- HOW TO USE WITH cURL ---
 * 
 * IMPORTANT: Replace 'YOUR_DEPLOYMENT_ID' with your actual web app URL ID.
 * 
 * 1. POST: Append a nested record mapped dynamically to column headers
 * 
 * curl -X POST "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec" \
 *      -H "Content-Type: application/json" \
 *      -d '{
 *            "token": "YOUR_SECRET_TOKEN",
 *            "spreadsheet_id": "111pOmcnd1Dotmu3yLQzLOzFwxvStPBCtt46yDg8Cjkk",
 *            "sheet_name": "Sheet1",
 *            "record": {
 *              "name": "Alice Developer",
 *              "email": "alice@example.com",
 *              "notes": "Testing dynamic column mapping API"
 *            }
 *          }'
 * 
 * 2. GET: Download sheet data as CSV
 * 
 * curl -L "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?token=YOUR_SECRET_TOKEN&spreadsheet_id=111pOmcnd1Dotmu3yLQzLOzFwxvStPBCtt46yDg8Cjkk&sheet_name=Sheet1" \
 *      -o sheet_export.csv
 * 
 * ============================================================================
 */

/**
 * GET Request Handler
 * Fetches data from a dynamically specified Google Sheet and returns it as a CSV file.
 */
function doGet(e) {
  try {
    var spreadsheetId = e.parameter.spreadsheet_id || null;
    var sheetName = e.parameter.sheet_name || "Sheet1";
    var token = e.parameter.token || null;

    var SECRET_TOKEN = "YOUR_SECRET_TOKEN"; 
    if (SECRET_TOKEN !== "YOUR_SECRET_TOKEN" && token !== SECRET_TOKEN) {
      return createJsonResponse({ status: "error", message: "Unauthorized. Invalid token." }, 401);
    }

    var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return createJsonResponse({ status: "error", message: "Could not resolve target spreadsheet." }, 404);

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return createJsonResponse({ status: "error", message: "Sheet name '" + sheetName + "' not found." }, 404);

    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var csvContent = "";

    if (lastRow > 0 && lastColumn > 0) {
      var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
      var csvRows = [];
      for (var i = 0; i < data.length; i++) {
        var row = data[i].map(function(value) {
          var stringValue = value instanceof Date ? value.toISOString() : String(value);
          if (stringValue.indexOf(',') !== -1 || stringValue.indexOf('"') !== -1 || stringValue.indexOf('\n') !== -1) {
            return '"' + stringValue.replace(/"/g, '""') + '"';
          }
          return stringValue;
        });
        csvRows.push(row.join(","));
      }
      csvContent = csvRows.join("\n");
    }

    var filename = sheetName + "_" + new Date().toISOString().substring(0, 10) + ".csv";
    
    // Fixed: Removed the invalid addMetaAllowedSetting method
    return ContentService.createTextOutput(csvContent)
      .setMimeType(ContentService.MimeType.CSV)
      .downloadAsFile(filename);

  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() }, 500);
  }
}

/**
 * POST Request Handler
 * Appends a new record to a dynamically specified Google Sheet, mapped to headers.
 */
function doPost(e) {
  var headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "No post data received" }, 400);
    }
    
    var payload = JSON.parse(e.postData.contents);

    var SECRET_TOKEN = "YOUR_SECRET_TOKEN"; 
    if (SECRET_TOKEN !== "YOUR_SECRET_TOKEN" && payload.token !== SECRET_TOKEN) {
      return createJsonResponse({ status: "error", message: "Unauthorized. Invalid token." }, 401);
    }

    // var spreadsheetId = payload.spreadsheet_id || null; 
    var spreadsheetId = null; 
    var sheetName = payload.sheet_name || "Sheet1"; 

    var ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return createJsonResponse({ status: "error", message: "Could not resolve target spreadsheet." }, 404);

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return createJsonResponse({ status: "error", message: "Sheet name '" + sheetName + "' not found." }, 404);

    var lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return createJsonResponse({ status: "error", message: "Target sheet is empty. Please define column headers in Row 1." }, 400);
    }

    // 1. Read the schema from Row 1
    var sheetHeaders = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    
    // 2. Initialize an empty array matching the exact width of the sheet
    var rowData = new Array(sheetHeaders.length).fill("");
    var record = payload.record || {};

    // 3. ENFORCE TIMESTAMP: First column (index 0) is ALWAYS the timestamp
    rowData[0] = new Date();

    // 4. Map JSON keys to the remaining Sheet Columns (starting at index 1)
    for (var i = 1; i < sheetHeaders.length; i++) {
      var headerName = String(sheetHeaders[i]).trim(); 
      
      // If the dictionary contains a key matching the column header, slot it in
      if (record.hasOwnProperty(headerName)) {
        rowData[i] = record[headerName];
      }
    }

    // 5. Append the fully constructed array
    sheet.appendRow(rowData);

    return createJsonResponse({
      status: "success",
      message: "Row successfully mapped and recorded!",
      insertedRow: rowData
    }, 200);

  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() }, 500);
  }
}

/**
 * Helper to generate JSON responses
 */
function createJsonResponse(data, statusCode) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}