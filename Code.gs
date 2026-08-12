function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Drive Chip Automator')
    .addItem('Convert & Refresh Smart Chips (Entire Sheet)', 'convertAllDynamicColumnsToChips')
    .addItem('🧪 Test Update on Selected Cell', 'testActiveCellChip')
    .addItem('Reset Script Progress', 'resetProgress')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function resetProgress() {
  PropertiesService.getUserProperties().deleteAllProperties();
  SpreadsheetApp.getUi().alert("Reset Complete", "Saved progress cleared.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// --- NEW TEST FUNCTION ---
function testActiveCellChip() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const cell = sheet.getActiveCell();
  const a1 = sheet.getName() + "!" + cell.getA1Notation();

  let rawUrl = null;

  try {
    let response = Sheets.Spreadsheets.get(ss.getId(), {
      ranges: [a1],
      fields: "sheets(data(rowData(values(userEnteredValue,chipRuns,hyperlink))))"
    });

    let cellData = response.sheets[0].data[0].rowData[0].values[0];

    if (cellData.chipRuns && cellData.chipRuns.length > 0) {
      rawUrl = cellData.chipRuns[0].chip.richLinkProperties.uri;
    } else if (cellData.hyperlink) {
      rawUrl = cellData.hyperlink;
    } else if (cellData.userEnteredValue && cellData.userEnteredValue.stringValue) {
      rawUrl = cellData.userEnteredValue.stringValue;
    }
  } catch (e) {
    // Cell might be empty or unreadable
  }

  const drivePattern = /https:\/\/(?:drive|docs)\.google\.com\/[^\s"']+/i;
  if (rawUrl && rawUrl.match(drivePattern)) {
    rawUrl = rawUrl.match(drivePattern)[0];
    // Clean up any lingering refresh tags from previous script versions
    rawUrl = rawUrl.replace(/([?&])refresh=\d+/g, '').replace(/[?&]$/, '');
  } else {
    SpreadsheetApp.getUi().alert("Test Failed", "No Drive URL found in cell " + cell.getA1Notation() + ".", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let rangeObj = {
    sheetId: sheet.getSheetId(),
    startRowIndex: cell.getRow() - 1,
    endRowIndex: cell.getRow(),
    startColumnIndex: cell.getColumn() - 1,
    endColumnIndex: cell.getColumn()
  };

  // STEP 1: Wipe the cell
  let wipeRequest = {
    updateCells: {
      range: rangeObj,
      rows: [{ values: [{ userEnteredValue: { stringValue: "" } }] }],
      fields: "userEnteredValue,chipRuns,hyperlink"
    }
  };

  // STEP 2: Write the new chip using the exact original URL
  let writeRequest = {
    updateCells: {
      range: rangeObj,
      rows: [{
        values: [{
          userEnteredValue: { stringValue: "@" },
          chipRuns: [{
            startIndex: 0,
            chip: {
              richLinkProperties: {
                uri: rawUrl
              }
            }
          }]
        }]
      }],
      fields: "userEnteredValue,chipRuns"
    }
  };

  try {
    Sheets.Spreadsheets.batchUpdate({ requests: [wipeRequest, writeRequest] }, ss.getId());
    SpreadsheetApp.getUi().alert("Success!", "Wiped and refreshed cell " + cell.getA1Notation() + "!\n\nURL sent to API:\n" + rawUrl, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error", "Failed to update chip in " + cell.getA1Notation() + ".\n\nError Message: " + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// --- MAIN FUNCTION ---
function convertAllDynamicColumnsToChips() {
  const startTime = new Date().getTime();
  const MAX_RUNTIME_MS = 240000;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const startRow = 2;
  const ROW_BLOCK_SIZE = 100;

  const userProperties = PropertiesService.getUserProperties();
  let savedSheetName = userProperties.getProperty('SAVED_SHEET_NAME');
  let savedRowIndex = parseInt(userProperties.getProperty('SAVED_ROW_INDEX'), 10) || startRow;
  let skipSheets = savedSheetName ? true : false;

  const targetedKeywords = ["attachment", "attachments", "exhibit"];
  let errorLog = [];
  let successCount = 0;

  for (let s = 0; s < sheets.length; s++) {
    let sheet = sheets[s];
    let sheetName = sheet.getName();

    if (skipSheets) {
      if (sheetName === savedSheetName) skipSheets = false;
      else continue;
    }

    let lastRow = sheet.getLastRow();
    let lastCol = sheet.getLastColumn();

    if (lastRow < startRow || lastCol < 1) continue;

    const headerValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let columnsToScan = [];

    for (let c = 0; c < headerValues.length; c++) {
      let headerText = String(headerValues[c]).toLowerCase().trim();
      if (targetedKeywords.some(keyword => headerText.includes(keyword))) {
        columnsToScan.push(c + 1);
      }
    }

    if (columnsToScan.length === 0) {
      savedRowIndex = startRow;
      continue;
    }

    let advancedData;
    try {
      advancedData = Sheets.Spreadsheets.get(ss.getId(), {
        ranges: [sheetName],
        fields: "sheets(data(rowData(values(userEnteredValue,chipRuns,hyperlink))))"
      });
    } catch (e) {
      continue;
    }

    let sheetRowData = [];
    if (advancedData.sheets && advancedData.sheets[0].data && advancedData.sheets[0].data[0].rowData) {
      sheetRowData = advancedData.sheets[0].data[0].rowData;
    }

    let currentRow = (sheetName === savedSheetName) ? savedRowIndex : startRow;

    while (currentRow <= lastRow) {
      if (new Date().getTime() - startTime > MAX_RUNTIME_MS) {
        userProperties.setProperty('SAVED_SHEET_NAME', sheetName);
        userProperties.setProperty('SAVED_ROW_INDEX', currentRow.toString());
        SpreadsheetApp.getUi().alert("Timeout Approaching", `Progress saved. Processed ${successCount} chips so far. Run again to continue.`, SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }

      let rowsRemaining = lastRow - currentRow + 1;
      let currentBlockSize = Math.min(ROW_BLOCK_SIZE, rowsRemaining);

      let apiRequests = [];
      let cellReferences = [];

      for (let col of columnsToScan) {
        for (let r = 0; r < currentBlockSize; r++) {
          let rowIndex = (currentRow - 1) + r;
          let colIndex = col - 1;

          let rawUrl = null;

          if (sheetRowData[rowIndex] && sheetRowData[rowIndex].values && sheetRowData[rowIndex].values[colIndex]) {
            let cellData = sheetRowData[rowIndex].values[colIndex];

            if (cellData.chipRuns && cellData.chipRuns.length > 0) {
              rawUrl = cellData.chipRuns[0].chip.richLinkProperties.uri;
            } else if (cellData.hyperlink) {
              rawUrl = cellData.hyperlink;
            } else if (cellData.userEnteredValue && cellData.userEnteredValue.stringValue) {
              rawUrl = cellData.userEnteredValue.stringValue;
            }
          }

          const drivePattern = /https:\/\/(?:drive|docs)\.google\.com\/[^\s"']+/i;
          if (rawUrl && rawUrl.match(drivePattern)) {
            let cleanUrl = rawUrl.match(drivePattern)[0];

            // Clean up any lingering refresh tags
            cleanUrl = cleanUrl.replace(/([?&])refresh=\d+/g, '').replace(/[?&]$/, '');

            let cellA1 = sheet.getRange(rowIndex + 1, col).getA1Notation();
            let rangeObj = {
              sheetId: sheet.getSheetId(),
              startRowIndex: rowIndex,
              endRowIndex: rowIndex + 1,
              startColumnIndex: colIndex,
              endColumnIndex: colIndex + 1
            };

            // STEP 1: Wipe request
            let wipeReq = {
              updateCells: {
                range: rangeObj,
                rows: [{ values: [{ userEnteredValue: { stringValue: "" } }] }],
                fields: "userEnteredValue,chipRuns,hyperlink"
              }
            };

            // STEP 2: Write request with exact original URL
            let writeReq = {
              updateCells: {
                range: rangeObj,
                rows: [{
                  values: [{
                    userEnteredValue: { stringValue: "@" },
                    chipRuns: [{
                      startIndex: 0,
                      chip: {
                        richLinkProperties: {
                          uri: cleanUrl
                        }
                      }
                    }]
                  }]
                }],
                fields: "userEnteredValue,chipRuns"
              }
            };

            apiRequests.push(wipeReq, writeReq);
            cellReferences.push({ cell: cellA1, wipe: wipeReq, write: writeReq });
          }
        }
      }

      if (apiRequests.length > 0) {
        try {
          Sheets.Spreadsheets.batchUpdate({ requests: apiRequests }, ss.getId());
          successCount += (apiRequests.length / 2);
        } catch (e) {
          // FALLBACK: Process one by one with a strict 1.5s delay to prevent Quota Exceeded errors
          for (let item of cellReferences) {
            try {
              Sheets.Spreadsheets.batchUpdate({ requests: [item.wipe, item.write] }, ss.getId());
              successCount++;
              Utilities.sleep(1500); // 1.5 seconds ensures we stay under the 60 requests/minute limit
            } catch (err) {
              // Properly extract the real error message without chopping it off
              let realError = err.message;
              if (realError.includes("error:")) {
                realError = realError.split("error:")[1].trim();
              } else {
                realError = realError.split("\n")[0];
              }
              errorLog.push(`${item.cell} (${realError})`);
            }
          }
        }
      }

      currentRow += currentBlockSize;
    }
    savedRowIndex = startRow;
  }

  userProperties.deleteAllProperties();

  if (errorLog.length > 0) {
    SpreadsheetApp.getUi().alert(
      "Finished with some errors",
      `Successfully updated ${successCount} chips.\n\nHowever, the following cells failed:\n${errorLog.join("\n")}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } else {
    SpreadsheetApp.getUi().alert("Success!", `Successfully generated and refreshed ${successCount} Smart Chips!`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// DO NOT DELETE: Forces Google Drive authorization scope.
function forceDriveScope() {
  DriveApp.getFiles();
}
