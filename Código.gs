/**
 * @OnlyCurrentDoc
 */

function onOpen() {

  SpreadsheetApp.getUi().createMenu('breakApart')
  .addItem('Break apart selected ranges', 'breakApart')
  .addToUi();
  
}

/**
 * Menu command
 */
function breakApart() {
  
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  breakApartAux(ss.getActiveSheet(), ss.getActiveSheet().getRange(ss.getRange('B1').getValue()));
  
}

/** 
* Breaks apart merged cells in range
* @param {sheet} sheet Sheet of range
* @param {range} Range Cell or cells that intersect with whole merged range
*/
function breakApartAux(sheet, rangeToSplit) {
  
  // Currente spreadsheet & sheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let activeSSId = ss.getId();
  let activeSId = sheet.getSheetId();
  
  // Get sheet's merges using advanced sheets service
  let merges = Sheets.Spreadsheets.get(activeSSId).sheets.find(s => s.properties.sheetId == activeSId).merges;
  
  // Logger.log(merges);
   
  // Cells to merge R/C
  let rowS = rangeToSplit.getRow();
  let rowE = rowS + rangeToSplit.getNumRows() - 1;
  let colS = rangeToSplit.getColumn();
  let colE = colS + rangeToSplit.getNumColumns() - 1;
  
  // Find overlapping merged range
  // Advanced service ranges start in 0 and are right-open [..)
  let merge = merges.find(m => {
                          let mRowS = m.startRowIndex + 1;
                          let mRowE = m.endRowIndex;
                          let mColS = m.startColumnIndex + 1;
                          let mColE = m.endColumnIndex;
      
                          // Check for overlapping
                          return ((rowS >= mRowS && rowS <= mRowE)  || 
                                  (rowE <= mRowE && rowE >= mRowS)  ||
                                  (rowS < mRowS && rowE > mRowE)) &&
                                 ((colS >= mColS && colS <= mColE)  ||
                                  (colE <= mColE && colE >= mColS)  ||
                                  (colS < mColS && colE > mColE));
  })
  
  // Overlapping range?
  if (merge != undefined) {
  // Break apart whole range
  ss.getActiveSheet().getRange(merge.startRowIndex + 1,
                               merge.startColumnIndex + 1,
                               merge.endRowIndex - merge.startRowIndex,
                               merge.endColumnIndex - merge.startColumnIndex).breakApart();
  } else SpreadsheetApp.getUi().alert('No merged cells found in specified range.');
 
} 