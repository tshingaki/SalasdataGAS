function autoFormatReport() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = spreadsheet.getSheetByName('変換(受注)');
  const targetSheet = spreadsheet.getSheetByName('BP1_受注');
  
  const salesPeople = {
    '松永': [14, 33],
    '永倉': [35, 54],
    '井野口': [56, 75],
    '瀧本': [77, 96],
    '森田': [98, 117]
  };
  
  // Get the data range from the source sheet
  const dataRange = sourceSheet.getDataRange();
  const data = dataRange.getValues();
  
  // Get the month row from the target sheet (assuming it's row 2)
  const monthRow = targetSheet.getRange(2, 1, 1, targetSheet.getLastColumn()).getValues()[0];
  
  // Loop through the data and process each row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const jobNo = row[0];  // Job No
    const dealName = row[1];  // 取引名
    const salesPerson = row[2];  // 取引担当者
    const amount = row[3];  // 金額
    const receivedDate = new Date(row[4]);  // 受注日

    const month = receivedDate.getMonth() + 1; // Get month from 1 (January) to 12 (December)
    const year = receivedDate.getFullYear();
    const currentYear = new Date().getFullYear();

    if (year === currentYear && salesPeople[salesPerson]) {
      // Find the corresponding column for the month
      let targetColumn = -1;
      for (let j = 1; j < monthRow.length; j++) {
        if (monthRow[j] == month + "月") {
          targetColumn = j + 1;
          break;
        }
      }
      
      if (targetColumn !== -1) {
        const [startRow, endRow] = salesPeople[salesPerson];

        // Find the first empty row in the salesPerson's range
        let targetRow = -1;
        for (let r = startRow; r <= endRow; r++) {
          if (!targetSheet.getRange(r, targetColumn).getValue()) {
            targetRow = r;
            break;
          }
        }
        
        if (targetRow !== -1) {
          // Fill in the data
          targetSheet.getRange(targetRow, targetColumn).setValue(jobNo);
          targetSheet.getRange(targetRow, targetColumn + 1).setValue(dealName);
          targetSheet.getRange(targetRow, targetColumn + 3).setValue(salesPerson);
          targetSheet.getRange(targetRow, targetColumn + 4).setValue(amount);
        }
      }
    }
  }

  // Hide fully empty rows in the specified ranges
  for (const [salesPerson, [startRow, endRow]] of Object.entries(salesPeople)) {
    for (let r = startRow; r <= endRow; r++) {
      let isEmpty = true;
      for (let c = 2; c <= targetSheet.getLastColumn(); c++) {
        if (targetSheet.getRange(r, c).getValue()) {
          isEmpty = false;
          break;
        }
      }
      if (isEmpty) {
        targetSheet.hideRows(r);
      } else {
        targetSheet.showRows(r); // Ensure non-empty rows are shown
      }
    }
  }
}

// Function to set a time-driven trigger to run the script
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger('autoFormatReport')
    .timeBased()
    .everyDays(1) // Adjust as needed
    .create();
}
