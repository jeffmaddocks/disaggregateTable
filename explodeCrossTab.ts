async function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();
  
    // Get the values from the range
    let rangeValues: Array<Array<string | number | boolean>> = range.getValues();
  
    // Process the data
    let lst: Array<[number | string | string]> = [];
    let icount = 0;
    for (let i = 0; i < rangeValues.length; i++) {  //looping through the rows
      let col0 = rangeValues[i][0].toString();
      // console.log(col0)
      if (i == 0) {
        lst.push([null, col0.toString(), null] as [number, string, string])
      }
      for (let j = 1; j < rangeValues[i].length; j++) {  //looping through the columns
        let colCount = rangeValues[i][j]; //this is the number in the cell of the crosstable
        let colName = rangeValues[0][j]; //this is the name of the column (row 0 of the column)
        // console.log(colName)
        if (typeof colCount === "number" && colCount > 0) {
          for (let k = 0; k < colCount; k++) {  //looping through the value in the cell X number of times creating rows
            icount = icount + 1
            lst.push([icount, col0.toString(), colName.toString()] as [number, string, string]);
          }
        }
      }
    }
  
    // Create a new sheet
    let newSheet = workbook.addWorksheet(sheet.getName() + '_exploded ');
    // newSheet.getRange().setNumberFormat('@');
  
    // Write data to the new sheet
    for (let i = 0; i < lst.length; i++) { //looping through the rows
      for (let j = 0; j < lst[i].length; j++) { //looping through the columns
        // Convert column number (j) to letter (A, B, C, ...)
        let columnLetter = String.fromCharCode(65 + j); // 65 is the ASCII value for 'A'
        if (j > 0) { //ensure that everything after column A is formatted as text
          newSheet.getRange(`${columnLetter}${i + 1}`).setNumberFormat('@');
        }
        newSheet.getRange(`${columnLetter}${i + 1}`).setValue(lst[i][j]);
      }
    }
  }