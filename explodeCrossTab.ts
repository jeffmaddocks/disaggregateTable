async function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();

    // Get the values from the range
    let rangeValues: Array<Array<string | number | boolean>> = range.getValues();

    // Process the data
    let lst: Array<[string, string]> = [];

    for (let i = 0; i < rangeValues.length; i++) {
        let col0 = rangeValues[i][0];
        // console.log(col0)
        for (let j = 1; j < rangeValues[i].length; j++) {
          let colCount = rangeValues[i][j];
          let colName = rangeValues[0][j];
          // console.log(colName)
            if (typeof colCount === "number" && colCount > 0) {
                for (let k = 0; k < colCount; k++) {
                    lst.push([col0.toString(), colName.toString()]);
                }
            }
        }
    }

    // Create a new sheet
    let newSheet = workbook.addWorksheet(sheet.getName() + '_exploded ');
    newSheet.getRange().setNumberFormat('@');

    // Write data to the new sheet
    for (let i = 0; i < lst.length; i++) {
        newSheet.getRange(`A${i + 1}`).setValue(lst[i][0]);
        newSheet.getRange(`B${i + 1}`).setValue(lst[i][1]);
    }
}