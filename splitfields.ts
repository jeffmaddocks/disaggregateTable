async function main(workbook: ExcelScript.Workbook) {
    let sheet = workbook.getActiveWorksheet();
    let range = sheet.getUsedRange();

    // Get the values in the range
    let rangeValues = range.getValues();
    let lst: string[] = [];

    // Process the data
    for (let i = 0; i < rangeValues.length; i++) {  //looping through the rows
        let thisrow: string[] = [];
        for (let j = 0; j < rangeValues[i].length; j++) {  //looping through the columns
            let colName = rangeValues[0][j]; //this is the name of the column (row 0 of the column)
            let thiscell: string = typeof rangeValues[i][j] === 'string' ? rangeValues[i][j].toLowerCase() : String(rangeValues[i][j]);
            if (colName == "Received Vaccine") {
              let received = "";
              if (thiscell.includes("covid")) { received = received + "Second (or later) dose of a COVID-19 vaccine; " }
              if (thiscell.includes("flu")) { received = received + "Flu/Influenza vaccine; " }
              if (thiscell.includes("pneumococcal")) { received = received + "Pneumococcal; " }
              if (thiscell.includes("shingles")) { received = received + "Shingles; " }
              if (thiscell.includes("rsv")) { received = received + "Respiratory Syncytial Virus (RSV); " }
              if (thiscell.includes("other")) { received = received + "Other; " }
              if (thiscell.includes("none")) { received = received + "None; " }
              received = received.slice(0, -2);  // Trim the last two characters
            
              thisrow.push(received);
            } else if (colName == "Patient Date of Birth") {
              thisrow.push(thiscell);
            } else {
              thisrow.push(thiscell);
            }
        }
        lst.push(thisrow);
    }

    // Create a new sheet
    let newSheet = workbook.addWorksheet(sheet.getName() + '_split');
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