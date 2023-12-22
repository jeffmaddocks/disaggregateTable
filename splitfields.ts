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
      let thiscell: string = typeof rangeValues[i][j] === 'string' ? rangeValues[i][j].toLowerCase() : String(rangeValues[i][j]);
      let colName: string = rangeValues[0][j].trim(); //this is the name of the column (row 0 of the column)

      if (colName == "Received Vaccine") {
        let received = "";
        // if (thiscell.includes("covid")) { received = received + "Second (or later) dose of a COVID-19 vaccine; " }
        if (thiscell.includes("covid") || thiscell.includes("pfizer") || thiscell.includes("moderna")) { received = received + "Second (or later) dose of a COVID-19 vaccine; " }
        if (thiscell.includes("flu")) { received = received + "Flu/Influenza vaccine; " }
        if (thiscell.includes("pneumococcal") || thiscell.includes("prevnar")) { received = received + "Pneumococcal; " }
        if (thiscell.includes("shingles") || thiscell.includes("shingrix")) { received = received + "Shingles; " }
        if (thiscell.includes("rsv") || thiscell.includes("arexvy")) { received = received + "Respiratory Syncytial Virus (RSV); " }
        if (thiscell.includes("engerix")) { received = received + "HepB; " }
        if (thiscell.includes("boostrix")) { received = received + "TDap; " }
        if (thiscell.includes("other")) { received = received + "Other; " }
        if (thiscell.includes("none")) { received = received + "None; " }
        received = received.slice(0, -2);  // Trim the last two characters
        if (i == 0) { received = "Vaccinations received" }
        thisrow.push(received);

      } else if (colName == "Gender") {
        let usetext = "";
        if (i == 0) {
          usetext = 'Gender identification';
        } else {
          if (thiscell.includes("male") && !thiscell.includes("female")) { usetext = usetext + "Male; " }
          if (thiscell.includes("female")) { usetext = usetext + "Female; " }
          if (thiscell.includes("trans")) { usetext = usetext + "Transgender; " }
          if (thiscell.includes("binary")) { usetext = usetext + "Non-binary or gender non-conforming person; " }
          if (thiscell.includes("different")) { usetext = usetext + "Different identity; " }
          if (thiscell.includes("answer")) { usetext = usetext + "I prefer not to answer; " }
          usetext = usetext.slice(0, -2);  // Trim the last two characters
        }
        thisrow.push(usetext);

      } else if (colName.includes("Demographic Information")) {
        let usetext = "";
        if (i == 0) {
          usetext = 'Race';
        } else {
          if (thiscell.includes("indian") || thiscell.includes("alaska") || thiscell.includes("indigenous")) { usetext = usetext + "American Indian, Alaska Native, or Indigenous; " }
          if (thiscell.includes("asian")) { usetext = usetext + "Asian or Asian American; " }
          if (thiscell.includes("black") || thiscell.includes("african")) { usetext = usetext + "Black or African American; " }
          if ((!thiscell.includes("not hispanic")) && (thiscell.includes("hispanic") || thiscell.includes("latin") || thiscell.includes("mexican"))) { usetext = usetext + "Hispanic, Latino/a/x, or Latin American; " }
          if (thiscell.includes("middle") || thiscell.includes("north")) { usetext = usetext + "Middle Eastern, or North African; " }
          if (thiscell.includes("multiple")) { usetext = usetext + "Multiple races or ethnicities; " }
          if (thiscell.includes("hawaiian") || thiscell.includes("islander")) { usetext = usetext + "Native Hawaiian or Other Pacific Islander; " }
          if (thiscell.includes("white")) { usetext = usetext + "White/Caucasian; " }
          if (thiscell.includes("other")) { usetext = usetext + "Other; " }
          if (thiscell.includes("answer")) { usetext = usetext + "I prefer not to answer; " }
          usetext = usetext.slice(0, -2);  // Trim the last two characters
        }
        thisrow.push(usetext);

      } else if (colName == "Age") {

        let age = rangeValues[i][j] as number;
        let ageCategory: string;
        if (age < 18) {
          ageCategory = 'Under age 18';
        } else if (age >= 18 && age < 50) {
          ageCategory = 'Age 18 - 49';
        } else if (age >= 50 && age < 55) {
          ageCategory = 'Age 50 - 54';
        } else if (age >= 55 && age < 60) {
          ageCategory = 'Age 55 - 59';
        } else if (age >= 60 && age < 65) {
          ageCategory = 'Age 60 - 64';
        } else if (age >= 65 && age < 75) {
          ageCategory = 'Age 65 - 74';
        } else if (age >= 75 && age < 85) {
          ageCategory = 'Age 75 - 84';
        } else if (age >= 85) {
          ageCategory = 'Age 85+';
        } else {
          if (i == 0) {
            ageCategory = 'Age range';
          } else {
            ageCategory = 'I prefer not to answer';
          }
        }
        thisrow.push(ageCategory);

      // } else if (colName == "Patient Date of Birth") {

      //   let excelDateValue = rangeValues[i][j] as number;
      //   let birthDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));

      //   let today = new Date();
      //   let age = today.getFullYear() - birthDate.getFullYear();

      //   let m = today.getMonth() - birthDate.getMonth();
      //   if (m < 0 || (m === 0 && today.getDate() < birthDate.getDate())) { //account for dates where the birthday hasn't happened yet
      //     age--;
      //   }
      //   let ageCategory: string;
        // if (age < 18) {
        //   ageCategory = 'Under age 18';
        // } else if (age >= 18 && age < 50) {
        //   ageCategory = 'Age 18 - 49';
        // } else if (age >= 50 && age < 55) {
        //   ageCategory = 'Age 50 - 54';
        // } else if (age >= 55 && age < 60) {
        //   ageCategory = 'Age 55 - 59';
        // } else if (age >= 60 && age < 65) {
        //   ageCategory = 'Age 60 - 64';
        // } else if (age >= 65 && age < 75) {
        //   ageCategory = 'Age 65 - 74';
        // } else if (age >= 75 && age < 85) {
        //   ageCategory = 'Age 75 - 84';
        // } else if (age >= 85) {
        //   ageCategory = 'Age 85+';
        // } else {
      //     if (i == 0) {
      //       ageCategory = 'Age range';
      //     } else {
      //       ageCategory = 'I prefer not to answer';
      //     }
      //   }
      //   thisrow.push(ageCategory);

      } else { // if the column is named anything else, just push it
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