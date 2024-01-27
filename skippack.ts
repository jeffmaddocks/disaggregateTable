async function main(workbook: ExcelScript.Workbook) {
  let sheet = workbook.getActiveWorksheet();
  let range = sheet.getUsedRange();

  // Get the values in the range
  let rangeValues = range.getValues();
  let lst: string[] = [];

  let defaultzip = "19473";

  // Process the data
  for (let i = 0; i < rangeValues.length; i++) {  //looping through the rows
    let thisrow: string[] = [
      "Zip code",
      "Vaccinations received",
      "Other vaccinations received",
      "Age range",
      "Accompanied to event",
      "Disabled",
      "Disability type",
      "Race",
      "Other race",
      "Gender identification",
      "Other gender identification",
      "Sexual orientation",
      "Other sexual orientation",
      "Primary language",
      "Other primary language",
      "Event Feedback"
    ];
    
    if (i>0) {

      let skiptonext = false; // start off each row assuming that there will be a vaccine received

      for (let k = 0; k < thisrow.length; k++) { // loop through all the columns of the cumulus sheet
        thisrow[k] = ""; // set all cells to be blank unless they are filled below
      }

      for (let j = 0; j < rangeValues[i].length; j++) {  //looping through the columns
        let thiscell: string = typeof rangeValues[i][j] === 'string' ? rangeValues[i][j].toLowerCase() : String(rangeValues[i][j]);
        let colName: string = rangeValues[0][j].trim(); //this is the name of the column (row 0 of the column)
        if (skiptonext) {
          break;
        }

        if (colName == "Zip") {
          let usezip = "";

          usezip = thiscell;
          if (usezip.includes('-')) {
            usezip = usezip.split('-')[0]; // Trim off anything after a dash
          }
          if (/^[0-9]+$/.test(usezip) && parseInt(usezip) >= 0) { // Check if usezip is a positive number
              usezip = usezip.padStart(5, '0'); // Format usezip as a 5-digit zip code
          } else {
              usezip = defaultzip; // set to the default zip code if negative or includes any letters
          }

          thisrow[0] = usezip;

        } else if (colName == "Vaccine") {
          let received = "";
          let received_oth = "";
          if (thiscell.includes("covid") || thiscell.includes("pfizer") || thiscell.includes("moderna") || thiscell.includes("novavax")) { received = received + "Second (or later) dose of a COVID-19 vaccine; " }
          if (thiscell.includes("flu")) { received = received + "Flu/Influenza vaccine; " }
          if (thiscell.includes("pneumococcal") || thiscell.includes("prevnar") || thiscell.includes("pneumonia")) { received = received + "Pneumococcal; " }
          if (thiscell.includes("shingles") || thiscell.includes("shingrix")) { received = received + "Shingles; " }
          if (thiscell.includes("rsv") || thiscell.includes("arexvy")) { received = received + "Respiratory Syncytial Virus (RSV); " }
          if (
              thiscell.includes("other") ||
              thiscell.includes("engerix") ||
              thiscell.includes("boostrix") ||
              thiscell.includes("tdap") ||
              thiscell.includes("dtap") ||
              thiscell.includes("ipv-child") ||
              thiscell.includes("menquadfi") ||
              thiscell.includes("m-m-r") ||
              thiscell.includes("tenivac") ||
              thiscell.includes("varivax")
            ) { 
              received = received + "Other; " 
              switch (true) {
                case thiscell.includes("engerix"):
                  received_oth = received_oth + "HepB; ";
                  break;
                case thiscell.includes("boostrix"):
                  received_oth = received_oth + "TDap; ";
                  break;
                case thiscell.includes("boostrix") || thiscell.includes("tdap"):
                  received_oth = received_oth + "TDap; ";
                  break;
                case thiscell.includes("dtap"):
                  received_oth = received_oth + "DTaP; ";
                  break;
                case thiscell.includes("ipv-child"):
                  received_oth = received_oth + "Polio; ";
                  break;
                case thiscell.includes("menquadfi"):
                  received_oth = received_oth + "Meningococcal; ";
                  break;
                case thiscell.includes("m-m-r"):
                  received_oth = received_oth + "MMR; ";
                  break;
                case thiscell.includes("tenivac"):
                  received_oth = received_oth + "Tetanus/Diphtheria; ";
                  break;
                case thiscell.includes("varivax"):
                  received_oth = received_oth + "Varicella; ";
                  break;
              }
            }
          if (thiscell.includes("none") || thiscell == "") { 
            received = received + "None; " 
            skiptonext = true;
          }
          received = received.slice(0, -2);  // Trim the last two characters
          received_oth = received_oth.slice(0, -2);  // Trim the last two characters
          thisrow[1] = received;
          thisrow[2] = received_oth;
          if (skiptonext) {
            // if no vaccine received then insert a blank row
            // thisrow[0] =  "";
            // thisrow[1] = "";
            break;
          }

        } else if (colName == "Age") {

          let age = rangeValues[i][j] as number;
          let ageCategory: string;
          if (thiscell == "" ) {
            ageCategory = 'I prefer not to answer';
          } else if (age >=0 && age < 18) {
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
              // ageCategory = 'I prefer not to answer';
            }
          }
          thisrow[3] = ageCategory;
          thisrow[4] = "I prefer not to answer"; // Accompanied to event
          thisrow[5] = "I prefer not to answer"; // Disabled

        } else if (colName.includes("Ethnicity")) {
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
            if (thiscell.includes("answer") || thiscell == "" || usetext == "") { usetext = usetext + "I prefer not to answer; " }
            usetext = usetext.slice(0, -2);  // Trim the last two characters
          }
          thisrow[7] = usetext;

        } else if (colName == "Gender") {
          let usetext = "";
          if (i == 0) {
            usetext = 'Gender identification';
          } else {
            if ((thiscell.includes("male") && !thiscell.includes("female")) || thiscell == "m") { usetext = usetext + "Male; " }
            if (thiscell.includes("female") || thiscell == "f") { usetext = usetext + "Female; " }
            if (thiscell.includes("trans")) { usetext = usetext + "Transgender; " }
            if (thiscell.includes("binary")) { usetext = usetext + "Non-binary or gender non-conforming person; " }
            if (thiscell.includes("different") || thiscell.includes("other")) { usetext = usetext + "Different identity; " }
            if (thiscell.includes("answer") || thiscell == "") { usetext = usetext + "I prefer not to answer; " }
            usetext = usetext.slice(0, -2);  // Trim the last two characters
          }
          // thisrow.push(usetext);
          thisrow[9] = usetext;
          thisrow[11] = "I prefer not to answer"; // sexual orientation
          thisrow[13] = "I prefer not to answer"; // primary language

        }
      }
    }
    lst.push(thisrow);
  }

  // Create a new sheet
  let newSheet = workbook.addWorksheet(sheet.getName() + '_split');
  newSheet.getRange().setNumberFormat('@');

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