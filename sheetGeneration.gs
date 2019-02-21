/* Author      : Haley Tortorich
 * Name        : sheetGeneration
 * Description : generates new sheets based on templates. stores values to
 *               a "scripting" work sheet and uses those to update week and
 *               spreadsheet names. Year and Week number are updated each
 *               time it is ran.
 */ 


/* set current year 
 * @param sh1 first sheet name
 * @param sh2 second sheet name
 * @param sh3 third sheet name
 */
function setYear(sh1, sh2, sh3) {
  var sh4 = getSheetName(sh1);
  var sh5 = getSheetName(sh2);
  var sh6 = getSheetName(sh3);
  var cell1 = sh4.getRange("E1:E1");
  var cell2 = sh5.getRange("E1:E1");
  var cell3 = sh6.getRange("E1:E1");
  
  // create array to store location that year will be added to
  var cellArray = [cell1, cell2, cell3]; 
  
  // loop through array and update year on sheet at cellArray location
  for(var i = 0; i < cellArray.length; i++) {
      var date = Utilities.formatDate(new Date(), "GMT+1", "yyyy");
      cellArray[i].setValue(date);
  }
}


/* sets week number 
 * @param sh1 first sheet name
 * @param sh2 second sheet name
 * @param sh3 third sheet name
 */
function setWeek(sh1, sh2, sh3) {
  // get the sheet from parameters and store to variable
  var sh4 = getSheetName(sh1);
  var sh5 = getSheetName(sh2);
  var sh6 = getSheetName(sh3);
  
  // get location to store new week from new sheets
  var cell1 = sh4.getRange("B1:B1");
  var cell2 = sh5.getRange("B1:B1");
  var cell3 = sh6.getRange("B1:B1");
  
  // get current week value from sheet
  var ss = getValueFromRange("scripting", "D2:D2");
  
  // set new week value on new sheets
  cell1.setValue(ss+1);
  cell2.setValue(ss+1);
  cell3.setValue(ss+1);
  
  // sets new week value into scripting sheet
  getSheetName("scripting").getRange("D2:D2").setValue(ss+1);
}

//creates new sheets, stores current sheet number and newest sheet names
function createSheetWeekly() {  
  // create variable to do work on sheet named "scripting"
  var scripting = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("scripting");
  
  // store base names of sheets
  var trucksBaseName = scripting.getRange("A2:A2");
  var bobcatsBaseName = scripting.getRange("B2:B2");
  var attachmentsBaseName = scripting.getRange("C2:C2");
  // track sheet number
  var number = scripting.getRange("J2:J2");
  // add 1 to current number
  var newnum = number.getValue() + 1;
  var setNumber = scripting.getRange("E2:E2");
  // updates scripting value tracking sheet number
  setNumber.setValue(newnum);
  
  // create new sheet name
  var newTruckSheet = trucksBaseName.getValue() + newnum; 
  var newBobcatSheet= bobcatsBaseName.getValue() + newnum;
  var newAttachmentSheet = attachmentsBaseName.getValue() + newnum;
  
  // create array containing new sheet names
  var sheetNameArray = [newTruckSheet, newBobcatSheet, newAttachmentSheet];
  
  // get location to store newest sheet names
  var locationA = scripting.getRange("F2:F2");
  var locationB = scripting.getRange("G2:G2");
  var locationC = scripting.getRange("H2:H2");

  
  // create array containing locations in the sheet to add new sheet name
  var locationArray = [locationA, locationB, locationC];
  
  // set value at locationArray[i] to sheetNameArray[i] value
  for(var i = 0; i < locationArray.length; i++) {
    locationArray[i].setValue(sheetNameArray[i]);
  }  

  // create new sheets using new sheet names
  for(var i = 0; i < sheetNameArray.length; i++) {
    SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetNameArray[i]);
  }
}

/* gets the value from a given sheetname and range
 * @param sheetName takes in a string to be used as the sheetname ie "scripting"
 * @param str takes in a string to be used as a range ie "A2:A2"
 */
function getValueFromRange(sheetName, str) {
  var scripting = getSheetName(sheetName);  
  return scripting.getRange(str).getValue();
}
/* gets sheet name from a given string
 * @param sheetName takes in a string to be used as the sheetname ie "scripting"
 */
function getSheetName(sheetName) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);  
}

/* Copies data from src file to target file
 * @param src takes in a sheetname as source 
 * @param trgt takes in a sheetname as target
 * @param takes in a bool, checks if it should getFormulas and store to new sheet 
 */
function copyDataToNewFile(src, trgt, checkFormula) {
  // copy sheet values from src to trgt 
  var temp1Range = src.getDataRange();
  // get A1 notation identifying the range
  var A1Range = temp1Range.getA1Notation();
  // get background of src
  var backGround = temp1Range.getBackgrounds();
  // get fontsize of src
  var fontSize = temp1Range.getFontSizes();
  // get fontcolors of src
  var fontColors = temp1Range.getFontColors();
  // get fontstyles of src
  var fontStyles = temp1Range.getFontStyles();
  // get the range of src that contains formulas
  var formula = src.getRange("D3:D9");
  // get formulas
  var formulas = formula.getFormulas();
  
  // get the values in the range
  var tValues = temp1Range.getValues();
  // set the range of target to values
  trgt.getRange(A1Range).setValues(tValues);
  // apply formula to correct sheet
  if(checkFormula == true) {
    // set formulas in trgt 
    trgt.getRange("D3:D9").setFormulas(formulas);
    // get range src that contains formatting
    var tt = src.getRange("D3:D9");
    // copy formatting of src to trgt
    tt.copyFormatToRange(trgt.getSheetId(), 4, 4, 3, 9);
  }
  // set background on trgt
  trgt.getRange(A1Range).setBackgrounds(backGround);
  // set fontsizes on trgt
  trgt.getRange(A1Range).setFontSizes(fontSize);
  // set fontcolors on trgt
  trgt.getRange(A1Range).setFontColors(fontColors);
  // set fontstyles on trgt
  trgt.getRange(A1Range).setFontStyles(fontStyles);

}

//XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
//-----------------------------DRIVER---------------------------------------
function main() {
  createSheetWeekly();
    
  // get newest sheet name
  var newTruckSheet = getValueFromRange("scripting", "F2:F2");
  var newBobcatSheet = getValueFromRange("scripting", "G2:G2");
  var newAttachmentSheet= getValueFromRange("scripting", "H2:H2");
    
  // get newest sheet names and store it into a variable
  var sh1 = getSheetName(newTruckSheet);
  var sh2 = getSheetName(newBobcatSheet);
  var sh3 = getSheetName(newAttachmentSheet);
  
  
  
  // get template sheetnames and store into a variable
  var temp1 = getSheetName("truck_template");
  var temp2 = getSheetName("bobcat_template");
  var temp3 = getSheetName("attachment_template");
  
  
  
  // copy templates to new sheets
  copyDataToNewFile(temp1, sh1, false);
  copyDataToNewFile(temp2, sh2, true);
  copyDataToNewFile(temp3, sh3, false);
  
  // create variable for active Spreadsheet
  var spread = SpreadsheetApp.getActiveSpreadsheet();
  
  // create an array with newest sheet names
  var sheetsArray = [sh1, sh2, sh3];
  
  // set active sheet and move active sheet
  for (var i = 0; i < sheetsArray.length; i++) {
    spread.setActiveSheet(sheetsArray[i]);
    spread.moveActiveSheet(i + 1);
  }
  
  //wait 3000 millisecs
  Utilities.sleep(3000);
  
  // setup new sheets (add year, weekly, other data)
  setYear(newTruckSheet, newBobcatSheet, newAttachmentSheet);
  setWeek(newTruckSheet, newBobcatSheet, newAttachmentSheet);
}

//-----------------------------END DRIVER-----------------------------------
//XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
