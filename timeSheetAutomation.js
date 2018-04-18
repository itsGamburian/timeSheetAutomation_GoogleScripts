/**********************
 Hambartzum Gamburian
 timeSheetAutomation.gs
 Google Scripts Code
 **********************/

function getEmployees() {
  var employeesArray = ['User1', 'User2', 'User3', 'User4'];
  return employeesArray;
}


function timeSheets() {
  
  //Today's Date
  var today = new Date(); //today.setDate(today.getDate() - 1); //DEBUGGING FOR RUNNING ON DIFFERENT TESTING DAYS
  var day = today.getDate();
  var month = today.getMonth() + 1; //January is 0
  var year = today.getFullYear();
  var dayOfWeek = today.getDay(); 
  
  //Adding a 0 before month and day, ie. July will be 07 instead of 7
  if (day < 10) day = '0' + day;
  if (month < 10) month = '0' + month;
 
  
  //The + 1 and the + 2 and so on is the amount of days you are adding to today
  var todayPlus1 = new Date();
      todayPlus1.setDate(todayPlus1.getDate() + 1); //todayPlus1.setDate(todayPlus1.getDate() - 1); //DEBUGGING FOR RUNNING ON DIFFERENT TESTING DAYS
  var todayPlus1Day = todayPlus1.getDate();
  var todayPlus1Month = todayPlus1.getMonth() + 1; //January is 0
  var todayPlus1Year = todayPlus1.getFullYear();
  
  
  var todayPlus13 = new Date();
      todayPlus13.setDate(todayPlus13.getDate() + 13); //todayPlus13.setDate(todayPlus13.getDate() - 1); //DEBUGGING FOR RUNNING ON DIFFERENT TESTING DAYS
  var todayPlus13Day = todayPlus13.getDate();
  var todayPlus13Month = todayPlus13.getMonth() + 1; //January is 0
  var todayPlus13Year = todayPlus13.getFullYear();
  
 
  var todayMinus14 = new Date();
      todayMinus14.setDate(todayMinus14.getDate() - 14); //todayMinus14.setDate(todayMinus14.getDate() - 1); //DEBUGGING FOR RUNNING ON DIFFERENT TESTING DAYS
  var todayMinus14Day = todayMinus14.getDate();
  var todayMinus14Month = todayMinus14.getMonth() + 1; //January is 0
  var todayMinus14Year = todayMinus14.getFullYear();
  
  var todayMinus1 = new Date();
      todayMinus1.setDate(todayMinus1.getDate() - 1); //todayMinus1.setDate(todayMinus1.getDate() - 1); //DEBUGGING FOR RUNNING ON DIFFERENT TESTING DAYS
  var todayMinus1Day = todayMinus1.getDate();
  var todayMinus1Month = todayMinus1.getMonth() + 1; //January is 0
  var todayMinus1Year = todayMinus1.getFullYear();
 
  
  if (todayPlus13Day < 10) todayPlus13Day = '0' + todayPlus13Day;
  if (todayPlus13Month < 10) todayPlus13Month = '0' + todayPlus13Month;
  
  if (todayMinus14Day < 10) todayMinus14Day = '0' + todayMinus14Day;
  if (todayMinus14Month < 10) todayMinus14Month = '0' + todayMinus14Month;
  
  if (todayMinus1Day < 10) todayMinus1Day = '0' + todayMinus1Day;
  if (todayMinus1Month < 10) todayMinus1Month = '0' + todayMinus1Month;
  
  
  var todaysDate = month + '/' + day;
  var todayPlus13Date = todayPlus13Month + '/' + todayPlus13Day;
  var timeFrame = todaysDate + ' to ' + todayPlus13Date;
  
  var todayMinus14Date = todayMinus14Month + '/' + todayMinus14Day;
  var todayMinus1Date = todayMinus1Month + '/' + todayMinus1Day;
  var lastTimeFrame = todayMinus14Date + ' to ' + todayMinus1Date;
  
              
  var employees = getEmployees();
  
  var employeeFileId = new Array(employees.length);
  
  var file, files;
  
  
  if (dayOfWeek == 0) { // Sunday is 0
    
    var i = 0;
    
    do {
     
      files = DriveApp.getFilesByName(employees[i] + ' TimeSheet ' + year);
      
      while (files.hasNext()) {
        file = files.next();
        employeeFileId[i] = file.getId();    
      }
    
      var ss = SpreadsheetApp.openById(employeeFileId[i]);
   
      var mostRecentSheet = ss.getSheetByName(lastTimeFrame);
    
      var tempMostRecentDateString = mostRecentSheet.getRange('P7').getValue().toLocaleString();
  
      var mostRecentDateString = tempMostRecentDateString.substring(0,16);
    
      var mostRecentDatePlus1 = new Date(Date.parse(mostRecentDateString));
          mostRecentDatePlus1.setDate(mostRecentDatePlus1.getDate() + 1);
      var mostRecentDatePlus1Day = mostRecentDatePlus1.getDate();
      var mostRecentDatePlus1Month = mostRecentDatePlus1.getMonth() + 1;
      var mostRecentDatePlus1Year = mostRecentDatePlus1.getFullYear();
    
      if (mostRecentDatePlus1Day < 10) mostRecentDatePlus1Day = '0' + mostRecentDatePlus1Day;
      if (mostRecentDatePlus1Month < 10) mostRecentDatePlus1Month = '0' + mostRecentDatePlus1Month;
    
      var mostRecentDatePlus1Date = mostRecentDatePlus1Month + '/' + mostRecentDatePlus1Day;

      Logger.log(str.concat(employees[i], "-", timeFrame, "-Most recent date: ", mostRecentDatePlus1Date));   
      
      //Copying sheet to new sheet
      var sheet = ss.getSheetByName('Template').copyTo(ss);
      
      //Naming sheet by the timeFrame deccared earlier
      sheet.setName(timeFrame);
      
      //Make the new sheet active
      ss.setActiveSheet(sheet);
      
      //Set the date on one cell, other cells will be that date plus 1 or 2 or 3 all setup already in "Template" sheet
      sheet.getRange('C7').setValue(todaysDate);
      
      //Clear the Projects List in the Main TimeSheet
      sheet.getRange('A8:A25').clearContent();
      
      sheet.getRange('A5').setValue(timeFrame);
      
      //If the sheet "Template" is hidden before copying,
      //Then we have to show the sheet because it automatically becomes hidden
      sheet.showSheet();  
      
      //Make the sheet all the way on the left
      ss.moveActiveSheet(1);
      
        i++;
    } while(i < employees.length);
  }
  
  SpreadsheetApp.flush();
}

function copyOverProjects() {
  
   //Today's Date
  var today = new Date();
  var day = today.getDate();
  var month = today.getMonth() + 1; //January is 0
  var year = today.getFullYear();
  var dayOfWeek = today.getDay(); 
  
  var employees = theEmployees();
  
  var employeeFileId = new Array(employees.length);
  
  var file, files;
  
  // max execution time = 6 mins
  var maxRowNumber = 103;
  var localTimeSheetProjectSheet = 'A3:A103';
  var globalProjectSheet = 'D4:D104';

  
  var i = 0;
    
    do {
     
      files = DriveApp.getFilesByName(employees[i] + ' TimeSheet ' + year);
      
      while (files.hasNext()) {
        file = files.next();
        employeeFileId[i] = file.getId();    
      }
   
      var ss = SpreadsheetApp.openById(employeeFileId[i]);
    
    
      //Copying Projects from Main Projects Spreadsheet to TimeSheet Spreadsheet
      var destination = ss.getSheetByName('Projects');
        
      destination.getRange(localTimeSheetProjectSheet).clearContent();
        
        
      var source = SpreadsheetApp.openById('1PE2lK2-1OPNjlzuA01Hldxb-MHYg6jgdtR_8mSOqXD0'); 
        
      var sourceSheet = source.getSheetByName('ForTech Projects');
        
      var sourceSheetRangeValues = sourceSheet.getRange(globalProjectSheet).getValues().toString();
        
      var sourceValuesArray = sourceSheetRangeValues.split(',');
        
        
      
      for (var k = 0, l = 3; k < sourceValuesArray.length, l < maxRowNumber; k++, l++) {
          destination.getRange('A' + l).setValue(sourceValuesArray[k]);
      }
        
      
      i++;
  } while(i < employees.length);
  
  
  SpreadsheetApp.flush();
}  

function createTrigger() {
  
  ScriptApp.newTrigger('timeSheets').timeBased().inTimezone('America/Los_Angeles')
  .everyWeeks(2).onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(9).create();  
  
}  