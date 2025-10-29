function columnTitleUpdates(){
  var app = SpreadsheetApp;

  var masterSS = app.openById("1nSK79KRlNgu8wMdsYFDnai9MeYYFBxVpDxJ8CmSrh24");
  var hrSS = app.openById("1rkufQVPUOOSE-Zyh0pt8Rdg_jpZ-LgUTg6ZtiT7R8QE");
  var prSS = app.openById("1KqYqSAz5cMUacvz19zUXC3jBe2bpesK6p-DFejnlQGs");

  var masterSheet = masterSS.getSheetByName("Master List");
  var hrSheet = hrSS.getSheetByName("HR Employees");
  var prSheet = prSS.getSheetByName("PR Employees");

Logger.log("TEST");

//These data sets are 2D arrays [row, column] with all data for each sheet
  var masterData = masterSheet.getDataRange().getValues(); 
  var hrData = hrSheet.getDataRange().getValues(); 
  var prData = prSheet.getDataRange().getValues(); 

//Convert the 2D arrays into 1D arrays of objects.  Column titles are object object keys, and the value in that column is the object value.
  var masterAO = createArrayOfObjects1(masterData);
  var hrAO = createArrayOfObjects1(hrData);
  var prAO = createArrayOfObjects1(prData);

//Get headers for each sheet as a 1D array
var masterHeaders = masterData[0];
var hrHeaders = hrData[0];
var prHeaders = prData[0];


//For each sheet, find the column number of each header title.  These header titles (the string inside indexOf()), need to be manually entered and identical to the column titles as they actually appear on the sheet.

var msFirstNameCol = masterHeaders.indexOf("First Name")+1;
var msLastNameCol = masterHeaders.indexOf("Last Name")+1;
var msAnnualSalaryCol = masterHeaders.indexOf("Annual Salary")+1;
var msTitleCol = masterHeaders.indexOf("Title")+1;
var msLocationCol = masterHeaders.indexOf("Location")+1;
var msDepositCodeCol = masterHeaders.indexOf("Deposit Code")+1;
var msBusinessUnitCol = masterHeaders.indexOf("Business Unit")+1;
var msHireDateCol = masterHeaders.indexOf("Hire Date")+1;

var hrFirstNameCol = hrHeaders.indexOf("First Name")+1;
var hrLastNameCol = hrHeaders.indexOf("Last Name")+1;
var hrAnnualSalaryCol = hrHeaders.indexOf("Annual Salary")+1;
var hrTitleCol = hrHeaders.indexOf("Title")+1;
var hrLocationCol = hrHeaders.indexOf("Location")+1;
var hrDepositCodeCol = hrHeaders.indexOf("Deposit Code")+1;
var hrBusinessUnitCol = hrHeaders.indexOf("Business Unit")+1;
var hrHireDateCol = hrHeaders.indexOf("Hire Date")+1;
var hrEmailCol = hrHeaders.indexOf("Email")+1;

var prFirstNameCol = prHeaders.indexOf("First Name")+1;
var prLastNameCol = prHeaders.indexOf("Last Name")+1;
var prAnnualSalaryCol = prHeaders.indexOf("Annual Salary")+1;
var prTitleCol = prHeaders.indexOf("Title")+1;
var prLocationCol = prHeaders.indexOf("Location")+1;
var prDepositCodeCol = prHeaders.indexOf("Deposit Code")+1;
var prBusinessUnitCol = prHeaders.indexOf("Business Unit")+1;
var prHireDateCol = prHeaders.indexOf("Hire Date")+1;
var prEmailCol = prHeaders.indexOf("Email Address")+1;




//Get a 1D array of [FirstName+" "+LastName] for all lists.  The slice(1) is used to remove the header when needed.
  var masterNames = masterData.slice(1).map(function(value){
    var mFirstName = value[msFirstNameCol-1];
    var mLastName = value[msLastNameCol-1];
    return(mFirstName+" "+mLastName)
  })

  var hrNames = hrData.slice(1).map(function(value){
    var hrFirstName = value[hrFirstNameCol-1];
    var hrLastName = value[hrLastNameCol-1];
    return(hrFirstName+" "+hrLastName)
  })

  var prNames = prData.slice(1).map(function(value){
    var prFirstName = value[prFirstNameCol-1];
    var prLastName = value[prLastNameCol-1];
    return(prFirstName+" "+prLastName)
  })



//------------------------------------------------------------------ ADD TO PAYROLL --------------------------------------------------------------------------------------------
//Find the names that are in the master list but not the Payroll list, return as a 1D array.  Find the object in the Master Array of Objects, so that the corresponding info can be retrieved and entered into the Payroll list

  const inMasterNotPr = masterNames.filter(item => !prNames.includes(item));
  var inMasterNotPrLength = inMasterNotPr.length;

  for(var i = 0; i<inMasterNotPrLength; i++){
    var listName = inMasterNotPr[i];

    Logger.log("listName is "+listName+" at i = "+i)

    for (var t=0; t<masterAO.length; t++){

      Logger.log("at t = "+t+", masterAO[First Name is ]"+masterAO[t]["First Name"])

      if(masterAO[t]["First Name"]+" "+masterAO[t]["Last Name"] === listName){
        var newRow = prSheet.getLastRow()+1;
        prSheet.getRange(newRow,prAnnualSalaryCol,1,1).setValue(masterAO[t]["Annual Salary"]);
        prSheet.getRange(newRow,prFirstNameCol,1,1).setValue(masterAO[t]["First Name"]);
        prSheet.getRange(newRow,prLastNameCol,1,1).setValue(masterAO[t]["Last Name"]);
        prSheet.getRange(newRow,prLocationCol,1,1).setValue(masterAO[t]["Location"]);
        prSheet.getRange(newRow,prDepositCodeCol,1,1).setValue(masterAO[t]["Deposit Code"]);
      }
    }
  }
//---------------------------------------------------- End ADD TO PAYROLL --------------------------------------------------------------------------------------------------------










//-------------------------------------------------------- REMOVE FROM PAYROLL ---------------------------------------------------------------------------------------------------
//Find the names that are in the Payroll list but not the Master, so that these can be removed from the Payroll list.  Remember that the index for prAO is actually TWO LESS than the row number, given the header row and zero indexing


//Get all of the Master and Payroll data from the sheets as 2D arrays
 masterData = masterSheet.getDataRange().getValues(); 
 prData = prSheet.getDataRange().getValues(); 

 //Take the Payroll 2D array and turn it into a 1D array of objects.  Find its length.
 prAO = createArrayOfObjects1(prData);
 var prAOLenth = prAO.length;


//Get the list of names [First Name+" "+Last Name] on the Master Sheet
    masterNames = masterData.slice(1).map(function(value){
    var mFirstName = value[msFirstNameCol-1];
    var mLastName = value[msLastNameCol-1];
    return(mFirstName+" "+mLastName)
  })


//Get the list of names [First Name+" "+Last Name] on the Payroll Sheet
  prNames = prData.slice(1).map(function(value){
    var prFirstName = value[prFirstNameCol-1];
    var prLastName = value[prLastNameCol-1];
    return(prFirstName+" "+prLastName)
  })


//Compare the two lists of names.  Make a new list of names, those that are on the Payroll List but not the Master List
  var inPrNotMaster = prNames.filter(item => !masterNames.includes(item)); 
  var inPrNotMasterLength = inPrNotMaster.length;
 



//Loop through Payroll array of objects.  For each [First + Last] name, check if it's on the list to delete, and if so delete it.  The list that contains rows to delete MUST be the outside loop and counted backwards, otherwise there will be unpredictable behavior.  Also, the prAO array uses i-1 for the index, and deleteRow uses i+1.  In total there is a difference of 2 between the index and the row to be deleted, due to the removal of the header row and zero-indexing of the array.

  for(var i =prAOLenth; i>0; i--){
    var prListName = prAO[i-1]["First Name"]+" "+prAO[i-1]["Last Name"]


    for(t=0; t<inPrNotMasterLength; t++){
      var prNameToDelete = inPrNotMaster[t];
      if(prListName === prNameToDelete){

        Logger.log("MATCH")

        prSheet.deleteRow(i+1);
              break;
      }

    }
  }
//--------------------------------------------------- End REMOVE FROM PAYROLL -----------------------------------------------------------------------------------------------------









//---------------------------------------------------- ADD EMAILS FROM 3RD SHEET ----------------------------------------------------------------------------------------------------

   hrData = hrSheet.getDataRange().getValues(); 
   prData = prSheet.getDataRange().getValues();

   for(i=1; i<prData.length; i++){
    if(prData[i][prEmailCol-1] == ""){
      var prName = prData[i][prFirstNameCol-1]+" "+prData[i][prLastNameCol-1];
      for(var t=0; t<hrData.length; t++){
        if(hrData[t][hrFirstNameCol-1]+" "+hrData[t][hrLastNameCol-1]===prName){
          prSheet.getRange(i+1,prEmailCol,1,1).setValue(hrData[t][hrEmailCol-1]);
        }
      }
    }
   } 



//----------------------------------------------------- End ADD EMAILS FROM 3RD SHEET -----------------------------------------------------------------------------------------

}



//This function takes a 2D array and converts it to a 1D array of objects.  The column titles are the object keys, and the value will be the object in that column(inner loop) for the given row (outer loop).

function createArrayOfObjects1(pageArray){

var result = [];  //Initialize 1D array

//Get headers from the first row in the whole sheet 2D arraty
var headers = pageArray[0];

  // Loop through each row of the Main Array
  for (var i = 1; i < pageArray.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = pageArray[i][j]; // Assign header as key
    }
    result.push(obj); // Add object to result array
  }
  return result;
}




