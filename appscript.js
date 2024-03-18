//Add a toolbar menu for manual solutions from the sheet rather than requiring users to access the code.
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Sync", functionName: "saveAsCSV"}];
  ss.addMenu("Sync List", csvMenuEntries);
};

//Attach file to email to submit to Clover systems.
function saveAsCSV() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  // decide the folder that you want to store your information
  var folder = DriveApp.getFolderById('INSERT FOLDER');
  //For this function to work, a sheet will need to be named Responses
  var sheet = ss.getSheetByName('Responses');
  const date = Utilities.formatDate(new Date(), "GMT+1", "MM/dd/yyyy")
  // append ".csv" extension to the sheet name
  fileName = 'Manual.csv';
  // convert all available sheet data to csv format
  var csvFile = convertRangeToCsvFile_(fileName, sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvFile);
  //File download
  const identification = file.getId();
  var file2 = DriveApp.getFileById(identification);

  GmailApp.sendEmail('INSERT EMAIL', '', 'Content-ID: <'+identification+'>', {
    name: 'ATTACH CITY',
    attachments: [file2.getAs(MimeType.CSV)]
});
}
  



function convertRangeToCsvFile_(csvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
