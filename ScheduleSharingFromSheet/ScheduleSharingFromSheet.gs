function ScheduleSharingFromSheet() {
/**
 * Turn on/off sharing by link for a list of documents with ID's from the spreadsheet 
 * https://docs.google.com/spreadsheets/d/1s3zZX43jd5H7yaAqLpRcTk4GTfaVN5t851ki4XFCWcE/edit
 */
 eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());
 var spreadsheetId = '1s3zZX43jd5H7yaAqLpRcTk4GTfaVN5t851ki4XFCWcE';
 var rangeName = 'Sheet1!A2:C';
 var values = Sheets.Spreadsheets.Values.get(spreadsheetId, rangeName).values;
 var dateFormat ='DD/MM/Y';
 if (!values) {
    Logger.log('No data found.');
  } else {
    Logger.log('Date-time data: ' + values.length +' records found');
    var now = moment();
    //var now = moment('2020-05-14');
    Logger.log(' Now is %s',now.format());
    for (var row = 0; row < values.length; row++) {
      fromDate = moment(values[row][0].toString());
      toDate = moment(values[row][1].toString()).add(1,'day');
      var URL = values[row][2].toString();
      Logger.log('Row ' + (row+1) +': Obtained URL: '+URL);
      if ((!fromDate.isValid()) || (!toDate.isValid())) {
        Logger.log('Row ' + (row+1) +': Either %s or %s contains invalid date skipping. ',values[row][0].toString(),values[row][1].toString());
        // Skip if either of From/To dates is invalid 
        continue;
      }
      try {
        
        // Extract the File or Folder ID from the Drive URL
        var id = URL.match(/[-\w]{25,}$/);
        Logger.log('Row ' + (row+1) + ': Extracted ID: '+id.toString());
        if (id) {
          asset = DriveApp.getFileById(id) ? DriveApp.getFileById(id) : DriveApp.getFolderById(id);
          if (asset) {
            //Logger.log('Asset obtained');
            // Modify sharing state based on the current date
            if (now.isBetween(fromDate,toDate)) {
              //Logger.log('Row ' + (row+1) +': %s is within the date range %s and %s',now.format(dateFormat),fromDate.format(dateFormat),toDate.format(dateFormat));
              // Enable access
              Logger.log('Enabling access to %s', id.toString());
              asset.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            } else {
              //Logger.log('Row ' + (row+1) +': %s is outside of date range %s and %s',now.format(dateFormat),fromDate.format(dateFormat),toDate.format(dateFormat));
              // Make the folder / file Private
              Logger.log('Disabling access to %s', id.toString());
              asset.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.NONE);
            }              
          }
        }
      } catch (e) {
        
        Logger.log(e.toString());
      } // try
    } // for (var row = 0; row < values.length; row++)
  } // if (!values) {
}  // end function
