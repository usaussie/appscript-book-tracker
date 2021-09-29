/**
 * 
 * CHANGE THESE VARIABLES PER YOUR PREFERENCES
 * 
 */
// sheet tab name containing the data
const DATA_TAB_NAME = 'Form Responses 1';

// Set sheet column order (using 0 index). IE: COLUMN A is 0, COLUMN B is 1
const COLUMN_NUMBER_ISBN = 2;
const COLUMN_NUMBER_TITLE = 5;

// Set sheet column order for updating from the API details that come back (starting with Column A = 1, Column B = 2)
const COLUMN_NUMBER_AUTHORS = 7;
const COLUMN_NUMBER_DESCRIPTION = 8;

// API URL
const API_URL = "https://www.googleapis.com/books/v1/volumes?country=US";

/**
 * 
 * DO NOT CHANGE ANYTHING UNDER THIS LINE
 * 
 * ONLY CHANGE THINGS IN THE CONFIG.GS FILE
 * 
 */

/**
* Do the lookup stuff
*/

function job_processGoogleSheetData() {
  
  // get current spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Log starting of the script
  Logger.log('Script has started');

  // get TimeZone
  var timeZone = ss.getSpreadsheetTimeZone();
  
  // get Data sheet
  var dataSheet = ss.getSheetByName(DATA_TAB_NAME);
  
  // get all data as a 2-D array
  var data = dataSheet.getDataRange().getValues();
  
  // create a name:value pair array to send the data to the next Function
  var spreadsheetData = {ss:ss, timeZone:timeZone, dataSheet:dataSheet, data:data};
  
  // run Function to create Google Folders
  var doIsbnLookup = isbnLookup_(spreadsheetData);
  
  // check success status
  if (doIsbnLookup) {
    // display Toast notification
    Logger.log('Finished Successfully');
  }
  else {
    // script completed with error
    // display Toast notification
    Logger.log('With errors. Please see Logs', 'Finished');
  }
  
  // Log starting of the script
  Logger.log('Script finished');
  
  
}


/**
* Loop through each row and create folders, set permissions
*/
function isbnLookup_(spreadsheetData) {
  
  // extract data from name:value pair array
  var ss = spreadsheetData['ss'];
  var timeZone = spreadsheetData['timeZone'];
  var dataSheet = spreadsheetData['dataSheet']; 
  var data = spreadsheetData['data'];

  // get last row number so we know when to end the loop
  var lastRow = dataSheet.getLastRow();

  var folderIdMap = new Object();

  // start of loop to go through each row iteratively
  for (var i=1; i<lastRow; i++) {
    
    // extract values from row of data for easier reference below
    var isbn = data[i][COLUMN_NUMBER_ISBN];
    var title = data[i][COLUMN_NUMBER_TITLE];
    
    // only perform this row if the title is blank
    if(title == '') {

      Logger.log('Looking Up ISBN:  ' + isbn);

      // run Function to get the book info from the API
      var bookData = getBookDetails_(isbn);

      // set the column number for the update method later
      var title_column_number = COLUMN_NUMBER_TITLE + 1;
      
      // check data came back correctly
      if (bookData) {

        // extract details into vars for easier reference later
        var title = (bookData["volumeInfo"]["title"]);
        var description = (bookData["volumeInfo"]["description"]);
        var authors = (bookData["volumeInfo"]["authors"]);

        // set values in apropriate columns
        dataSheet.getRange(i+1, title_column_number).setValue(title);
        dataSheet.getRange(i+1, COLUMN_NUMBER_AUTHORS).setValue(authors);
        dataSheet.getRange(i+1, COLUMN_NUMBER_DESCRIPTION).setValue(description);
        
        // write all pending updates to the google sheet using flush() method
        SpreadsheetApp.flush();
        
      } else {
        // write error into Title cell and return false value
        dataSheet.getRange(i+1, title_column_number).setValue('Error finding ISBN data. Please see Logs');
        return false;
      }

    } else {

      Logger.log('Skipping Row - ISBN Data already set - Parsing Next Row');

    }
    
  } // end of loop to go through each row in turn **********************************
  
  // completed successfully
  return true;
  
  
}

function getBookDetails_(isbn) {

  // Query the book database by ISBN code.
  //isbn = isbn || "9781451648546"; // Steve Jobs book

  var url = API_URL + "&q=isbn:" + isbn;

  var response = UrlFetchApp.fetch(url);
  var results = JSON.parse(response);

  if (results.totalItems) {

    // There'll be only 1 book per ISBN
    var book = results.items[0];

    // var title = (book["volumeInfo"]["title"]);
    // var description = (book["volumeInfo"]["description"]);
    // var subtitle = (book["volumeInfo"]["subtitle"]);
    // var authors = (book["volumeInfo"]["authors"]);
    // var printType = (book["volumeInfo"]["printType"]);
    // var pageCount = (book["volumeInfo"]["pageCount"]);
    // var publisher = (book["volumeInfo"]["publisher"]);
    // var publishedDate = (book["volumeInfo"]["publishedDate"]);
    // var webReaderLink = (book["accessInfo"]["webReaderLink"]);

    // // For debugging
    // Logger.log(description);
    // Logger.log(isbn);
    // Logger.log(book);

    return book;

  } 

  return false;

}
