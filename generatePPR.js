const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
var pdfFillForm = require('pdf-fill-form');
var moment = require("moment");
var sourcePDF = "CoroonBlankPPR.pdf";
var maxRows = 12;



// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

//var orderSheetId = '1MznAbKY0VS9M4S1rhw_RJUoZ85AhWTJ1BtzU3BptTfg';
var orderSheetId = '1_AtNK9XKil42hGTc4bwOQWPhzCBfzeedn2uxZ5EIob8'; // 2020 sheet

// print process.argv
process.argv.forEach(function (val, index, array) {
  console.log(index + ': ' + val);
});

var startRow = process.argv[2];
var endRow = process.argv[3];

if(startRow == "list"){
  listPDFFields();
}else if(!startRow || !endRow){
  throw new Error("You need to specify the start and end row numbers from the spreadsheet");
}else{

  // Load client secrets from a local file.
  fs.readFile('credentials.json', (err, content) => {
    if (err) return console.log('Error loading client secret file:', err);
    // Authorize a client with credentials, then call the Google Sheets API.
    authorize(JSON.parse(content), function(auth){
    listOrderItems(auth, startRow, endRow, writeToPPR)});
  });
}
/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return console.error('Error while trying to retrieve access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}

/**
 * Prints the names and majors of students in a sample spreadsheet:
 * @see https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
 * @param {google.auth.OAuth2} auth The authenticated Google OAuth client.
 */
function listOrderItems(auth, startRow, endRow, callback) {
  const sheets = google.sheets({version: 'v4', auth});
  sheets.spreadsheets.values.get({
    spreadsheetId: orderSheetId,
    range: 'Order Requests!A2:I',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    var rows = res.data.values;
    if (rows.length) {
      console.log(rows.length);
      console.log(startRow + " : " + endRow);
      rows = rows.slice(startRow - 2, endRow);
      // Print columns A and E, which correspond to indices 0 and 4.
      var itemsArray = rows.map((row) => {
        console.log(row);
        var newRow = {
           "Order No" : row[1],
           "ITEM NUMBER  DESCRIPTION" : row[2],
           "Price" : row[3] ? row[3].replace("$","") : "",
           "Quantity" : row[4],
           "Total" : row[5] ? row[5].replace("$","") : "",
           "Vendor" : row[8],
        };
        return newRow;
      });
      console.log(JSON.stringify(itemsArray, null, "  "));
      callback(itemsArray);
    } else {
      console.log('No data found.');
    }
  });
}



function writeToPPR(itemsArray, addtlFields){
  if(itemsArray.length > 12){
    throw new Error("max number of valid row is 12. You're trying to process " + itemsArray.length + " rows");
  }
  var fields = {};
  var i = 1;
  var grandTotal = 0;
  var rowNum = 1;
  var vendor = false;
  var filename = "";
  for(var i = 0; i < itemsArray.length; i++){
    var numName = "Row"+rowNum;
    var item = itemsArray[i];
    if(
      //!item["ITEM NUMBER  DESCRIPTION"+numName] || 
      !item["Price"] ||
      item["Price"] == "" ||
      !item["Quantity"] ||
      item["Quantity"] == "" ||
      !item["Total"] ||
      false
      ){
      console.log("skipping");
      console.log(item);
      continue;
    }else{
      fields["Order No"+numName] = item["Order No"];
      fields["ITEM NUMBER  DESCRIPTION"+numName] = item["ITEM NUMBER  DESCRIPTION"];
      fields["Price"+numName] = "$"+item["Price"];
      fields["Quantity"+numName] = item["Quantity"];
      fields["TotalRow"+ (rowNum > 1 ? rowNum.toString().padStart(2,"0") : rowNum)] = "$"+item["Total"];
      grandTotal += parseFloat(item["Total"])
      rowNum++;      
    }
    if(item["Vendor"] && item["Vendor"].trim() != ""){
      if(vendor && vendor != item["Vendor"].trim()){
        throw new Error("Inconsistent Vendors: " + vendor + " and " + item["Vendor"].trim());
      }
      vendor = item["Vendor"].trim();
    }
  }
  fields["TotalTOTAL"] = Math.round(grandTotal * 100) / 100;
  fields["Name of Vendor 1"] = "";
  fields["Date"] = moment().format('MM/DD/YYYY');
  if(vendor){
    var vendorFields = vendor.split(",");
    fields["Name of Vendor 1"] = vendorFields[0];
    if(vendorFields[1]){
     fields["Vendor's Email"] = vendorFields[1];      
    }
    if(vendorFields[2]){
     fields["Vendor's Phone Number"] = vendorFields[2];      
    }
    if(vendorFields[3]){
      fields["Vendor's Contact"] = vendorFields[3];
    }
    if(vendorFields[4]){
      fields["Business Purpose"] = vendorFields[4];
    }

  }
  console.log("fields");
  console.log(JSON.stringify(fields, null, "  "));
  filename = fields["Name of Vendor 1"].replace(/[^0-9a-zA-Z]/g,"") + "_"+ moment().format('YYYY_MM_DD')+"_CoroonPPR.pdf";
  doWrite(fields,filename);


}


function doWrite(fields, filename){
  pdfFillForm.write(sourcePDF, fields, { "save": "pdf", 'cores': 4, 'scale': 0.2, 'antialias': true } )
  .then(function(result) {
      fs.writeFile(filename, result, function(err) {
          if(err) {
            return console.log(err);
          }
          console.log("The file was saved!");
      }); 
  }, function(err) {
      console.log(err);
  });
}


function listPDFFields(callback){

  pdfFillForm.read(sourcePDF)
    .then(function(result) {
      console.log ("listPDFFields");
      console.log(result);
    }, function(err) {
      console.log(err);
  });
}