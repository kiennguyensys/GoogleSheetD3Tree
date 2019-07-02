var express = require('express');
var router = express.Router();
"use strict";
if(typeof require !== 'undefined') XLSX = require('xlsx');

  var fs = require('fs');


/*person class*/
function handleXLSX(workbook){

class Person {
    constructor (type, mail, condition) {
        this.type = type;
        this.mail = mail;
        this.condition = condition;
    }
}

class Node {
    constructor (person, name , r , c) {
        this.person = person;
        this.name = name;
        this.r = r;
        this.c = c;
        this.children = [];
    }
}





/* Get worksheet */
var worksheet = workbook.Sheets[workbook.SheetNames[1]];
var configsheet = workbook.Sheets[workbook.SheetNames[2]];

var metadata_firstcol = (configsheet['D2'] ? configsheet['D2'].v : undefined);
var metadata_lastcol = (configsheet['D3'] ? configsheet['D3'].v : undefined);
var metadata_firstrow = (configsheet['F2'] ? configsheet['F2'].v : undefined);
var metadata_lastrow = (configsheet['F3'] ? configsheet['F3'].v : undefined);

var data_firstcol = (configsheet['D4'] ? configsheet['D4'].v : undefined);
var data_lastcol = (configsheet['D5'] ? configsheet['D5'].v : undefined);
var data_firstrow = (configsheet['F4'] ? configsheet['F4'].v : undefined);
var data_lastrow = (configsheet['F5'] ? configsheet['F5'].v : undefined);

/* Find desired cell */
var persons = [];
//var range1 = sheet['!ref'];
var range = {s:{c:metadata_firstcol, r:metadata_firstrow}, e:{c:metadata_lastcol, r:metadata_lastrow}};
for(var R = range.s.r; R <= range.e.r; ++R) {
  var person = new Person();
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
    /* if an A1-style address is needed, encode the address */
    var cell_ref = XLSX.utils.encode_cell(cell_address);
  //  console.log(cell_ref);
    var cell_value = worksheet[cell_ref]? worksheet[cell_ref].v : undefined;
    if (C == range.s.c) person.type = cell_value;
    if (C == range.s.c+1) person.mail = cell_value;
    if (C == range.e.c) person.condition = cell_value;
  }
  persons[R]=person;
}
// console.log(persons);
range = {s:{c:data_firstcol, r:data_firstrow}, e:{c:data_lastcol, r:data_lastrow}};
var nodes = [];
for(var R = range.s.r; R <= range.e.r; ++R) {
  var node = new Node();
  for(var C = range.s.c; C <= range.e.c; ++C) {
    var cell_address = {c:C, r:R};
    /* if an A1-style address is needed, encode the address */
    var cell_ref = XLSX.utils.encode_cell(cell_address);
  //  console.log(cell_ref);
    if(worksheet[cell_ref])
    {
    var cell_value = worksheet[cell_ref]? worksheet[cell_ref].v : undefined;
    node.person = persons[R];
    node.name = cell_value;
    node.r = R;
    node.c = C;
    }
  }
  nodes[R]=node;
}



for(var C = range.e.c; C >= range.s.c; --C) {
  var node = new Node();
  for(var i = range.s.r; i <= range.e.r; ++i) {

  if(nodes[i].c == C)
    for(var j = nodes[i].r; j >= range.s.r; --j){
      if(nodes[j].c == C-1)
      {
        nodes[j].children.push(nodes[i]);
        break;

      }

    }

  }
}
  var json = JSON.stringify(nodes[data_firstrow]);//
  //console.log(json);
  fs.writeFile('./public/nodes.json', json, 'utf8',(err) => {
      if (err) {
          console.error(err);
          return;
      };
      console.log("File has been created");
  });
}


/* *************** Google Drive */

const readline = require('readline');
const {google} = require('googleapis');
var auth1;

// If modifying these scopes, delete token.json.
const SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/drive.metadata.readonly'];
// The file token.json stores the user's access and refresh tokens, and is
// created automatically when the authorization flow completes for the first
// time.
const TOKEN_PATH = 'token.json';

// Load client secrets from a local file.
fs.readFile('credentials.json', (err, content) => {
  if (err) return console.log('Error loading client secret file:', err);
  // Authorize a client with credentials, then call the Google Drive API.
  authorize(JSON.parse(content), listFiles);
});

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
    if (err) return getAccessToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    auth1 = oAuth2Client;
    callback(oAuth2Client);
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 * @param {google.auth.OAuth2} oAuth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback for the authorized client.
 */
function getAccessToken(oAuth2Client, callback) {
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
      if (err) return console.error('Error retrieving access token', err);
      oAuth2Client.setCredentials(token);
      // Store the token to disk for later program executions
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      auth1 = oAuth2Client;
      callback(oAuth2Client);
    });
  });
}

/**
 * Lists the names and IDs of up to 10 files.
 * @param {google.auth.OAuth2} auth An authorized OAuth2 client.
 */
function listFiles(auth) {
  const drive = google.drive({version: 'v3', auth});
  drive.files.list({
    pageSize: 10,
    fields: 'nextPageToken, files(id, name)',
  }, (err, res) => {
    if (err) return console.log('The API returned an error: ' + err);
    const files = res.data.files;
    if (files.length) {
      console.log('Files:');
      files.map((file) => {
        console.log(`${file.name} (${file.id})`);
      });
    } else {
      console.log('No files found.');
    }
  });
}
/* ***************GoogleDrive */


function downloadFile(auth, fileId) {
  const drive = google.drive({version: 'v3', auth});
 var dest = fs.createWriteStream('data_set.xlsx');
 // example code here
 drive.files.export({fileId: fileId, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}, {responseType: 'stream'},
 function(err, res){
    res.data
    .on('end', () => {
       console.log('Done');

    })
    .on('Error', err => {
       console.log('Error', err);
    })
    .pipe(dest);
 });
}




/* GET home page. */
router.get('/', (req, res, next) => {
    res.render('index', {title: 'Project X'});

});

router.get('/filedownload', (request, response, next) => {
    var raw_url = request.query.text.toString();
    var arr = raw_url.split('/');
    var fileId = arr[5];
    downloadFile(auth1, fileId.toString());
//    response.send('(Drive Public) Spreadsheet Id: '+fileId);
setTimeout(function () {
     // after 2 seconds

     var workbook = XLSX.readFile('data_set.xlsx');
     handleXLSX(workbook);
     response.redirect('..');
  }, 5000)




});


router.put('/update_a_food', (request, response, next) => {
    response.end("PUT requested => update_a_food");
});

router.delete('/delete_a_food', (request, response, next) => {
    response.end("DELETE requested => delete_a_food");
});

module.exports = router;
