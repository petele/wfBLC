'use strict';

let fs = require('fs');
let blc = require('broken-link-checker');
let chalk = require('chalk');
let readline = require('readline');
let google = require('googleapis');
let moment = require('moment');
let googleAuth = require('google-auth-library');

process.env['NODE_TLS_REJECT_UNAUTHORIZED'] = '0';

const VERBOSE = false;
const SITE_URL = 'https://web-central.appspot.com/web/';
const SPREADSHEET_ID = '1ObBKWXu0KQ7yaew8VvG-eArXXyIX64sSSseXRZRADuU';
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const TOKEN_DIR = (process.env.HOME || process.env.HOMEPATH ||
    process.env.USERPROFILE) + '/.credentials/';
const TOKEN_PATH = TOKEN_DIR + 'sheets.googleapis.com-blc.json';

let pageData;
let authToken;
let siteChecker;
let pagesCompleted = 0;
let startedAt = moment().format();
let sheets = google.sheets('v4');

console.log('Broken Link Checker for', chalk.bold('/web'));
console.log('Started at:', chalk.cyan(startedAt));
console.log('');

let opts = {
  cacheExpiryTime: 3 * 60 * 60 * 1000,
  cacheResponses: true,
  excludedKeywords: [],
  excludedSchemes: [],
  excludeExternalLinks: false,
  filterLevel: 3,
  honorRobotExclusions: false,
  rateLimit: 10,
  requestMethod: 'get' 
};

function padString(msg) {
  return (msg + '                    ').slice(0, 7);
}

function resetPageData(pageUrl) {
  console.log(pageUrl);
  pageData = {
    currentPage: pageUrl,
    linkCount: 0,
    linkOK: 0,
    linkBroken: 0,
    linkExcluded: 0,
    brokenLinks: []
  };
}

function handleUrlResult(result) {
  pageData.linkCount++;
  let status;
  let resolved = result.url.resolved;
  let original = result.url.original;
  let simpleUrl = resolved || original;
  if (result.broken) {
    pageData.linkBroken++;
    logBrokenLink('ERROR', resolved, original, result.brokenReason);
  } else if (result.excluded) {
    pageData.linkExcluded++;
    if (VERBOSE) {
      status = chalk.yellow(padString('SKIPPED'));
      let reason = chalk.gray(result.excludedReason)
      console.log('->', status, simpleUrl, reason);
    }
  } else {
    pageData.linkOK++;
    if (VERBOSE) {
      status = chalk.green(padString('OK'));
      if (result.http.cached === true) {
        status = chalk.green(padString('CACHED'));
      }
      console.log('->', status, simpleUrl);
    }
  }
}

function saveSummary() {
  console.log('');
  console.log('Broken Link Check Completed.');
  let finishedAt = moment().format();
  console.log('Finished at:', chalk.cyan(finishedAt));
  console.log('Checked', chalk.cyan(pagesCompleted), 'pages.');
  console.log('');
  let res = {
    range: 'Summary!B3',
    majorDimension: 'ROWS',
    values: [
      ['Finished'],
      [startedAt],
      [finishedAt]
    ]
  }
  return updateSheet(res);
}

function logBrokenLink(status, resolved, original, reason) {
  let url = resolved || original;
  console.log('->', chalk.red(padString(status)), url, reason);
  let label = '/web/' + pageData.currentPage.replace(SITE_URL, '');
  let sourceUrl = '=hyperlink("' + pageData.currentPage + '", "' + label + '")';
  pageData.brokenLinks.push([sourceUrl, reason, resolved, original]);
}

function saveErrorsToSheet() {
  if (pageData.brokenLinks.length > 0) {
    let range = 'Errors!A2:D2';
    let resource = {
      range: range,
      majorDimension: 'ROWS',
      values: pageData.brokenLinks
    }
    sheets.spreadsheets.values.append({
      auth: authToken,
      spreadsheetId: SPREADSHEET_ID,
      range: range,
      valueInputOption: 'USER_ENTERED',
      resource: resource
    }, function(err, response) {
      if (err) {
        console.log(chalk.red('saveErrorsToSheet FAILED'), err);
      }
    });
  }
}

function logPageCompleted() {
  saveErrorsToSheet();
  pagesCompleted++;
  let msg = '';
  msg += 'Links: ' + chalk.cyan(pageData.linkCount) + ' | ';
  msg += 'OK: ' + chalk.green(pageData.linkOK) + ' | ';
  msg += 'Broken: ' + chalk.red(pageData.linkBroken) + ' | ';
  msg += 'Skipped: ' + chalk.yellow(pageData.linkExcluded) + ' | ';
  console.log('', msg);
  msg = 'Pages Completed: ' + chalk.cyan(pagesCompleted) + ' | ';
  msg += 'In Queue: ' + chalk.cyan(siteChecker.numPages());
  console.log('', msg);
  console.log('');
  let range = 'Pages!A2:E2';
  let label = '/web/' + pageData.currentPage.replace(SITE_URL, '');
  let sourceUrl = '=hyperlink("' + pageData.currentPage + '", "' + label + '")';
  let resource = {
    range: range,
    majorDimension: 'ROWS',
    values: [[
      sourceUrl,
      pageData.linkCount,
      pageData.linkOK,
      pageData.linkBroken,
      pageData.linkExcluded
    ]]
  }
  sheets.spreadsheets.values.append({
    auth: authToken,
    spreadsheetId: SPREADSHEET_ID,
    range: range,
    valueInputOption: 'USER_ENTERED',
    resource: resource
  }, function(err, response) {
    if (err) {
      console.log(chalk.red('logPageCompleted FAILED'), err);
    }
  });
}

let handlers = {
  html: function(tree, robots, response, pageUrl) {
    resetPageData(pageUrl);
  },
  junk: handleUrlResult,
  link: handleUrlResult,
  page: function(error, pageUrl) {
    if (error) {
      resetPageData(pageUrl);
      if (error.code !== 200) {
        pageData.linkCount++;
        pageData.linkBroken++;
        let errorCode = 'HTTP_' + error.code;
        logBrokenLink('ERROR', error.message, '', errorCode);
      }
    }
    logPageCompleted();
  },
  end: function() { process.exit(0); },
  robots: function() { console.log('robots'); },
  site: saveSummary
};


// Load client secrets from a local file.
fs.readFile('client_secret.json', function processClientSecrets(err, content) {
  if (err) {
    console.log(chalk.red('Error loading client secret file: '), err);
    return;
  }
  // Authorize a client with the loaded credentials, then call the
  // Google Sheets API.
  authorize(JSON.parse(content), onAuthorized);
});

/**
 * Create an OAuth2 client with the given credentials, and then execute the
 * given callback function.
 *
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
  console.log('Authorizing...');
  var clientSecret = credentials.installed.client_secret;
  var clientId = credentials.installed.client_id;
  var redirectUrl = credentials.installed.redirect_uris[0];
  var auth = new googleAuth();
  var oauth2Client = new auth.OAuth2(clientId, clientSecret, redirectUrl);

  // Check if we have previously stored a token.
  fs.readFile(TOKEN_PATH, function(err, token) {
    if (err) {
      getNewToken(oauth2Client, callback);
    } else {
      oauth2Client.credentials = JSON.parse(token);
      callback(oauth2Client);
    }
  });
}

/**
 * Get and store new token after prompting for user authorization, and then
 * execute the given callback with the authorized OAuth2 client.
 *
 * @param {google.auth.OAuth2} oauth2Client The OAuth2 client to get token for.
 * @param {getEventsCallback} callback The callback to call with the authorized
 *     client.
 */
function getNewToken(oauth2Client, callback) {
  var authUrl = oauth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES
  });
  console.log('Authorize this app by visiting this url: ', authUrl);
  var rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });
  rl.question('Enter the code from that page here: ', function(code) {
    rl.close();
    oauth2Client.getToken(code, function(err, token) {
      if (err) {
        console.log('Error while trying to retrieve access token', err);
        return;
      }
      oauth2Client.credentials = token;
      storeToken(token);
      callback(oauth2Client);
    });
  });
}

/**
 * Store token to disk be used in later program executions.
 *
 * @param {Object} token The token to store to disk.
 */
function storeToken(token) {
  try {
    fs.mkdirSync(TOKEN_DIR);
  } catch (err) {
    if (err.code != 'EEXIST') {
      throw err;
    }
  }
  fs.writeFile(TOKEN_PATH, JSON.stringify(token));
  console.log('Token stored to ' + TOKEN_PATH);
}

function readRange(range) {
  return new Promise(function(resolve, reject) {
    sheets.spreadsheets.values.get({
      auth: authToken,
      spreadsheetId: SPREADSHEET_ID,
      range: range,
    }, function(err, response) {
      if (err) {
        console.log(chalk.red('readRange FAILED'), err);
        reject(err);
      }
      resolve(response.values);
    });  
  });
}

function updateSheet(resource) {
  return new Promise(function(resolve, reject) {
    sheets.spreadsheets.values.update({
      auth: authToken,
      spreadsheetId: SPREADSHEET_ID,
      range: resource.range,
      valueInputOption: 'USER_ENTERED',
      resource: resource
    }, function(err, response) {
      if (err) {
        console.log(chalk.red('updateSheet FAILED'), err);
        reject(err);
      }
      resolve(response);
    });
  });
}

function resetWorkbook(workbook) {
  console.log('Resetting workbook...');
  return new Promise(function(resolve, reject) {
    var requests = [];
    // Summary Page
    requests.push({
      updateCells: {
        start: {sheetId: 635298754, rowIndex: 2, columnIndex: 0},
        rows: [{
          values: [{
            userEnteredValue: {stringValue: 'Current Status'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Running'},
          }]
        }],
        fields: 'userEnteredValue,userEnteredFormat.textFormat'
      }
    });
    requests.push({
      updateCells: {
        start: {sheetId: 635298754, rowIndex: 3, columnIndex: 0},
        rows: [{
          values: [{
            userEnteredValue: {stringValue: 'Started At'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: moment().format()},
          }]
        }],
        fields: 'userEnteredValue,userEnteredFormat.textFormat'
      }
    });
    requests.push({
      updateCells: {
        start: {sheetId: 635298754, rowIndex: 4, columnIndex: 0},
        rows: [{
          values: [{
            userEnteredValue: {stringValue: 'Finished At'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: ''},
          }]
        }],
        fields: 'userEnteredValue,userEnteredFormat.textFormat'
      }
    });
    // Errors
    requests.push({
      updateCells: {
        range: {
          sheetId: 0
        },
        fields: 'userEnteredValue'
      }
    });
    requests.push({
      updateCells: {
        start: {sheetId: 0, rowIndex: 0, columnIndex: 0},
        rows: [{
          values: [{
            userEnteredValue: {stringValue: 'Source URL'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Issue'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Resolved URL'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Original URL'},
            userEnteredFormat: {textFormat: {bold: true}}
          }]
        }],
        fields: 'userEnteredValue,userEnteredFormat.textFormat'
      }
    });
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: 0,
          gridProperties: {
            frozenRowCount: 1
          }
        },
        fields: 'gridProperties.frozenRowCount'
      }
    });
    // Pages
    requests.push({
      updateCells: {
        range: {
          sheetId: 352617043
        },
        fields: 'userEnteredValue'
      }
    });
    requests.push({
      updateCells: {
        start: {sheetId: 352617043, rowIndex: 0, columnIndex: 0},
        rows: [{
          values: [{
            userEnteredValue: {stringValue: 'Source URL'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Links'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'OK'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Broken'},
            userEnteredFormat: {textFormat: {bold: true}}
          }, {
            userEnteredValue: {stringValue: 'Skipped'},
            userEnteredFormat: {textFormat: {bold: true}}
          }]
        }],
        fields: 'userEnteredValue,userEnteredFormat.textFormat'
      }
    });
    requests.push({
      updateSheetProperties: {
        properties: {
          sheetId: 0,
          gridProperties: {
            frozenRowCount: 1
          }
        },
        fields: 'gridProperties.frozenRowCount'
      }
    });

    let batchUpdateRequest = {requests: requests};
    sheets.spreadsheets.batchUpdate({
      auth: authToken,
      spreadsheetId: SPREADSHEET_ID,
      resource: batchUpdateRequest
    }, function(err, response) {
      if (err) {
        console.log(chalk.red('resetWorkbook FAILED'), err);
        reject(err);
      }
      console.log('->', 'Workbook reset');
      console.log('');
      resolve(response);
    });
  });
}

function getExcludes() {
  console.log('Retreiving excludes...');
  return new Promise(function(resolve, reject) {
    readRange('ExcludeKeywords!A2:B')
      .then(function(rows) {
        if (rows && rows.length > 0) {
          rows.forEach(function(row) {
            if (row[0]) {
              let urlExclude = row[0].trim();
              if (urlExclude.length > 0) {
                opts.excludedKeywords.push(urlExclude);
              }
            }

            if (row[1]) {
              let schemeExclude = row[1].trim();
              if (schemeExclude.length > 0) {
                opts.excludedSchemes.push(schemeExclude);
              }
            }
          });
        }
        console.log('->', 'Keywords Excluded:', chalk.yellow(opts.excludedKeywords.join(', ')));
        console.log('->', 'Schemes  Excluded:', chalk.yellow(opts.excludedSchemes.join(', ')));
        console.log('');
        resolve(true);
      });
  });
}

function onAuthorized(auth) {
  console.log('->', 'Authorization:', chalk.green('OK'));
  console.log('');
  authToken = auth;
  getExcludes()
    .then(resetWorkbook)
    .then(function() {
      console.log('Starting link checker at:', chalk.cyan(SITE_URL));
      console.log('');
      siteChecker = new blc.SiteChecker(opts, handlers);
      siteChecker.enqueue(SITE_URL);
    })
    .catch(function(err) {
      console.log(chalk.red('CRITIAL FAILURE in onAuthorized.'));
      console.log(err);
    });
}
