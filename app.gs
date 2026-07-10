// Google Apps Script - Keramas Tracker Backend
// Deployment: Publish → New Deployment → Web app

let spreadsheet = null;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('keramas_tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Initialize Spreadsheet
function initializeSheet() {
  spreadsheet = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  
  // Create sheets if not exist
  const sheets = spreadsheet.getSheets();
  const sheetNames = sheets.map(s => s.getName());
  
  if (!sheetNames.includes('users')) {
    spreadsheet.insertSheet('users');
    spreadsheet.getSheetByName('users').appendRow(['username', 'password', 'name', 'role', 'lastWash', 'rewardsClaimed']);
  }
  
  if (!sheetNames.includes('proofs')) {
    spreadsheet.insertSheet('proofs');
    spreadsheet.getSheetByName('proofs').appendRow(['username', 'date', 'imageUrl', 'status', 'rewardClaimed']);
  }
  
  if (!sheetNames.includes('rewards')) {
    spreadsheet.insertSheet('rewards');
    spreadsheet.getSheetByName('rewards').appendRow(['imageUrl', 'uploadDate', 'active']);
  }
  
  if (!sheetNames.includes('broadcasts')) {
    spreadsheet.insertSheet('broadcasts');
    spreadsheet.getSheetByName('broadcasts').appendRow(['message', 'sendDate', 'readBy']);
  }
}

// User Management
function getAllUsers() {
  const sheet = spreadsheet.getSheetByName('users');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    username: row[0],
    password: row[1],
    name: row[2],
    role: row[3],
    lastWash: row[4],
    rewardsClaimed: row[5]
  }));
}

function getUserByUsername(username) {
  const users = getAllUsers();
  return users.find(u => u.username === username);
}

function addUser(username, password, name) {
  const sheet = spreadsheet.getSheetByName('users');
  sheet.appendRow([username, password, name, 'user', '', 0]);
}

// Proof Management
function submitProof(username, imageUrl) {
  const sheet = spreadsheet.getSheetByName('proofs');
  sheet.appendRow([username, new Date().toISOString(), imageUrl, 'pending', false]);
  
  // Update user's lastWash
  const userSheet = spreadsheet.getSheetByName('users');
  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username) {
      userSheet.getRange(i + 1, 5).setValue(new Date().toISOString());
      break;
    }
  }
}

function getPendingProofs() {
  const sheet = spreadsheet.getSheetByName('proofs');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(row => row[3] === 'pending').map(row => ({
    username: row[0],
    date: row[1],
    imageUrl: row[2],
    status: row[3],
    rewardClaimed: row[4]
  }));
}

function verifyProof(username, date) {
  const sheet = spreadsheet.getSheetByName('proofs');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === username && data[i][1] === date) {
      sheet.getRange(i + 1, 4).setValue('verified');
      break;
    }
  }
}

// Reward Photos
function addRewardPhoto(imageUrl) {
  const sheet = spreadsheet.getSheetByName('rewards');
  sheet.appendRow([imageUrl, new Date().toISOString(), true]);
}

function getRewardPhotos() {
  const sheet = spreadsheet.getSheetByName('rewards');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).filter(row => row[2] === true).map(row => ({
    imageUrl: row[0],
    uploadDate: row[1]
  }));
}

// Broadcast
function sendBroadcast(message) {
  const sheet = spreadsheet.getSheetByName('broadcasts');
  sheet.appendRow([message, new Date().toISOString(), '']);
}

function getBroadcasts() {
  const sheet = spreadsheet.getSheetByName('broadcasts');
  const data = sheet.getDataRange().getValues();
  return data.slice(1).map(row => ({
    message: row[0],
    sendDate: row[1],
    readBy: row[2] ? row[2].split(',') : []
  }));
}

// Utility Functions
function uploadImage(blob) {
  const folder = DriveApp.getFoldersByName('Keramas Tracker').hasNext() 
    ? DriveApp.getFoldersByName('Keramas Tracker').next()
    : DriveApp.createFolder('Keramas Tracker');
  
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  return file.getUrl();
}

// Initialize on deployment
function onInstall() {
  onOpen();
}

function onOpen() {
  // Initialize sheet when script opens
  try {
    initializeSheet();
  } catch (e) {
    Logger.log('Sheet initialization error: ' + e);
  }
}
