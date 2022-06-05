/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sjekk mapper')
    .addItem('Kjør funksjon', 'main')
    .addItem('Opprett tidsutløsere', 'TriggerCreation')
    .addToUi();
}

/**
 * Send email once a file is created or updated in the folders
 */ 
function main(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Main')
  var emailAddress = sheet.getRange(1, 2).getValue();
  var urlRow = 4;
  var currentDate = new Date();

  while (true){
    var folderUrl = String(sheet.getRange(urlRow, 1).getValue());
    //Logger.log('folderUrl: ' + folderUrl);
    if (folderUrl.length < 1){
      break;
    }

    var folderInfo = GetFolderInfo(folderUrl = folderUrl); 
    var folderName = folderInfo.folderName;
    var fileList = folderInfo.fileList;

    Logger.log('folderName: ' + folderName);
    sheet.getRange(urlRow, 2).setValue(folderName);

    Logger.log('Printing folder info: ');
    Logger.log(JSON.stringify(folderInfo));

    var compareUpdated = sheet.getRange(urlRow,3).getValue();
    if (String(compareUpdated).length < 1){
      compareUpdated = currentDate
    };
    Logger.log('compareUpdated: ' + compareUpdated);

    var newFiles = FilterNewFiles(fileList, compareUpdated);
    Logger.log('Printing updated files: \n' + JSON.stringify(newFiles));
    
    fileLastUpdatedDatetime = GetLastUpdatedDatetime(fileList);
    sheet.getRange(urlRow, 3).setValue(fileLastUpdatedDatetime);

    if (newFiles.length > 0){
      SendEmail(emailAddress=emailAddress,folderName=folderName, newFiles=newFiles)
    }

    urlRow++;
  }
}


function SendEmail(emailAddress, folderName, newFiles){
  var emailMessage = "";
  for (let i in newFiles){
    file = newFiles[i];
    emailMessage += "Fil \'" + file.fileName + "\' sist oppdatert " + file.fileLastUpdated + " med lenke: \n" + file.fileUrl + " \n\n"
  }

  var subject = 'Nye eller oppdaterte filer i: \'' + folderName +'\'';

  GmailApp.sendEmail(recipient=emailAddress, subject=subject, body=emailMessage)
}


function GetIdFromUrl(url) { return url.match(/[-\w]{25,}/); }


function GetFolderInfo(folderUrl){
  var folder, folderId, folderName, contents, fileName, fileLastUpdated, fileUrl;

  const fileList = [];

  folderId = GetIdFromUrl(folderUrl);
  //Logger.log('folderId: ' + folderId);
  folder = DriveApp.getFolderById(folderId);
  folderName = folder.getName();

  contents = folder.getFiles();  
  while (contents.hasNext()){
    var file = contents.next();
    fileName = file.getName();
    fileLastUpdated = file.getLastUpdated();
    fileUrl = file.getUrl();
    //Logger.log('fileName: ' + fileName);
    //Logger.log('Last Updated:' + fileLastUpdated);
    fileList.push({fileUrl: fileUrl, fileName: fileName, fileLastUpdated: fileLastUpdated});
  }

  return {folderName: folderName, fileList: fileList};
}


function GetLastUpdatedDatetime(inputList){
  var lastUpdated;

  if (inputList.length == 0){
    lastUpdated = new Date();
  }

  lastUpdated = inputList[0].fileLastUpdated;

  for (let i=1; i < inputList.length; i++){
    ts = inputList[i].fileLastUpdated;
    if (ts > lastUpdated){
      lastUpdated = ts
    }
  }

  return lastUpdated
}


function FilterNewFiles(fileList, compareDate) {
  result = [];
  for (let i in fileList){
    file = fileList[i];
    if (file.fileLastUpdated > compareDate){
      result.push(file);
    }
  }

  return result
}


function CreateTriggers() {
 ScriptApp.newTrigger("main")
   .timeBased()
   .atHour(5)
   .everyDays(2)
   .inTimezone("Europe/Oslo")
   .create();
}


function DeleteIfTriggersExists(eventType, handlerFunction) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if(String(trigger.getEventType()) === eventType &&
      String(trigger.getHandlerFunction()) === handlerFunction)
      ScriptApp.deleteTrigger(trigger);
  });
}


function TriggerCreation(){
  DeleteIfTriggersExists(eventType="CLOCK", handlerFunction="main");
  ScriptApp.newTrigger("main")
   .timeBased()
   .atHour(5)
   .everyDays(1)
   .inTimezone("Europe/Oslo")
   .create();
  ScriptApp.newTrigger("main")
   .timeBased()
   .atHour(16)
   .everyDays(1)
   .inTimezone("Europe/Oslo")
   .create();
}


