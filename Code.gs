/**
 * A special function that runs when the spreadsheet is first
 * opened or reloaded. onOpen() is used to add custom menu
 * items to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sjekk mapper')
    .addItem('Kjør funksjon', 'main')
    .addItem('Opprett tidsutløsere', 'TriggerCreationFromMenu')
    .addItem('Slett tidsutløsere', 'DeleteIfTriggersExistsFromMenu')
    .addToUi();
}


function getEmailAddresses(sheet){
  let myEmailAddress = Session.getActiveUser().getEmail();
  let otherEmailAddresses = sheet.getRange(1, 2).getValue();

  if (otherEmailAddresses.length >1){
    return myEmailAddress + ', ' +  otherEmailAddresses;
  } else {
    return myEmailAddress;
  }
}


/**
 * Send email once a file is created or updated in the folders
 */ 
function main(){
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('Main');
  var emailAddress = getEmailAddresses(sheet);
  var urlRow = 4;
  var currentDate = new Date();

  while (true){
    let folderUrl = String(sheet.getRange(urlRow, 1).getValue());
    //Logger.log('folderUrl: ' + folderUrl);
    if (folderUrl.length < 1){
      break;
    }

    let mainFolderId = GetIdFromUrl(folderUrl);
    Logger.log('mainFolderId: ' + mainFolderId);

    let mainFolder = DriveApp.getFolderById(mainFolderId);
    let mainFolderName = mainFolder.getName();
    Logger.log('mainFolderName: ' + mainFolderName);
    sheet.getRange(urlRow, 2).setValue(mainFolderName);

    let fileList = GetAllFolderInfo(folder = mainFolder, prefix = '', isMain = true); 
    Logger.log('Printing total fileList: ');
    Logger.log(JSON.stringify(fileList));

    var compareUpdated = sheet.getRange(urlRow,3).getValue();
    if (String(compareUpdated).length < 1){
      compareUpdated = currentDate
    };
    Logger.log('compareUpdated: ' + compareUpdated);

    var newFiles = FilterNewFiles(fileList, compareUpdated);
    Logger.log('Printing all updated files: \n' + JSON.stringify(newFiles));
    
    fileLastUpdatedDatetime = GetLastUpdatedDatetime(fileList);
    sheet.getRange(urlRow, 3).setValue(fileLastUpdatedDatetime);

    if (newFiles.length > 0){
      SendEmail(emailAddress=emailAddress,folderName=mainFolderName, newFiles=newFiles)
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


function GetFileInfo(folder, prefix){
    let folderFileList = [];
    let contents = folder.getFiles();  

    while (contents.hasNext()){
      let file = contents.next();
      let fileName = file.getName();
      let fullFileName = prefix + fileName;
      let fileLastUpdated = file.getLastUpdated();
      let fileUrl = file.getUrl();
      //Logger.log('fullFileName: ' + fullFileName);
      //Logger.log('Last Updated: ' + fileLastUpdated);
      folderFileList.push({fileUrl: fileUrl, fileName: fullFileName, fileLastUpdated: fileLastUpdated});
    }

  return folderFileList;
}


function GetAllFolderInfo(folder, prefix, isMain){
  let folderName = folder.getName();
  
  if (isMain == false){
    prefix = prefix + folderName + '/';
  }
  isMain = false;

  let fileList = GetFileInfo(folder, prefix);
  Logger.log('fileList from GetFileInfo:');
  Logger.log(JSON.stringify(fileList))

  let folderContents = folder.getFolders();
  Logger.log('folderContents.hasNext:' + folderContents.hasNext())
  while (folderContents.hasNext()){
    let newfolder = folderContents.next();
    let newFileList = GetAllFolderInfo(folder=newfolder, prefix, isMain);
    Logger.log('newFileList:')
    Logger.log(JSON.stringify(newFileList));
    fileList = fileList.concat(newFileList);
    Logger.log('Concatenated fileList:');
    Logger.log(JSON.stringify(fileList));
  }
  
  return fileList;
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


function DeleteIfTriggersExists(eventType, handlerFunction) {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (trigger) {
    if(String(trigger.getEventType()) === eventType &&
      String(trigger.getHandlerFunction()) === handlerFunction)
      ScriptApp.deleteTrigger(trigger);
  });
}


function DeleteIfTriggersExistsFromMenu(){
  DeleteIfTriggersExists(eventType="CLOCK", handlerFunction="main");
  SpreadsheetApp.getUi().alert("Tidsutløsere har blitt slettet");
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


function TriggerCreationFromMenu(){
  TriggerCreation();
  SpreadsheetApp.getUi().alert("Tidsutløsere har blitt opprettet");
}

