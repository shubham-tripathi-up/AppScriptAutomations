function processAdmitCards() {
  var SHEET_NAME = "MAIN_SHEET"; // Change this if needed
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    Logger.log("Sheet not found: " + SHEET_NAME);
    return;
  }
  
  var data = sheet.getDataRange().getValues();
  var header = data[0];
  
  var columns = {};
  header.forEach((colName, index) => {
    columns[colName] = index;
  });
  
  var requiredColumns = ["student_code", "name", "photo", "sign", "new_photo_link", "new_sign_link"];
  
  for (var col of requiredColumns) {
    if (!(col in columns)) {
      Logger.log("Missing required column: " + col);
      return;
    }
  }
  
  var parentFolder = DriveApp.getRootFolder(); // Change this if needed
  var mainFolder = getOrCreateFolder(parentFolder, "Offline Exam - Student Documents");

  Logger.log(`All files will be stored in: ${mainFolder.getUrl()}`);
  
  var totalRows = data.length - 1;
  var processedRows = 0;
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var studentCode = row[columns["student_code"]];
    var studentName = row[columns["name"]];
    var photoUrl = row[columns["photo"]];
    var signUrl = row[columns["sign"]];
    var existingPhotoLink = row[columns["new_photo_link"]];
    var existingSignLink = row[columns["new_sign_link"]];

    if (existingPhotoLink && existingSignLink) {
      processedRows++;
      Logger.log(`${processedRows}/${totalRows} processed`);
      continue;
    }
    
    if (studentName && studentCode) {
      var newPhotoName = `${studentName}_${studentCode}_Photograph`;
      var newSignName = `${studentName}_${studentCode}_Signature`;
      
      // Process Photo
      if (photoUrl && !existingPhotoLink) {
        var existingPhoto = findFileInFolder(mainFolder, newPhotoName);
        if (existingPhoto) {
          sheet.getRange(i + 1, columns["new_photo_link"] + 1).setValue(existingPhoto.getUrl());
        } else {
          var copiedPhoto = copyFile(photoUrl, newPhotoName, mainFolder);
          if (copiedPhoto) {
            sheet.getRange(i + 1, columns["new_photo_link"] + 1).setValue(copiedPhoto.getUrl());
            Utilities.sleep(2000);
          }
        }
      }
      
      // Process Signature
      if (signUrl && !existingSignLink) {
        var existingSign = findFileInFolder(mainFolder, newSignName);
        if (existingSign) {
          sheet.getRange(i + 1, columns["new_sign_link"] + 1).setValue(existingSign.getUrl());
        } else {
          var copiedSign = copyFile(signUrl, newSignName, mainFolder);
          if (copiedSign) {
            sheet.getRange(i + 1, columns["new_sign_link"] + 1).setValue(copiedSign.getUrl());
            Utilities.sleep(2000);
          }
        }
      }
      
      processedRows++;
      Logger.log(`${processedRows}/${totalRows} processed`);
    }
  }

  Logger.log(`Processing complete. Total rows processed: ${processedRows}/${totalRows}`);
}

function getOrCreateFolder(parent, folderName) {
  var folders = parent.getFoldersByName(folderName);
  return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
}

function findFileInFolder(folder, fileName) {
  var files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    var file = files.next();
    Logger.log("File already exists: " + file.getName());
    return file;
  }
  return null;
}

function copyFile(fileUrl, newFileName, destinationFolder) {
  try {
    var fileId = extractDriveFileId(fileUrl);
    if (!fileId) {
      Logger.log("Invalid file URL: " + fileUrl);
      return null;
    }
    
    var file = DriveApp.getFileById(fileId);
    var copiedFile = file.makeCopy(newFileName, destinationFolder);
    Logger.log("Copied file: " + newFileName);
    return copiedFile;
  } catch (e) {
    Logger.log("Error processing file: " + fileUrl + " - " + e.toString());
    return null;
  }
}

function extractDriveFileId(url) {
  var match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}
