function processAdmitCards() {
    var SHEET_NAME = "MAIN_SHEET"; // <-- Change this if needed
  
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    
    if (!sheet) {
      Logger.log("Sheet not found: " + SHEET_NAME);
      return;
    }
  
    var data = sheet.getDataRange().getValues();
    var header = data[0];
  
    var studentCodeIndex = header.indexOf("student_code");
    var nameIndex = header.indexOf("name");
    var photoIndex = header.indexOf("photo");
    var signIndex = header.indexOf("sign");
    var newPhotoIndex = header.indexOf("new_photo_link");
    var newSignIndex = header.indexOf("new_sign_link");
  
    if (
      studentCodeIndex === -1 || nameIndex === -1 ||
      photoIndex === -1 || signIndex === -1 ||
      newPhotoIndex === -1 || newSignIndex === -1
    ) {
      Logger.log("One or more required columns are missing.");
      return;
    }
  
    var parentFolder = DriveApp.getRootFolder(); // Change this if needed
    var mainFolder = getOrCreateFolder(parentFolder, "Offline Exam - Student Documents");
  
    Logger.log(`All files will be stored in: ${mainFolder.getUrl()}`);
  
    var totalRows = data.length - 1; // Exclude header row
    var processedRows = 0;
  
    for (var i = 1; i < data.length; i++) {
      var studentCode = data[i][studentCodeIndex];
      var studentName = data[i][nameIndex];
      var photoUrl = data[i][photoIndex];
      var signUrl = data[i][signIndex];
      var existingPhotoLink = data[i][newPhotoIndex];
      var existingSignLink = data[i][newSignIndex];
  
      // Skip if both files are already processed
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
            sheet.getRange(i + 1, newPhotoIndex + 1).setValue(existingPhoto.getUrl());
          } else {
            var copiedPhoto = copyFile(photoUrl, newPhotoName, mainFolder);
            if (copiedPhoto) {
              sheet.getRange(i + 1, newPhotoIndex + 1).setValue(copiedPhoto.getUrl());
              Utilities.sleep(2000); // Wait to ensure Google Drive updates
            }
          }
        }
  
        // Process Signature
        if (signUrl && !existingSignLink) {
          var existingSign = findFileInFolder(mainFolder, newSignName);
          if (existingSign) {
            sheet.getRange(i + 1, newSignIndex + 1).setValue(existingSign.getUrl());
          } else {
            var copiedSign = copyFile(signUrl, newSignName, mainFolder);
            if (copiedSign) {
              sheet.getRange(i + 1, newSignIndex + 1).setValue(copiedSign.getUrl());
              Utilities.sleep(2000); // Wait to ensure Google Drive updates
            }
          }
        }
  
        processedRows++;
        Logger.log(`${processedRows}/${totalRows} processed`);
      }
    }
  
    Logger.log(`Processing complete. Total rows processed: ${processedRows}/${totalRows}`);
  }
  
  // Function to create or retrieve a folder by name
  function getOrCreateFolder(parent, folderName) {
    var folders = parent.getFoldersByName(folderName);
    return folders.hasNext() ? folders.next() : parent.createFolder(folderName);
  }
  
  // Function to check if a file already exists in the folder
  function findFileInFolder(folder, fileName) {
    var files = folder.getFilesByName(fileName);
    while (files.hasNext()) {
      var file = files.next();
      Logger.log("File already exists: " + file.getName());
      return file;
    }
    return null;
  }
  
  // Function to copy a file from a given Drive link
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
  
  // Extracts the Google Drive file ID from the URL
  function extractDriveFileId(url) {
    var match = url.match(/[-\w]{25,}/);
    return match ? match[0] : null;
  }