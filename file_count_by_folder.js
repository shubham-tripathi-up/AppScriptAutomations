function countFilesInFolder() {
    var folderId = "<FOLDER ID HERE"; // Replace with your actual folder ID
    var folder = DriveApp.getFolderById(folderId);
    var files = folder.getFiles(); // Get all files in the folder
    
    var count = 0;
    while (files.hasNext()) {
      files.next(); // Move to the next file
      count++;
    }
    
    Logger.log("Total number of files in the folder: " + count);
  }