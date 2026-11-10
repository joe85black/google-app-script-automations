/**
 * Prefixes every Google Sheets file inside each subfolder of a parent Drive
 * folder with "Back-Up " (skipping files already prefixed).
 *
 * Meant to run after a backup routine (see runBackup in dailyBackup.js) has
 * copied spreadsheets into per-day subfolders, so the copies are clearly
 * distinguishable from the originals if the whole folder tree is browsed.
 *
 * FOLDER_ID is a placeholder — set it to the Drive folder ID of your own
 * backup parent folder before running.
 */
function renameSpreadsheets() {
  var FOLDER_ID = "YOUR_BACKUP_FOLDER_ID"; // Parent Drive folder ID (backup root)
  var parentFolder = DriveApp.getFolderById(FOLDER_ID);

  // Get all subfolders inside the parent
  var subFolders = parentFolder.getFolders();

  while (subFolders.hasNext()) {
    var subFolder = subFolders.next();

    // Get all spreadsheet files in the subfolder
    var files = subFolder.getFilesByType(MimeType.GOOGLE_SHEETS);

    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();

      // Check if name already starts with "Back-Up"
      if (!name.startsWith("Back-Up")) {
        var newName = "Back-Up " + name;
        file.setName(newName);
        Logger.log("Renamed: " + name + " → " + newName);
      } else {
        Logger.log("Skipped (already renamed): " + name);
      }
    }
  }
}
