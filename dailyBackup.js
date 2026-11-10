/**
 * Copies a fixed list of source Google Sheets into a rotating pair of
 * day-of-week subfolders under a parent Drive folder, so each spreadsheet
 * has a same-day-last-week and current-day backup on hand at all times.
 *
 * Folder layout under PARENT_FOLDER_ID: "1_1Sunday", "1_2Monday", ...,
 * "2_1Sunday", "2_7Saturday" — the "1_"/"2_" prefix is a week slot that
 * flips every 7 days from ROTATION_ANCHOR_ISO, so writing to the same
 * weekday folder twice in a row never happens back-to-back.
 *
 * All IDs below are placeholders — set PARENT_FOLDER_ID and
 * SOURCE_FILE_IDS to your own Drive folder / spreadsheet IDs before use.
 */

/***** CONFIG *****/

// Folder where all the day-of-week subfolders live
// Example: "1_1Sunday", "2_3Tuesday", etc.
const PARENT_FOLDER_ID = 'YOUR_BACKUP_FOLDER_ID'; // Backup root folder

// Explicit list of Sheets to copy (only these will be backed up)
const SOURCE_FILE_IDS = [
  'YOUR_SPREADSHEET_ID_1', // Contracts & Calls Log
  'YOUR_SPREADSHEET_ID_2', // Completed Projects (Automated)
  'YOUR_SPREADSHEET_ID_3', // Client Scheduling
  'YOUR_SPREADSHEET_ID_4', // Confirmation Notices (Automated)
  'YOUR_SPREADSHEET_ID_5', // Follow-Up: Business Plus, Annual Review, and Quick Check Clients
  'YOUR_SPREADSHEET_ID_6', // Follow-Up: Welcome Form Clients
];

// Leave blank unless you want to copy *all* Sheets from a folder instead of the explicit list above.
const SOURCE_FOLDER_ID = '';

// Rotation anchor (a Sunday). Week slot flips every 7 days from here: 1 ↔ 2.
const ROTATION_ANCHOR_ISO = '2025-08-17'; // YYYY-MM-DD

// Behavior toggles
const CLEAN_TARGET_BEFORE_COPY = true;   // trash old files in target folder before copying

// Timezone for stamping backups
const TIMEZONE = 'America/New_York'; // set to your own timezone


/***** MAIN BACKUP FUNCTION *****/
function runBackup() {
  // Prevent overlapping runs of the script
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another backup is already running. Try again later.');

  try {
    const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
    const today = new Date();

    // Figure out which "week slot" we're in (1 or 2) based on anchor date
    const weekSlot = getWeekSlot(today, ROTATION_ANCHOR_ISO, TIMEZONE);

    // Day of week (Sunday=0 … Saturday=6)
    const dayIdx = today.getDay();
    const dayNum = dayIdx + 1; // shift to 1..7
    const dayName = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'][dayIdx];

    // Match your folder structure exactly: e.g. "1_1Sunday", "2_7Saturday"
    const targetFolderName = `${weekSlot}_${dayNum}${dayName}`;
    const targetFolder = getOrCreateSubfolder(parent, targetFolderName);

    // Optionally clear out old files before copying new ones
    if (CLEAN_TARGET_BEFORE_COPY) trashAllFilesInFolder(targetFolder);

    // Get list of source spreadsheets
    const filesToCopy = getSourceFiles();

    // Format date as MM/dd for naming backups
    const displayStamp = Utilities.formatDate(today, TIMEZONE, 'MM/dd');

    // Copy each source file into the target folder with backup prefix
    while (filesToCopy.hasNext()) {
      const f = filesToCopy.next();
      const baseName = f.getName();

      // Prepend "Back-Up MM/DD -" before the original name
      const newName = `Back-Up ${displayStamp} - ${baseName}`;

      f.makeCopy(newName, targetFolder);
    }

    Logger.log(`Backed up ${targetFolderName}`);
  } finally {
    lock.releaseLock();
  }
}


/***** HELPER FUNCTIONS *****/

// Calculates which week slot we're in (1 or 2) based on an anchor Sunday
function getWeekSlot(dateObj, anchorISO, tz) {
  const midnight = new Date(Utilities.formatDate(dateObj, tz, 'yyyy-MM-dd') + 'T00:00:00');
  const anchor = new Date(anchorISO + 'T00:00:00');
  const days = Math.floor((midnight.getTime() - anchor.getTime()) / 86400000);
  const weekIndex = Math.floor(days / 7);
  // Normalize parity for negatives too
  const parity = ((weekIndex % 2) + 2) % 2; // 0 even, 1 odd
  return parity === 0 ? 1 : 2;
}

// Get or create a subfolder under the parent
function getOrCreateSubfolder(parent, name) {
  const it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

// Trash all files in a given folder
function trashAllFilesInFolder(folder) {
  const it = folder.getFiles();
  while (it.hasNext()) it.next().setTrashed(true);
}

// Gather source spreadsheet files
function getSourceFiles() {
  // Prefer explicit list
  if (SOURCE_FILE_IDS && SOURCE_FILE_IDS.length > 0) {
    const files = SOURCE_FILE_IDS
      .map(id => DriveApp.getFileById(id))
      .filter(f => f.getMimeType() === MimeType.GOOGLE_SHEETS);

    return iteratorFromArray_(files);
  }

  // Or take all Sheets from a folder
  if (SOURCE_FOLDER_ID) {
    const srcFolder = DriveApp.getFolderById(SOURCE_FOLDER_ID);
    const all = srcFolder.getFiles();
    const keep = [];
    while (all.hasNext()) {
      const f = all.next();
      if (f.getMimeType() === MimeType.GOOGLE_SHEETS) keep.push(f);
    }
    return iteratorFromArray_(keep);
  }

  throw new Error('No sources defined. Add SOURCE_FILE_IDS or SOURCE_FOLDER_ID.');
}

// Turns an array into an iterator (so it matches Apps Script style)
function iteratorFromArray_(arr) {
  return {
    _i: 0,
    hasNext() { return this._i < arr.length; },
    next() { return arr[this._i++]; }
  };
}


/***** OPTIONAL QUALITY OF LIFE HELPERS (disabled by default) *****/

// Creates a daily time-based trigger (~02:00) to run backups automatically
// function createDailyTrigger() {
//   deleteExistingTriggers_('runBackup');
//   ScriptApp.newTrigger('runBackup').timeBased().atHour(2).everyDays(1).create();
// }

// Deletes the daily trigger
// function deleteDailyTrigger() { deleteExistingTriggers_('runBackup'); }

// Deletes existing triggers for a given function
// function deleteExistingTriggers_(funcName) {
//   ScriptApp.getProjectTriggers().forEach(t => {
//     if (t.getHandlerFunction() === funcName) ScriptApp.deleteTrigger(t);
//   });
// }

// Optional helper to pre-create the 14 subfolders up front
// function createAllSubfolders() {
//   const parent = DriveApp.getFolderById(PARENT_FOLDER_ID);
//   const days = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
//   [1,2].forEach(week => {
//     days.forEach((d,i) => getOrCreateSubfolder(parent, `${week}_${i+1}${d}`));
//   });
// }
