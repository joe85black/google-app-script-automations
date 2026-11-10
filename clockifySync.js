/**
 * Pulls detailed Clockify time entries for one user and writes them into
 * monthly tabs of a Google Sheet, grouped by project/client.
 *
 * Each month tab (Jun–Dec below) gets its existing rows in A63:E200 cleared
 * and rewritten with that month's entries: Project, Client, Description,
 * Time (h:mm), and Time (decimal hours) — sorted by project name.
 *
 * Setup:
 *   1. In the Script Editor, go to Project Settings > Script Properties and
 *      add a CLOCKIFY_API_KEY property with your Clockify API key
 *      (found in Clockify under Profile Settings > API).
 *   2. Set USER_ID, WORKSPACE_ID, and SHEET_ID below to your own values.
 *      USER_ID/WORKSPACE_ID come from Clockify's API/URL; SHEET_ID is the
 *      destination spreadsheet's ID (from its URL).
 *
 * The API key is intentionally kept out of source (Script Properties
 * instead of a hardcoded constant) since this file may end up in a shared
 * or public repo.
 */
const CLOCKIFY_API_KEY = PropertiesService.getScriptProperties().getProperty("CLOCKIFY_API_KEY");
const USER_ID = "YOUR_CLOCKIFY_USER_ID";
const WORKSPACE_ID = "YOUR_CLOCKIFY_WORKSPACE_ID";
const SHEET_ID = "YOUR_SPREADSHEET_ID";

const MONTH_SHEETS = ["Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function fetchAndWriteDetailedEntries() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const currentYear = new Date().getFullYear();

  // Step 1: Build a map of projectId → { projectName, clientName }
  const projectMap = getProjectClientMap();

  // Step 2: Loop through months and sheets
  MONTH_SHEETS.forEach((sheetName, index) => {
    const monthIndex = 5 + index; // June = 5

    const startDate = new Date(currentYear, monthIndex, 1);
    const endDate = new Date(currentYear, monthIndex + 1, 0, 23, 59, 59);

    const isoStart = startDate.toISOString();
    const isoEnd = endDate.toISOString();

    const url = `https://api.clockify.me/api/v1/workspaces/${WORKSPACE_ID}/user/${USER_ID}/time-entries?start=${isoStart}&end=${isoEnd}&page-size=5000`;

    const options = {
      method: "get",
      headers: {
        "X-Api-Key": CLOCKIFY_API_KEY
      },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(url, options);
    const content = response.getContentText();

    let entries;
    try {
      entries = JSON.parse(content);
      if (!Array.isArray(entries)) throw new Error("Invalid structure");
    } catch (e) {
      Logger.log("Error parsing response: " + content);
      return;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Clear old data and write headers
    sheet.getRange("A63:E200").clearContent();
    sheet.getRange("A63:E63").setValues([["Project", "Client", "Description", "Time (h:mm)", "Time (decimal)"]]);

    // Build rows from time entries
    const rows = entries.map(entry => {
      if (!entry.timeInterval?.start || !entry.timeInterval?.end) return null;

      const projectId = entry.projectId;
      const project = projectMap[projectId]?.projectName || "No Project";
      const client = projectMap[projectId]?.clientName || "No Client";
      const desc = entry.description || "";

      const start = new Date(entry.timeInterval.start);
      const end = new Date(entry.timeInterval.end);
      const seconds = (end - start) / 1000;
      const decimalHours = (seconds / 3600).toFixed(4);

      const hours = Math.floor(seconds / 3600);
      const minutes = Math.floor((seconds % 3600) / 60);
      const secondsRemain = Math.floor(seconds % 60);
      const timeFormatted = `${hours}:${minutes.toString().padStart(2, '0')}:${secondsRemain.toString().padStart(2, '0')}`;

      return [project, client, desc, timeFormatted, parseFloat(decimalHours)];
    }).filter(row => row !== null);

    // Sort rows by Project name
    rows.sort((a, b) => a[0].localeCompare(b[0]));

    // Write to sheet starting at row 64
    if (rows.length > 0) {
      sheet.getRange(64, 1, rows.length, 5).setValues(rows);
    }
  });
}

// Helper: Get map of projectId → { projectName, clientName }
function getProjectClientMap() {
  const url = `https://api.clockify.me/api/v1/workspaces/${WORKSPACE_ID}/projects`;
  const options = {
    method: "get",
    headers: {
      "X-Api-Key": CLOCKIFY_API_KEY
    }
  };

  const response = UrlFetchApp.fetch(url, options);
  const projects = JSON.parse(response.getContentText());

  const map = {};
  projects.forEach(p => {
    map[p.id] = {
      projectName: p.name || "Unnamed Project",
      clientName: p.clientName || "No Client"
    };
  });

  return map;
}
