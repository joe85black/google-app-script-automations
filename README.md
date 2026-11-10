# Google App Script Automations

Google Apps Script automations built for a spreadsheet-based CRM/workflow — syncing data between sheets, sending templated notifications, backing up spreadsheets, tracking client stage progress, and pulling in time-tracking data. These are reference copies of scripts bound to Google Sheets, written for a client-services business that sells several service packages (Starter, Business Plus, Quick Check, Annual Review).

All names and data in this repository are fictional examples.

## Scripts

| File | What it does |
|---|---|
| [followUp.js](./followUp.js) | Classifies free-text follow-up notes into status labels, and syncs/dedupes client rows across multiple package tabs and external spreadsheets (matching on name + date). |
| [confirmationNotice.js](./confirmationNotice.js) | `onEdit` trigger that copies a new value from one sheet to a destination sheet in another spreadsheet, plus a name-formatting utility. |
| [stagesCRM.js](./stagesCRM.js) | Trigger functions tied to spreadsheet cell-color changes that send templated onboarding emails, create Google Tasks, and log progress to a tracking spreadsheet. |
| [dailyBackup.js](./dailyBackup.js) | Copies a fixed list of source spreadsheets into a rotating pair of day-of-week Drive subfolders, so each sheet always has a same-day-last-week and current-day backup on hand. |
| [renameSpreadsheets.js](./renameSpreadsheets.js) | Walks the backup folder tree and prefixes any un-prefixed spreadsheet copies with "Back-Up ", so backups are clearly distinguishable from originals at a glance. |
| [clockifySync.js](./clockifySync.js) | Pulls a user's detailed Clockify time entries via the Clockify API and writes them into per-month sheet tabs, grouped and sorted by project/client. |

## Usage

These are Apps Script files meant to be pasted into the Script Editor of a bound Google Sheet (Extensions > Apps Script). They reference sheet/tab names, spreadsheet/folder IDs, and API keys that need to be set to match your own environment — values such as `YOUR_SPREADSHEET_ID`, `YOUR_BACKUP_FOLDER_ID`, and `[Company Name]` mark where your own values go. See each file's header comment for its specific setup steps.

`clockifySync.js` reads its API key from Script Properties rather than a hardcoded constant (Project Settings > Script Properties > `CLOCKIFY_API_KEY`), since that file is more likely to be copied into shared or public projects.
