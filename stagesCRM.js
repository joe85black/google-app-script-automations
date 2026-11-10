function stage1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const c1 = sheet.getRange("C1");

  // If C1 is already green, clear formatting on second click
  if (c1.getBackground() === "#00ff00") {
    c1.setBackground(null);
    return;
  }

  const name = sheet.getRange("A2").getValue();
  const email = sheet.getRange(2, 2).getValue(); // row 2, column B
  const drafts = GmailApp.getDrafts();

  // Map draft subjects (Gmail drafts) to email subject lines (recipients will see these)
  const emailTemplates = [
    {
      draftSubject: "Stage 1 Template",
      sendSubject: "[Company Name] - Welcome!"
    },
    {
      draftSubject: "Stage 1 Template 2",
      sendSubject: "[Company Name] - Starter: Getting Started"
    },
    {
      draftSubject: "Stage 1 Template 3",
      sendSubject: "[Company Name] - Client Intake Form"
    }
  ];

  // Loop through each template and send the email
  emailTemplates.forEach(template => {
    const draft = drafts.find(d => d.getMessage().getSubject() === template.draftSubject);
    if (draft) {
      let body = draft.getMessage().getBody();
      body = body.replace("{{name}}", name);
      GmailApp.sendEmail(email, template.sendSubject, "", { htmlBody: body });
    } else {
      Logger.log("Draft not found: " + template.draftSubject);
    }
  });

  // Turn C1 green after sending
  c1.setBackground("#00ff00");

// Get the first available task list
const taskLists = Tasks.Tasklists.list().items;
if (!taskLists || taskLists.length === 0) {
  Logger.log("No task lists available.");
  return;
}

const taskListId = taskLists[0].id; // Use the first available list

// Set due date 7 days from now
const dueDate = new Date();
dueDate.setDate(dueDate.getDate() + 7);

Tasks.Tasks.insert({
  title: `Create WhatsApp contact and welcome ${name}`,
  notes: `Create contact for ${name} (${email}) in WhatsApp and send client the welcome message.`,
  due: dueDate.toISOString()
}, taskListId);


// Open the target spreadsheet and sheet
const targetSpreadsheet = SpreadsheetApp.openById("YOUR_TRACKING_SPREADSHEET_ID");
const targetSheet = targetSpreadsheet.getSheetByName("2025 - Tasks");

// Find last filled row in column A, even if there are gaps
const colA = targetSheet.getRange("A:A").getValues();
let lastRow = 0;
for (let i = colA.length - 1; i >= 0; i--) {
  if (colA[i][0] !== "") {
    lastRow = i + 1;
    break;
  }
}
const nextRow = lastRow + 1;

// Write data to columns A–C
targetSheet.getRange(nextRow, 1).setValue(name);
targetSheet.getRange(nextRow, 2).setValue("Starter Full");
targetSheet.getRange(nextRow, 3).setValue("Create Client Folder and Send Intro Emails");


}


function stage2() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const d1 = sheet.getRange("D1");
  const name = sheet.getRange("A2").getValue();

  // Second click clears formatting only
  if (d1.getBackground() === "#00ff00") {
    d1.setBackground(null);
    return;
}

  // Turn D1 green
  d1.setBackground("#00ff00");

  // Get the first available task list
const taskLists = Tasks.Tasklists.list().items;
if (!taskLists || taskLists.length === 0) {
  Logger.log("No task lists available.");
  return;
}

const taskListId = taskLists[0].id; // Use the first available list

// Set due date 7 days from now
const dueDate = new Date();
dueDate.setDate(dueDate.getDate() + 7);

Tasks.Tasks.insert({
  title: `Wait for Documents from ${name}`,
  notes: `Keep in contact with ${name} , give help when needed`,
  due: dueDate.toISOString()
}, taskListId);
}
