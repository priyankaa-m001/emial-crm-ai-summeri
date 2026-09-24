/**
 * EMAIL CRM SCRIPT
 * Paste this into: Google Sheet -> Extensions -> Apps Script
 * Reads directly from your own Gmail (no BCC / logging inbox needed).
 *
 * SETUP:
 * 1. Replace GROQ_API_KEY below with your free key from console.groq.com
 * 2. Run testLookup() once to authorize Gmail + Sheets access
 * 3. Reload the Sheet -> a new menu "Email CRM" appears at the top
 * 4. Use menu "Look up specific email" any time you want one person's summary shown as a popup
 * 5. (Optional) Set up a time-based trigger on autoScan() to summarize ALL recent senders automatically
 */

const GROQ_API_KEY = "YOUR_GROQ_API_KEY_HERE"; // get free key at console.groq.com
const SHEET_NAME = "Contacts"; // tab name in your Google Sheet

// ---------- Adds a menu to the Sheet so you can trigger a lookup by typing an email ----------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Email CRM")
    .addItem("Look up specific email", "lookupSpecificEmail")
    .addItem("Scan all recent emails", "autoScan")
    .addToUi();
}

// ---------- User types ONE email address -> only that person's summary is shown + saved ----------
function lookupSpecificEmail() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.prompt(
    "Look up an email",
    "Enter the email address whose conversation summary you want:",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return; // user cancelled

  const contactEmail = result.getResponseText().trim();
  if (!contactEmail) {
    ui.alert("No email entered.");
    return;
  }

  const summary = processContact(contactEmail); // filters to ONLY this email's threads
  if (summary) {
    ui.alert("Summary for " + contactEmail, summary, ui.ButtonSet.OK);
  } else {
    ui.alert("No emails found for " + contactEmail);
  }
}

// ---------- Manual test from the script editor (Run button) ----------
function testLookup() {
  const contactEmail = "clientemail@example.com"; // <-- change this to test
  processContact(contactEmail);
}

// Looks up ONLY the Gmail threads involving this one email address
// (both sent-to and received-from), summarizes them, saves to the Sheet,
// and returns the summary text so callers (like the popup) can show it.
function processContact(contactEmail) {
  const threads = GmailApp.search(
    'from:"' + contactEmail + '" OR to:"' + contactEmail + '"'
  );

  if (threads.length === 0) {
    Logger.log("No emails found for " + contactEmail);
    return null;
  }

  // Collect message text from matching threads for THIS sender only.
  // Limit to the most recent 5 threads and trim each email body, so the
  // request stays under Groq's free-tier size limit.
  const recentThreads = threads.slice(0, 5);
  let allText = "";
  let contactName = "";
  let latestSubject = "";
  let latestDate = null;

  recentThreads.forEach(function (thread) {
    if (!latestSubject) latestSubject = thread.getFirstMessageSubject(); // most recent thread's subject
    const messages = thread.getMessages().slice(-3); // last 3 messages per thread
    messages.forEach(function (msg) {
      // Pick up the contact's display name from whichever message they sent
      if (!contactName && msg.getFrom().indexOf(contactEmail) !== -1) {
        contactName = extractName(msg.getFrom());
      }
      // Track the most recent message date across everything we looked at
      if (!latestDate || msg.getDate() > latestDate) {
        latestDate = msg.getDate();
      }
      allText +=
        "Date: " + msg.getDate() +
        "\nSubject: " + msg.getSubject() +
        "\nBody: " + msg.getPlainBody().substring(0, 400) + // trim long emails
        "\n---\n";
    });
  });

  // Hard cap on total combined text sent to the AI
  if (allText.length > 6000) {
    allText = allText.substring(0, 6000);
  }

  const summary = summarizeWithGroq(allText, contactEmail);
  writeToSheet(contactName || contactEmail, contactEmail, latestSubject, summary, latestDate);
  return summary;
}

// Helper: extracts just the display name from a "Name <email>" string (falls back to the email itself)
function extractName(fromField) {
  const match = fromField.match(/^(.*?)</);
  return match ? match[1].trim().replace(/"/g, "") : fromField;
}

// ---------- Calls Groq API (free) to summarize the email history ----------
function summarizeWithGroq(emailText, contactEmail) {
  const prompt =
    "Here are all emails exchanged with " + contactEmail + ":\n\n" +
    emailText +
    "\n\nSummarize ONLY the important points from this conversation, in 2-4 sentences: " +
    "the key topics/requests/decisions, in chronological order, and the current status or next step. " +
    "Ignore greetings, sign-offs, pleasantries, and any other unimportant details.";

  const response = UrlFetchApp.fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "post",
    contentType: "application/json",
    headers: {
      "Authorization": "Bearer " + GROQ_API_KEY
    },
    payload: JSON.stringify({
      model: "openai/gpt-oss-120b",
      max_tokens: 300,
      messages: [{ role: "user", content: prompt }]
    })
  });

  const data = JSON.parse(response.getContentText());
  return data.choices[0].message.content;
}

// ---------- Writes or updates a row in the Sheet ----------
// Columns: Name, Email, Thread (latest subject), Summary, Last Conversation Date
function writeToSheet(contactName, contactEmail, thread, summary, lastDate) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME)
    || SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);

  // Add headers if sheet is empty
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Name", "Email", "Thread", "Summary", "Last Conversation Date"]);
  }

  // Check if this contact already has a row (matched by email) -> update it, else add new row
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === contactEmail) {
      rowIndex = i + 1; // sheet rows are 1-indexed
      break;
    }
  }

  const rowValues = [contactName, contactEmail, thread, summary, lastDate];

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, 5).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

// ---------- Full scan: summarizes EVERY sender you've emailed with recently ----------
// Run from the menu any time, or set up under Triggers -> Add Trigger -> autoScan -> Time-driven -> every 15-30 min
function autoScan() {
  // Scans your real inbox (sent + received) for recent threads and re-summarizes each sender
  const threads = GmailApp.search("newer_than:1d"); // adjust window as needed
  const seenSenders = {};

  threads.forEach(function (thread) {
    const messages = thread.getMessages();
    messages.forEach(function (msg) {
      const sender = msg.getFrom();
      if (!seenSenders[sender]) {
        seenSenders[sender] = true;
        processContact(extractEmail(sender));
      }
    });
  });
}

// Helper: extracts just the email address from a "Name <email>" string
function extractEmail(fromField) {
  const match = fromField.match(/<(.+)>/);
  return match ? match[1] : fromField;
}
