# Email CRM — AI Conversation Summarizer

A lightweight Google Sheets tool that reads your Gmail conversation history with any contact and generates an AI summary on demand — no manual logging, no third-party CRM subscription.

## How it works

1. Open the connected Google Sheet and use the **Email CRM** menu.
2. Type in an email address to look up.
3. The script (Google Apps Script) searches your Gmail for every thread involving that address.
4. The collected conversation text is sent to an LLM (via the Groq API) for summarization.
5. The summary is shown instantly as a popup, and saved as a row in the Sheet: **Name, Email, Thread, Summary, Last Conversation Date**.

## Tech stack

- **Backend:** Google Apps Script (JavaScript), Gmail API (via `GmailApp`), Google Sheets API (via `SpreadsheetApp`)
- **AI:** Groq API (`openai/gpt-oss-120b`)

## Setup

1. Create a new Google Sheet.
2. Open **Extensions → Apps Script**.
3. Paste `email-crm-script.gs` into `Code.gs`.
4. Add a free API key from [console.groq.com](https://console.groq.com) into the `GROQ_API_KEY` constant.
5. Run `testLookup()` once to authorize Gmail and Sheets access.
6. Reload the Sheet — a new **Email CRM** menu appears at the top.
7. Use **"Look up specific email"** to summarize one contact, or **"Scan all recent emails"** to summarize everyone you've emailed recently.

## Notes

- Reads only the Gmail account it's authorized under — no external inbox access, no BCC logging trick, no third-party CRM.
- Uses Groq's free tier for summarization, which has request-size and rate limits suitable for personal use.

## Author

Priyanka Mhaske — [GitHub](https://github.com/priyankaa-m001) · [LinkedIn](https://www.linkedin.com/in/priyankamhaske)
