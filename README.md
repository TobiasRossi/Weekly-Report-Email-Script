# 📊 Weekly Report Email Script

A Google Apps Script that automatically sends weekly personalized HTML emails to team members with their GPH Data Monitoring summary, including CC to their respective Team Leads.

---

## 🚀 Overview

This script reads performance data from a Google Spreadsheet and sends each team member a formatted email summarizing their monitored DTPs (Data Tracking Points), including remaining usage and variance indicators. Rows with variance outside the ±10% threshold are visually highlighted in red.

---

## 🛠️ Built With

- Google Apps Script
- Google Sheets API (`SpreadsheetApp`)
- Gmail API (`MailApp`)
- HTML email templating

---

## 📋 How It Works

1. Reads metric data from the **`Summary`** sheet
2. Reads user-to-Team Lead mappings from the **`Resilience`** sheet
3. Groups all rows by username (supports multiple users per row via `/` separator)
4. Sends one email per user with a full HTML table of their metrics
5. Automatically CCs the user's Team Lead if found in the mapping

---

## 📁 Sheet Structure

### `Summary` sheet
| Column | Index | Description |
|--------|-------|-------------|
| A | 0 | DTP Name |
| B | 1 | Frequency |
| C | 2 | Current value |
| D | 3 | Expected value |
| E | 4 | Variance (decimal, e.g. `0.08`) |
| K | 10 | Username(s) — separated by `/` for multiple |

### `Resilience` sheet
| Column | Index | Description |
|--------|-------|-------------|
| O | 14 | Username |
| Q | 16 | Team Lead username |

> Both sheets must exist in the same spreadsheet. The script will throw an error if either is missing.

---

## ⚙️ Setup

1. Open your Google Spreadsheet
2. Go to **Extensions > Apps Script**
3. Paste the script and save
4. Run `sendWeeklyEmailsToPerformersGrouped` manually or set up a **time-based trigger** for weekly execution

---

## 📬 Email Output

Each recipient gets an email like:

**Subject:** `Weekly GPH Data Monitoring Summary - username (yyyy-MM-dd)`

**Body:** An HTML table with the following columns:
- **DTP Name**
- **Frequency**
- **Remaining Use** *(calculated as Current − Expected)*
- **Variance** *(color-coded: 🟢 green if within ±10%, 🔴 red if outside)*

---

## ✅ Validations & Error Handling

- Skips rows where values are missing, non-numeric, or empty strings
- Wraps each email send in a `try/catch` — a failure on one user won't stop the rest
- Logs all sent emails and errors to the Apps Script Logger
- Supports `DRY_RUN` mode to preview output without sending real emails

---

## 📌 Notes

- All email addresses are assumed to follow the `@google.com` domain format
- `MailApp` has a daily sending limit (~100 for personal accounts, ~1500 for Workspace)
- Script timezone follows `Session.getScriptTimeZone()`