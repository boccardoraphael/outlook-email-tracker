# Outlook Email Tracker

Automated tracker that reads client folders from a Microsoft 365 Outlook mailbox and exports structured project data to an Excel file on OneDrive — updated every 2 hours via Windows Task Scheduler.

---

## What It Does

- Scans all client subfolders inside your Outlook Inbox (EA, Ubisoft, Locpick, Sony, Thallium, etc.)
- Filters out internal/HR/Teams/P&C emails — only project-related emails are tracked
- Parses structured subject lines to extract deadline, word count, task type, provider, and project name automatically
- Summarizes each email body in English using a local smart extractor (no API needed)
- Appends only new emails on each run (no duplicates)
- Saves everything to a color-coded Excel file on OneDrive

---

## Excel Columns

| Column | Description |
|---|---|
| Date Received | Date the email arrived |
| Client | Outlook folder name (EA, Ubisoft, etc.) |
| Project | Project name parsed from subject |
| PM Responsible | PM name extracted from email greeting |
| From | Sender name and email |
| Task Type | Translation, Proofreading, Delivery, Bug Report, etc. |
| Provider | Translator or proofreader name |
| Deadline | Parsed from subject or body |
| Word Count | Parsed from subject `[111 w]` or body |
| Reply/FW | RE / FW / blank |
| Project Code | Code in brackets e.g. `[i38]`, `[EA]` |
| Batch / File | Job ID or filename |
| Subject Topic | Clean human-readable subject |
| Summary | Auto-generated English summary of the email |

**Color coding:**
- 🔵 Blue rows — Translation, Proofreading, Delivery, Review
- 🟢 Green rows — General project communication
- 🟡 Yellow cell — Translation row missing word count

---

## Subject Patterns Supported

**Quoted delivery format:**
```
03-20-10h00 - [EA] - 4629909 Star Wars The Old [111 w]
03-20-10h00 - [Locpick] - LOC148208 Festival Batch 4 [46 w] [John - Translation]
```

**Native Prime format:**
```
03-19-13h00 - [Native Prime | Hawaii West (Kiln)] - filename [Translator - Translation]
```

**Generic project codes:**
```
[i38] B1 timing discrepancies
RE: [Thallium] TL_batch_33 - HO
```

---

## Setup

### Requirements

- Windows 10/11
- Microsoft Outlook desktop app (logged into your M365 account)
- Python 3.x

### Install dependencies

```bash
py -m pip install pywin32 openpyxl langdetect
```

### Configure

Open `email_tracker.py` and update these two lines at the top:

```python
EXCEL_PATH = Path(r"C:\Users\YOU\OneDrive - Your Company\Documents\Client Email Tracker.xlsx")
RAPHAEL_EMAIL = "your.email@yourcompany.com"
```

Also update the `QUOTED_PMS` set with your team's PM names.

### Run manually

```bash
py email_tracker.py
```

Or double-click `run_email_tracker.bat`.

---

## Scheduling (Windows Task Scheduler)

To run every 2 hours automatically, open PowerShell and run:

```powershell
$action   = New-ScheduledTaskAction -Execute "C:\path\to\run_email_tracker.bat"
$trigger  = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Hours 2) -Once -At (Get-Date)
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10) -StartWhenAvailable
Register-ScheduledTask -TaskName "EmailTracker" -Action $action -Trigger $trigger -Settings $settings -RunLevel Limited -Force
```

---

## Notes

- Outlook must be open and logged in for the script to run
- The Excel file must be closed when the script runs, otherwise the save will fail
- Epic Smartling folder is excluded by default (high volume, not project-specific)
- The Inbox itself is skipped — only its subfolders are scanned
