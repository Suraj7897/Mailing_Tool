📧 Outlook Last-7-Days Mail Export → Excel

Python tool using Microsoft Graph + MSAL

Export Outlook emails (from any folder) for the last N days, filter by keywords, extract dates from the body, and save everything to Excel.

✨ Features

🔑 Microsoft Graph device code flow → No client secrets required.

📂 Query any folder → e.g., Inbox, Inbox/Invoices/2025.

🔍 Keyword search (from/subject/body via Graph $search).

📅 Extract dates inside the email body (NLP + regex).

📊 Excel output with:

Subject

Received (IST)

Extracted Dates

From

Direct Outlook link

⚡ Prerequisites

🐍 Python 3.9+ (tested with 3.10+)

📬 Outlook / Microsoft 365 account (Work/School/Outlook.com)

🌐 Internet access



🔄 Architecture / Workflow
Mermaid Diagram (Markdown GitHub-ready)
flowchart TD
    A[👨‍💻 User runs Python script<br/>fetch_outlook.py] --> B[🔐 Microsoft Graph Auth<br/>(Device Code Flow)]
    B --> C[📬 Outlook Mailbox<br/> (Inbox / Subfolders)]
    C --> D[🔍 Apply Filters<br/>Keywords + Last N days]
    D --> E[📑 Extract Dates<br/>from Email Body (Regex + NLP)]
    E --> F[📊 Save to Excel<br/>(Subject, From, Dates, Link)]

    

🔑 1. Azure App Registration (one-time)

Go to Entra admin center → App registrations → New registration

Name it: Outlook-Mail-Export-POC

Supported account types → choose “Accounts in any directory + personal Microsoft accounts”

Authentication → Add platform → Mobile and desktop applications

Redirect URI: https://login.microsoftonline.com/common/oauth2/nativeclient

Enable ✅ Allow public client flows

API Permissions → Microsoft Graph → Delegated →

Mail.Read

offline_access (for token refresh)
(Optional: Files.ReadWrite if you want Graph-based Excel writes — not required here)

Save your Application (Client) ID + Tenant ID (or use common)

⚙️ 2. Project Configuration

Copy .env.example → .env and set:

CLIENT_ID=your_client_id_here
TENANT=common
FOLDER_PATH=Inbox
KEYWORDS=invoice,PO
DAYS=7
OUTPUT_XLSX=outlook_last7days.xlsx


FOLDER_PATH → supports nested like Inbox/Invoices/2025

KEYWORDS → comma-separated (invoice,PO)

DAYS → default 7

TENANT → use tenant ID or common

🚀 3. Install & Run
# Setup environment
python -m venv .venv
# Activate (Windows)
.venv\Scripts\Activate.ps1
# Activate (macOS/Linux)
source .venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run (first time → device code flow login)
python src/fetch_outlook.py

🔧 Optional CLI flags
python src/fetch_outlook.py \
  --folder "Inbox/Invoices" \
  --keywords "invoice,PO" \
  --days 7 \
  --out report.xlsx \
  --tenant common \
  --client-id YOUR_ID

📂 4. Output

✅ Excel file with columns:

Subject

Received (IST)

Extracted Dates

From

Link (click → opens message in Outlook Web)

💡 5. Notes & Tips

Keyword search → uses Graph $search (ConsistencyLevel: eventual).

Date extraction → NLP + regex for better accuracy.

HTTP 429 (throttling) → handled with auto-retry + backoff.

Change accounts/permissions → delete msal_cache.bin.

IST conversion uses Asia/Kolkata (UTC+5:30).

🛠 6. Troubleshooting

🔒 Auth issue → Use the device code shown in terminal → open link → paste code.

⚠️ Insufficient privileges → Ensure Mail.Read delegated is consented.

📁 Folder not found → Double-check path (case-sensitive).

📜 7. License

MIT — ✅ free to use for college POC or production experiments.