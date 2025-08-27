# Outlook Last-7-Days Mail Export → Excel (Microsoft Graph + MSAL)

Export Outlook emails (from a chosen folder) for the last _N_ days that match your keywords,
extract dates mentioned **inside the email body**, and write everything to **Excel**.

---

## What you get
- Microsoft Graph **device code flow** (no secrets needed)
- Query **Inbox or any subfolder** (e.g., `Inbox/Invoices/2025`)
- Keyword search on **from/subject/body** (Graph `$search`)
- Extract **dates** from message body (natural-language + regex)
- Output to **Excel**: `Subject | Received (IST) | Extracted Dates | From | Link`

---

## 0) Prerequisites
- Python 3.9+ (tested with 3.10+)
- An Outlook/Microsoft 365 account (Work/School or Outlook.com)
- Internet access

---

## 1) Azure App Registration (one-time)
1. Go to **Entra admin center** → *App registrations* → **New registration**.
2. Name it (e.g., `Outlook-Mail-Export-POC`).
3. Supported account types: choose based on your needs (e.g., *Accounts in any organizational directory and personal Microsoft accounts*).
4. **Authentication** → Add a platform → **Mobile and desktop applications**.
   - Add Redirect URI: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Enable: **Allow public client flows** (a toggle called “Enable the following mobile and desktop flows”).
5. **API permissions** → Add a permission → **Microsoft Graph** (Delegated):
   - `Mail.Read`
   - `offline_access` is implied when using delegated scopes (token refresh), but we’ll request it explicitly.
   - (Optional for OneDrive/Excel writes via Graph: `Files.ReadWrite` — not required here since we write locally.)
6. Save your **Application (client) ID**. Tenant can be your directory ID or `"common"`.

> You don’t need a client secret for device-code flow.

---

## 2) Configure the project
Copy `.env.example` to `.env` and fill:
```
CLIENT_ID=your_client_id_here
TENANT=common
FOLDER_PATH=Inbox
KEYWORDS=invoice,PO
DAYS=7
OUTPUT_XLSX=outlook_last7days.xlsx
```

- `FOLDER_PATH` supports nested paths like `Inbox/Invoices/2025`.
- `KEYWORDS` is comma-separated. Leave empty to skip keyword search.
- `DAYS` default 7.
- `TENANT`: use your tenant ID if you want to restrict; `common` works for most.

---

## 3) Install & run
```bash
# inside this folder
python -m venv .venv
# Windows PowerShell: .venv\Scripts\Activate.ps1
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt

# First run will show a device code + URL. Open the link, paste the code, finish login.
python src/fetch_outlook.py
```

Optional CLI flags override `.env`:
```bash
python src/fetch_outlook.py --folder "Inbox/Invoices" --keywords "invoice,PO" --days 7 --out report.xlsx --tenant common --client-id YOUR_ID
```

---

## 4) Output
The script writes an Excel file with columns:
- `Subject`
- `Received (IST)`
- `Extracted Dates`
- `From`
- `Link` (opens the message in Outlook on the web)

---

## 5) Notes & tips
- Keyword search uses Graph `$search` (KQL). We send header `ConsistencyLevel: eventual` only when `$search` is used.
- Date extraction uses both **natural-language parsing** and **regex** to increase recall.
- If you see HTTP 429 (throttling), the script retries with backoff automatically.
- If you change permissions or sign in with a different account, you may need to delete `msal_cache.bin`.
- IST conversion is done with `Asia/Kolkata` (fixed UTC+5:30 offset). Change in the code if needed.

---

## 6) Troubleshooting
- **Auth window doesn’t open**: Device-code flow prints a URL + code right in your terminal; open the URL manually and paste the code.
- **Insufficient privileges**: Ensure the app has **Mail.Read (Delegated)** and you consented.
- **Folder not found**: Check the exact folder path (`Inbox/Subfolder` names must match).

---

## 7) License
MIT — free to use for your college POC or production experiments.
