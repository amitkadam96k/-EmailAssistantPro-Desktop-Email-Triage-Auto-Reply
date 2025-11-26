# 🚀 EmailAssistantPro — Desktop Email Triage & Auto-Reply

A compact, friendly desktop utility that helps you fetch emails, categorize incoming messages, auto-reply with templates, save attachments, and generate a PDF summary from a CSV log. Perfect for small teams or solo professionals who want to save time responding to common requests.

✨ Features
- 📥 Fetch emails from your IMAP inbox
- 🧠 Classify emails using keyword-based categories (Billing, Order, Support, Lead, Other)
- ✉️ Auto-reply using category-specific templates over SMTP
- 💾 Save attachments to `attachments/email_<uid>/`
- 🗂️ Log replies to `logs/email_log.csv` and export a PDF summary `logs/email_log_summary.pdf`
- 🖥️ Modern desktop GUI built with `customtkinter`

🛠️ Prerequisites
- Python 3.10 or newer (type annotations use the `|` union operator)
- `tkinter` must be available in the Python installation (Windows installer often includes it)
- Valid IMAP & SMTP credentials for your email provider (for example Gmail app password if using Gmail)

📦 Dependencies
Install from the `requirements.txt` file — the key packages are:
- `customtkinter` — Modern theme wrapper for Tkinter
- `keyring` — Secure credential storage (system keychain integration)
- `fpdf2` — PDF generation library (importable as `from fpdf import FPDF`)

⚡ Quick setup (Windows PowerShell)
```powershell
cd 'C:\Users\***\OneDrive\Desktop\email'
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

📝 How to use
1. Start the app (`python app.py`).
2. Fill in your email address, app password (or allow keyring to autofill), IMAP and SMTP server and ports.
   - Example (Gmail): IMAP: `imap.gmail.com:993`, SMTP: `smtp.gmail.com:465`, Use SSL checked.
3. Click Connect (credentials are saved automatically using the system keyring).
4. Click "Fetch latest" to load emails, then use "Classify all" to apply categories.
5. Select an email in the list to view details and click "Auto-reply selected" to send the auto-reply.
6. Use "Generate PDF Log Summary" to create a PDF summary from the CSV logs.

💡 Notes & configuration
- 🔐 Keyring: On Windows, `keyring` typically uses the Windows Credential Manager — passwords are stored securely by the system.
- 📎 Attachments: Saved under `attachments/email_<uid>/` where `uid` is the email id (unique per fetch).
- 🧾 Logs: `logs/email_log.csv` is appended automatically — the CSV header is created if the file doesn't exist.
- 🧰 Classifier: A simple rule-based keyword classifier is used by default — you can replace `classify_email` with any other logic or model.

⚠️ Troubleshooting
- `tkinter` missing: Reinstall Python and ensure Tcl/Tk is installed or install the OS package that provides it.
- Gmail connection errors: Create an App Password and use `imap.gmail.com:993` and `smtp.gmail.com:465` if using Gmail.
- `keyring` errors: Consult the `python-keyring` documentation for OS backend settings and debugging tips.

📁 Project structure
```
app.py
attachments/          # saved attachments per email subfolder (email_<uid>)
logs/
  email_log.csv        # CSV log containing replies
  email_log_summary.pdf
README.md
requirements.txt
```

🤝 Contributing
- Pull requests are welcome — suggested improvements include better classification rules, error handling, automated tests, and an installer/packaging setup.


