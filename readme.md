# 📬 Email Activity Monitor

This project automatically monitors the inbox of **procurementhub361@gmail.com** and logs any emails containing `@Question` in the subject line.

Every **15 minutes**, it checks for new messages and updates an Excel log file.

---

## ✨ Features

✅ Monitors Gmail for new emails with `@Question` in the headline  
✅ Extracts:
- Sender
- Date
- Time received
- Wait time (how long since the email arrived)
- Status (Not started, In Progress, Done)

✅ Appends entries to an Excel file: result/log.xlsx


✅ Runs automatically via **GitHub Actions**:
- Every 15 minutes (`cron` schedule)
- On every push to `main`
- Or manually triggered in the Actions tab

---

## 🛠 How It Works

1. **Gmail API** retrieves unread emails.
2. Emails are filtered by subject line.
3. Details are formatted and written to `log.xlsx`.
4. If the file doesn’t exist, it is created with headers and dropdown validation for the Status column.
5. Changes are automatically committed back to the repository.

---

## 🧩 Project Structure
.
├── main.py # Main script to fetch and log emails
├── append_to_excel.py # Utility to append rows and manage Excel formatting
├── requirements.txt # Python dependencies
├── .github
│ └── workflows
│ └── email_monitor.yml # GitHub Actions workflow
└── result
└── log.xlsx # The Excel log file (auto-generated)



---

## ⚙️ GitHub Actions Workflow

The workflow is defined in:
.github/workflows/email_monitor.yml

markdown
Copy
Edit
It performs:
- Checkout repository
- Set up Python
- Install dependencies
- Restore credentials (`credentials.json` and `token.pickle`)
- Run `main.py`
- Commit and push any updates to `log.xlsx`

---

## 📝 Requirements

- Python 3.10+
- Google API credentials (`credentials.json`) for Gmail and Sheets access
- A `token.pickle` file (created after the first OAuth login)
- GitHub repository secrets:
  - `GOOGLE_CREDENTIALS_JSON`
  - `GMAIL_TOKEN_BASE64`
  - (Optional fallback) `PAT_TOKEN` for pushing changes

---

## 🚀 How to Run Locally

1. Install dependencies:
   ```bash
   pip install -r requirements.txt

2. Authenticate the first time to create token.pickle:

bash
Copy
Edit
python main.py

3. Confirm that result/log.xlsx is created and updated.


🙌 Contributing
Feel free to open issues or pull requests to improve functionality or automation.

📄 License
This project is provided under the MIT License.

yaml
Copy
Edit

---

✅ **Pro tip:**
Save this as `README.md` in your project root.  
It’s Markdown—GitHub will render it beautifully.

---

If you want, I can help you tweak it or add badges (e.g., build status).








