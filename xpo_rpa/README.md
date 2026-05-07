# XPO Billing RPA

Automates downloading bills from the XPO LTL portal (https://ext-web.ltl-xpo.com) in batches of 10, organized by download date, with a full Excel report.

---

## 📁 Project Structure

```
xpo_rpa/
├── main.py                    ← Run this to start the RPA
├── requirements.txt
├── README.md
├── input_data/
│   └── credentials.json       ← Put your login here
├── downloads/
│   └── YYYY-MM-DD/            ← Bills saved here (date-organized)
│       └── Bills_Report_YYYY-MM-DD.xlsx
├── logs/
│   └── rpa_YYYYMMDD_HHMMSS.log
└── rpa/
    ├── config.py
    ├── browser.py
    ├── runner.py
    ├── report.py
    └── logger.py
```

---

## ⚙️ Setup

### 1. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 2. Install Google Chrome

Download and install from: https://www.google.com/chrome/

### 3. Add your credentials

Edit `input_data/credentials.json`:

```json
{
  "login_id": "your_actual_username",
  "password": "your_actual_password"
}
```

> ⚠️ Keep this file private. Never commit it to version control.

---

## ▶️ Run the RPA

```bash
python main.py
```

That's it. The RPA will:
1. Open Chrome and navigate to the XPO portal
2. Log in with your credentials
3. Navigate to the Billing section
4. Sort bills by latest date
5. Download bills 10 at a time
6. Save all files in `downloads/YYYY-MM-DD/`
7. Generate `Bills_Report_YYYY-MM-DD.xlsx` with full details

---

## 📊 Excel Report Contents

| Sheet | Description |
|-------|-------------|
| **Downloaded Bills** | Full list: bill number, date, batch, file name, status |
| **Batch Summary** | Totals per batch — successful vs failed |
| **Failed Downloads** | (if any) Bills that failed — for manual follow-up |

---

## 🗂️ Downloads Folder

Bills are saved in:
```
downloads/
└── 2025-01-15/        ← Date of download
    ├── invoice_001.pdf
    ├── invoice_002.pdf
    └── Bills_Report_2025-01-15.xlsx
```

---

## 🔧 Troubleshooting

| Issue | Fix |
|-------|-----|
| `ChromeDriver not found` | Chrome auto-installs driver via `webdriver-manager`. Ensure Chrome is installed. |
| `Login failed` | Check credentials in `input_data/credentials.json` |
| `No bills found` | The portal's HTML may have changed. Check `logs/` for details and open a GitHub issue. |
| `Download timed out` | Slow internet or portal. Increase `DOWNLOAD_TIMEOUT` in `browser.py` |

---

## 🛡️ Security Notes

- Credentials are stored locally in `input_data/credentials.json`
- Never share this file or commit it to git
- Add `input_data/` to your `.gitignore`
