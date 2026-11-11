# 🏛️ ESIC Supplementary Monthly Contribution & Challan Automation Suite

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Selenium](https://img.shields.io/badge/Selenium-Automation-43B02A?logo=selenium&logoColor=white)](https://www.selenium.dev/)
[![Playwright](https://img.shields.io/badge/Playwright-Automation-2EAD33?logo=playwright&logoColor=white)](https://playwright.dev/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

An automated solution for **ESIC Supplementary Monthly Contributions** and **Challan Generation** directly on the official ESIC portal. Automates multi-establishment employee entry, challan calculations, CAPTCHA processing, and submission tracking.

---

## 🌟 Key Capabilities

- **⚡ Bulk Supplementary Entry**: Automated entry of supplementary contributions for employees across establishments (`esic_supplimentry_entry.py`, `esic_full_supplimentry_entry 1.py`).
- **🧾 Automated Challan Generation**: End-to-end flow to compute, verify, and generate monthly contribution challans (`challangeneration.py`).
- **🤖 Robust Browser Automation**: Uses Selenium & Playwright with intelligent element waits and automated CAPTCHA handling.
- **🛡️ Data Security & Protection**: Live employee credentials, salaries, and ESIC numbers (`esic_users.xlsx`, `suplimentry_esic_users.xlsx`) are strictly git-ignored to prevent data leakage.

---

## 📂 Project Architecture

```
Supplementary Monthly Contributio/
├── esic_supplimentry_entry.py         # Primary supplementary contribution entry engine
├── esic_full_supplimentry_entry 1.py  # Production extended supplementary entry flow
├── challangeneration.py               # Automated challan generation script
├── esic_users.example.csv             # Template format for employer/employee inputs
├── .gitignore                         # Security rules excluding credentials & runtime data
└── README.md                          # Project documentation
```

---

## 📋 Input Data Format

The system expects an input spreadsheet with the following column structure (see [`esic_users.example.csv`](esic_users.example.csv)):

| Column Header | Description |
|---|---|
| `ESIC Number` | 10-digit employee Insurance Number (IP) |
| `Paid Days` | Number of days worked in the contribution period |
| `Basic Payable` | Basic wages/salary payable for the period |
| `loginid` | ESIC employer portal username |
| `password` | ESIC employer portal password |
| `Month` | Contribution month (e.g. `10`) |
| `year` | Contribution year (e.g. `2026`) |
| `Update Status` | Auto-populated execution status per row |

---

## 🚀 Setup & Execution

### 1. Prerequisites
- Python 3.10 or higher
- Google Chrome & compatible ChromeDriver / Playwright browsers

### 2. Installation
```bash
git clone https://github.com/KDI4U/ESIC-Supplementary-Monthly-Contribution.git
cd ESIC-Supplementary-Monthly-Contribution

python -m venv .venv
.venv\Scripts\activate

pip install selenium playwright pandas openpyxl pillow
playwright install
```

### 3. Execution
```bash
# Run supplementary contribution entry
python esic_supplimentry_entry.py

# Or run challan generation
python challangeneration.py
```

---

## 🔒 Security Notice
All operational data files, user credentials (`*.xlsx`), downloaded challans, and captcha captures are strictly excluded via `.gitignore` to maintain confidential data protection.

---

## 📄 License
This project is licensed under the MIT License.

---

<div align="center">
  <b>Developed & Maintained by <a href="https://github.com/KDI4U">Kundan Kumar</a></b>
</div>
