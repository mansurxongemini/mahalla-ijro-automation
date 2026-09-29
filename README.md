# 🏢 Mahalla Ijro — Public Records & Civil Ingestion Automation

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-blue.svg?logo=python&logoColor=white)](https://www.python.org/)
[![Playwright](https://img.shields.io/badge/Playwright-Automation-45ba4b.svg?logo=playwright&logoColor=white)](https://playwright.dev/python/)
[![OpenPyXL](https://img.shields.io/badge/OpenPyXL-Excel-217346.svg?logo=microsoft-excel&logoColor=white)](https://openpyxl.readthedocs.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Mahalla Ijro Automation** is a high-reliability civic record verification and spreadsheet synchronization utility built with asynchronous Python, Playwright browser automation, and OpenPyXL.

---

## 🌟 Features
- **🔍 Automated PINFL/JSHSHIR Verification**: Validates citizen residency and demographic entries against administrative portal endpoints.
- **📊 Spreadsheet Ingestion & Deduplication**: High-speed processing of structured excel registers with persistent caching.
- **🛡️ Error Resilient**: Auto-recovery on network timeouts, multi-retry logic, and dynamic state checkpointing.

---

## 🚀 Quickstart

```bash
# 1. Setup environment
python3 -m venv venv
source venv/bin/activate
pip install openpyxl playwright
playwright install chromium

# 2. Run automation
python test.py
```

---

## 📜 License
Distributed under the **MIT License**.
