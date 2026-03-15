# 📊 Data Fetcher → Recap Sheet Sync (Google Sheets Automation)

Automating the **second stage of a BI data pipeline inside Google Sheets** ⚙️

This project synchronizes dynamically fetched datasets into structured **recap / reporting sheets**, ensuring dashboards and analytics tables remain stable even when the underlying dataset size changes.

It complements the repository **google-sheets-data-fetcher-automation**, which focuses on **retrieving data from BI platforms or analytics APIs into Google Sheets**.

Together, the two projects demonstrate a lightweight **BI data pipeline built entirely with Google Apps Script** 🚀

---

# 💡 Why This Project Exists

When working with analytics platforms or BI tools, data is often exported or fetched into Google Sheets automatically.

However, the number of rows in the dataset frequently changes depending on query results 📈📉

This causes common issues such as:

❌ dashboards referencing incorrect ranges
❌ recap sheets breaking when row counts change
❌ manual copy-paste workflows
❌ inconsistent reporting tables

This script solves the problem by **automatically synchronizing the fetched dataset into a structured recap sheet** while preserving formulas used for reporting.

---

# 🔗 Data Pipeline

This project represents the **second stage of a reporting pipeline**.

```text
BI Platform / Analytics API
        ↓
Google Apps Script Data Fetcher
        ↓
Raw Dataset Sheet
        ↓
Recap Sheet Sync Automation (this project)
        ↓
Reporting / Dashboard Sheets
```

The result: **stable reporting tables even when source datasets change** 📊

---

# ⚙️ What This Automation Does

The script automatically:

✅ detects the latest dataset in the sheet
✅ handles **dynamic dataset sizes**
✅ inserts additional rows when required
✅ copies values into recap tables
✅ preserves reporting formulas
✅ optionally sends notifications 🔔

This removes repetitive manual data preparation steps.

---

# 🧩 Example Use Cases

This automation is useful for:

📊 BI reporting pipelines
📈 KPI recap sheets
📉 operational dashboards
⚡ automated analytics workflows
📑 staging → reporting sheet synchronization

Any workflow where **data is fetched dynamically but reports require stable structures**.

---

# 🔄 Example Workflow

1️⃣ Data is retrieved from an analytics platform using Apps Script
2️⃣ The dataset is written into a **raw data sheet**
3️⃣ This script synchronizes that dataset into a **recap sheet**
4️⃣ The recap sheet feeds dashboards, summaries, or metrics

This allows reporting sheets to remain consistent while the source dataset evolves 🔁

---

# 📁 Project Structure

```text
data-fetcher-to-recap-sheet-sheets-sync
│
├── src
│   └── syncRecapSheet.js
│
├── examples
│   └── sheet-structure.md
│
├── README.md
└── appsscript.json
```

---

# 🔗 Related Project

This repository complements:

**google-sheets-data-fetcher-automation**

That project demonstrates how to connect Google Sheets with analytics platforms and retrieve data through APIs 🔌

This repository focuses on **processing and synchronizing the fetched dataset into recap sheets used for reporting**.

Together they form a simple **Google Sheets–based data automation pipeline**.

---

# 🛠 Technologies Used

🟢 Google Apps
