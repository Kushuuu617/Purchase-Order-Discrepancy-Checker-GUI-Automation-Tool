# 🔍 Purchase Order Discrepancy Checker – GUI Automation Tool

A robust **Python GUI application** built with **Tkinter** and **Selenium** to automate the process of verifying Purchase Orders (POs) from an ERP system.

This tool checks if the **billed amount** is less than the **total PO amount**, skipping “Global Blanket Agreement” POs, and exports discrepancies to an Excel file. It comes with a **user-friendly GUI** and real-time logging.

---

## ✅ Features

-  GUI-based for ease of use (no CLI needed)
-  Automates ERP login, navigation, and PO access
-  Compares **billed** vs **total** amount
-  Skips **Global Blanket Agreement** POs
-  Supports paginated PO tables
-  Exports discrepant POs to Excel
-  Real-time status updates within the app
-  Reduces manual checking effort by 90%

---
## 🧠 How It Works

Behind the scenes:

-  Launches Chrome via Selenium
-  Logs into the ERP system
-  Navigates paginated PO results
-  Opens each PO in a new tab
-  Compares total and billed amounts
-  Logs discrepancies
-  Saves under-billed PO numbers in an Excel workbook

## 🛡️ Disclaimer
⚠️ This automation script is tailored for Indus Towers ERP system.
If your ERP interface differs, manual adjustments may be needed.
This project is meant for internal use, testing, or learning purposes.
Always ensure compliance with your organization’s automation policies before use.

## 🖼️ Preview

```text
Discrepancy found: PO123456
Discrepancy found: PO123987
✔ Excel file saved successfully.
✅ All pages processed.
Browser closed.



