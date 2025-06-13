Purchase Order Discrepancy Checker (Selenium Automation)

A Python automation tool to streamline the manual process of verifying Purchase Orders (POs) in an ERP system. This script logs into the portal, navigates through POs, checks if the billed amount is less than the total amount, and exports such POs to an Excel file.

##• Features

- Automates ERP login and navigation
- Parses all paginated PO entries
- Skips "Global Blanket Agreement" document types
- Compares billed vs total amount
- Exports results to Excel
- Reduces manual effort by 90%

##• Preview

> Example:  
> Checks POs like  
> `PO12345 | Total: ₹100000 | Billed: ₹75000 → Saved to Excel`

##• Requirements

- Python 3.7+
- Google Chrome
- ChromeDriver (installed and added to system PATH)


