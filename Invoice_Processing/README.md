# Automated Invoice Processing Bot (UiPath)

An end-to-end Robotic Process Automation (RPA) solution built using UiPath Studio to automate the extraction, consolidation, and reporting of PDF invoice data.

---

## 📌 Overview

This UiPath workflow automatically monitors Gmail for incoming PDF invoices, extracts key fields using Regex patterns, consolidates the structured data into an Excel report, and emails the compiled report to the Finance team.

---

## 🚀 Key Features

* **Automated Attachment Extraction:** Iterates through Gmail inbox items to download incoming PDF invoice attachments.
* **Data Cleansing:** Deletes older output Excel files prior to execution to ensure fresh data processing.
* **PDF Text Extraction & Parsing:** Reads PDF contents and applies `System.Text.RegularExpressions` to accurately extract fields.
* **Data Consolidation:** Appends extracted data directly into a structured Data Table (`Dt_Invoices`) and exports it to `Excel\Invoices.xlsx`.
* **Automated Email Reporting:** Sends the compiled Excel report to the designated Finance team via Gmail Integration.

---

## 🛠 Target Extracted Fields

Using Regex patterns, the workflow extracts the following data points from each PDF invoice:
* **Invoice Number**
* **Vendor Name**
* **Invoice Date**
* **Due Date**
* **Total Amount**
* **Tax Amount**

---

## ⚙️ Process Workflow

1. **Email Monitoring & Download:**
   * Scans Gmail Inbox (`For Each Email`).
   * Downloads attachments (`Download Email Attachments`).

2. **Initialization:**
   * Clears prior reports using `Delete File` (`Excel\Invoices.xlsx`).
   * Initializes target data schema using `Build Data Table`.

3. **PDF Processing Loop:**
   * Loops through PDF files (`For Each File in Folder`).
   * Extracts raw text via `Read PDF Text`.
   * Applies Regular Expressions in a `Multiple Assign` activity for field extraction.

4. **Data Aggregation & Output:**
   * Appends fields to `Dt_Invoices` via `Add Data Row`.
   * Writes the consolidated table to `Excel\Invoices.xlsx` using `Write Range Workbook`.

5. **Notification:**
   * Uses `Send Email` (GSuite) to deliver the generated Excel report to the Finance team.

---

## 📋 Prerequisites & Dependencies

* **UiPath Studio Community / Enterprise**
* **UiPath Packages:**
  * `UiPath.GSuite.Activities`
  * `UiPath.PDF.Activities`
  * `UiPath.Excel.Activities`
* **Gmail Account Integration** configured within UiPath Integration Service or GSuite scope.

---

## 📂 File Structure

```text
├── Main.xaml                      # Primary workflow file
├── Excel/
│   └── Invoices.xlsx              # Output generated report
└── project.json                   # UiPath project dependencies and settings
