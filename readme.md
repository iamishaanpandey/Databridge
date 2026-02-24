# ST DataBridge

**Version:** 2.6
**Author:** Ishaan Pandey
**Organization:** STMicroelectronics  

## Overview
ST DataBridge is an intelligent, event-driven Excel VBA solution designed to safely bridge the gap between complex source datasets and user-friendly editing. It allows users to query external workbooks, filter data using up to 12 criteria, edit specific columns in a safe cache environment, and push updates back to the source file using a high-speed Scripting Dictionary hash map—all without manually opening or corrupting the master files.

## Repository Structure
This repository contains the raw, uncompiled VBA source code for the tool. The code is divided into three class/basic files:

* `Module1_MainLogic.bas`: The core engine. Contains the file I/O operations, data caching, AutoFilter logic, and the dictionary-based write-back algorithms.
* `Sheet_Dashboard.cls`: The UI controller. Contains `Worksheet_Change` event listeners that natively simulate multi-select dropdown menus and handle real-time user validation.
* `ThisWorkbook.cls`: The security wrapper. Contains `Workbook_Open` events that enforce version control checks and lock down the user interface.

## Prerequisites
To rebuild this tool from scratch, your target Excel file must meet the following structural requirements:
1. **File Type:** Saved as an Excel Macro-Enabled Workbook (`.xlsm`).
2. **Sheet Names:** Must contain three exactly named sheets:
   * `Dashboard`
   * `Cache` (Should be set to `xlSheetVeryHidden` in production)
   * `Dropdowns`
3. **Named Ranges:** The code relies heavily on predefined ranges on the Dashboard. You must define these in the Excel Name Manager (e.g., `Path_Source`, `Sheet_Source`, `Header_Row`, `Key1_Fetch` through `Key12_Fetch`, `Col1_Update` through `Col12_Update`, etc.).

## Installation Instructions
To deploy this code into a fresh workbook:

1. Open your target `.xlsm` file in Excel.
2. Press `ALT + F11` to open the VBA Editor.
3. **Install Main Logic:** Right-click your VBAProject -> `Insert` -> `Module`. Copy the contents of `Module1_MainLogic.bas` into this new module.
4. **Install Dashboard Events:** In the Project Explorer, double-click the sheet object representing your Dashboard. Copy the contents of `Sheet_Dashboard.cls` into this code window.
5. **Install Workbook Events:** Double-click the `ThisWorkbook` object in the Project Explorer. Copy the contents of `ThisWorkbook.cls` into this code window.
6. Save the workbook and restart Excel to trigger the initialization routines.

## System Requirements & Limitations
* **OS:** Windows (relies on Windows Scripting Host for the `Scripting.Dictionary` and `Scripting.FileSystemObject`).
* **IME Note:** Users on non-Latin layouts (Japanese, Korean, Chinese) experiencing character artifacting (e.g., random smileys or dots) must revert to the "previous version of Microsoft IME" in their Windows Language settings.
