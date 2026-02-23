# BMW DSR Generator User Guide

## Introduction
The BMW DSR Generator is a desktop utility designed for the Docs Team to automate the creation of the BMW Daily Status Report (DSR). It takes a raw data export from the Logisys system (Excel format) and automatically transforms, calculates duties, extracts VINs, and styles the data into the finalized Nagarkot-branded BMW DSR format.

## How to Use

### 1. Launching the App
Simply double-click the `bmw_dsr_gui.exe` application to launch the full-screen interface. No internet connection or special credentials are required.

### 2. The Workflow (Step-by-Step)
1. **Prepare Input File**: Ensure you have downloaded your standard Logisys export as an `.xlsx` file.
2. **Select File**: Click the **Browse** button and locate your Logisys Excel file on your computer.
3. **Generate**: Click the **Generate BMW DSR** button.
   - *Note: The system requires specific column headers to exist in your Logisys export (e.g., 'Job No', 'Product Desc', 'Unit Price', 'Total Duty (INR)').*
4. **Retrieve Output**: The application will show a spinning cursor while processing. Once complete, it will display a Success popup and automatically open the folder containing your brand new BMW DSR file.

## Interface Reference

| Control / Input | Description | Expected Format |
| :--- | :--- | :--- |
| **Browse (Button)** | Opens a file explorer to select the input data. | `.xlsx` file |
| **File Path (Text Box)** | Displays the location of the selected file. | Read-only text |
| **Generate BMW DSR (Button)** | Triggers the data extraction and formatting process. | N/A |

## Data Processing Rules & Outputs
When generating the report, the application applies these specific business rules:
- **VIN Extraction**: Automatically scans the `Product Desc` field and extracts the 17-character VIN.
- **Duty Calculation**: Automatically calculates `TOTAL DUTY` by adding `Total Duty (INR)` (Bank Payment) and `BCD Foregone` (Licence Debit).
- **Group Tracking**: Vehicles are grouped by `Job No`. If a shipment contains more than 1 car, the `NO OF CARS` total cell will be automatically highlighted in bold yellow.
- **Output Naming**: The final file is saved in the exact same folder as your input file, named automatically with the date and time (e.g., `BMW_DSR_20260223_1330.xlsx`).

## Troubleshooting & Validations

If you see an error, check this table:

| Message | What it means | Solution |
| :--- | :--- | :--- |
| **"Please select the Logisys Excel file."** | You clicked Generate before choosing a file snippet. | Click the "Browse" button to select your input file first. |
| **"Failed to generate BMW DSR: KeyError: '[Column Name]'"** | The uploaded Logisys Excel file is missing a mandatory column header. | Ensure your export includes standard columns like `Job No`, `Product Desc`, `Total Duty (INR)`, `Unit Price`, etc. Do not manually delete columns from the export. |
| **"Failed to generate BMW DSR: PermissionError"** | The system tried to save the new DSR, but a file with that exact name is already open. | Close the pending Excel files in your taskbar and try generating the report again. |
| **Empty VIN in Output** | The application could not find a standard 17-character VIN in the product description. | The input `Product Desc` must strictly contain the text "VIN NO " followed by 17 uppercase alphanumeric characters. |
