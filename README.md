# KPI Report Automator

This is a local Python tool to automate the generation of KPI reports from raw Excel data, matching a specific template format and applying predefined conditional formatting logic.

## Confidentiality
This tool operates **100% locally**. It does not use any external APIs or services, ensuring that your data remains fully confidential and secure on your local machine.

## Prerequisites
- Python 3.7+ installed.

## How to Run

### Windows (Simplest method)
1. Double click on `run.bat`.
2. This will automatically install the required Python packages (`pandas`, `openpyxl`) and launch the Graphical User Interface (GUI).

### Manual method
1. Open a terminal or command prompt in this directory.
2. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the script:
   ```bash
   python main.py
   ```

## Usage
1. **Template File:** Click "Browse" and select the template Excel file.
2. **Raw Data File:** Click "Browse" and select the raw data Excel file containing the latest KPI metrics.
3. **Save As:** Click "Browse" and select the location and filename where the generated report will be saved.
4. Click **Generate Excel Report**.

## Logic and Rules applied
- The script processes data for the **last 10 hours** found in the raw data file.
- It matches nodes and KPI variables between the raw data and the template.
- It applies formatting based on standard rules:
  - **Availability:** >= 98.5 is Green.
  - **CSSR (CS and PS):** < 98.5 is Faded Red.
  - **Call Drop (CS and PS):** <= 0.7 is Green.
  - **Traffic:** Remains White (no color applied).
