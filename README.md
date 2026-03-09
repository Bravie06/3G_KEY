# KPI Report Automator

This application automates the generation of an Excel KPI report. It reads raw data and maps it to a standard template, filtering for the last 10 hours and applying proper formatting and color coding.

## Confidentiality
This tool runs strictly locally on your machine. It does **not** send any data to external APIs or cloud services, ensuring total confidentiality of your data.

## Prerequisites
- Python 3.8 or higher installed on your computer.

## Installation and Execution
1. Double click on `run.bat` on Windows. This script will automatically install the necessary local libraries from `requirements.txt` and launch the application.
2. Alternatively, open a terminal in this directory and run:
   ```bash
   pip install -r requirements.txt
   python main.py
   ```

## Usage
1. Open the application.
2. Select the Raw Data file (Excel format).
3. Select the Template file (Excel format).
4. Choose the Output file location and name.
5. Click "Generate Report".
6. The application will output an Excel file with a single sheet containing the formatted and color-coded KPIs.
