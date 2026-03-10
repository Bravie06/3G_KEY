# 3G KPI Report Generator

This application automatically generates a formatted Excel report for 3G KPIs, based on a raw data file.

It strictly respects the provided template and applies the same formatting and conditional logic without requiring an internet connection, thus guaranteeing the confidentiality of your data.

## Prerequisites
- [Python 3.x](https://www.python.org/downloads/) must be installed on your machine. Ensure you check "Add Python to PATH" during installation.

## Installation and Quick Start (Windows)
1. Double-click on the `run.bat` file included in this folder.
2. The script will automatically install the required dependencies (pandas, openpyxl).
3. The graphical user interface of the application will open automatically.

## Usage
1. Enter the name of the event (e.g., `Concert`).
2. Click on **Select Raw Data File** to choose your raw Excel file (either site-level or city-level data).
3. Click on **Generate Report**.
4. Wait while the tool processes the data (extracts the last 10 hours, applies formatting).
5. A success message will be displayed once the file has been successfully generated in the same folder as the program.

## Confidentiality
This tool runs 100% locally on your machine. No data is sent to an external server or API, strictly respecting your security requirements.
