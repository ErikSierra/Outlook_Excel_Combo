# Outlook Email Data Extractor

## Introduction
This Python project automates the extraction of specific email attachments from Microsoft Outlook, combines their data, and saves the result into an Excel file. It's particularly useful for aggregating data from emails received over a recent period, allowing for automated data processing and analysis.

## Installation
To use this script, you need to have Python installed on your system along with the following dependencies:
- pandas
- pywin32
- colorama

You can install these dependencies via pip:
```bash
pip install pandas pywin32 colorama
```


## Usage
To run this script, execute the following command in your terminal:
```bash
python OutlookCombine.py
```

## Features
- Connects to Microsoft Outlook and accesses the Inbox.
- Filters emails based on the received time (last 6 days by default).
- Identifies and processes attachments from specific emails.
- Combines data from multiple CSV files into a single DataFrame.
- Saves the combined data into an Excel file.

## Dependencies
The script requires the following Python libraries:
- `pandas` for data manipulation.
- `win32com.client` for interacting with Outlook.
- `datetime` for date and time operations.
- `os` for interacting with the operating system.
- `colorama` for colored terminal output.

## Configuration
No additional configuration is needed to run the script as long as Microsoft Outlook is installed and configured on your machine.

## Documentation
The script is self-documented with comments explaining the steps involved in processing the emails.

## Troubleshooting
- Ensure Outlook is open and configured on your machine.
- Check if the script has permissions to access Outlook data.
- Verify the path and permissions for saving files on your system.

## Note: When running the script multiple times, make sure to move the previously generated Excel file to a different directory or rename it to avoid overwriting it with the new file.
