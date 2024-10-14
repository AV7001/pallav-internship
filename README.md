# Excel Convertor

## Overview
`Excel Convertor` is a Python utility designed to streamline the conversion and processing of Excel files. This tool provides options to convert both single and multiple Excel files by reformatting and adjusting their structure based on specific columns and data.It was mainly designed for the telecommunication companies.
It sorts the files by `SiteID` and organizes the output with essential columns for better clarity.

## Features
- **Single File Conversion**: Process an individual Excel file to adjust the data format and remove the original file after conversion.
- **Multiple File Conversion**: Automatically processes all `.xlsx` files in a given folder.
- **Customizable Columns**: Keeps the relevant columns such as `Site`, `SiteID`, `Name`, `OccuredOn`, and `ClearedOn` in the output file.
- **User-friendly Interface**: Provides a simple menu-driven interface for ease of use.

## Requirements

- Python 3.x
- Required libraries:
  - `pandas`
  - `openpyxl`
  - `tqdm`
  - `art`

You can install the required libraries using the following command:

```bash
pip install pandas openpyxl tqdm art



### Instructions:
- Replace `/path/to/excel-file.xlsx` and `/path/to/excel-folder` with your actual file paths.
- Make sure the required Python libraries are installed as listed in the requirements section.

Let me know if you need any changes or additional details!
