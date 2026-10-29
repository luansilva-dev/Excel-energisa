# SIGOD Pole Correction Automation

This project is a Python script created to help with the pole correction process using Excel spreadsheets and SIGOD data exports.

The idea is to make the process easier by comparing pole codes, finding possible mismatches, and generating a report so the results can be checked without having to compare everything manually.

## How it works

The script reads two Excel sheets: one containing the pole correction requests and another containing data exported from SIGOD. It compares the information, filters records from Mato Grosso do Sul (MS), and generates an Excel file with the results for review.

## Requirements

- Python 3.10+
- Pandas
- OpenPyXL

Install the required libraries:

```bash
pip install pandas openpyxl
```

## Running the script

1. Download this repository.
2. Fill in the Excel template with the required data.
3. Make sure the workbook has two sheets named `Requests` and `SIGOD_Export`.
4. Run the script:

```bash
python sigod_pole_correction_prototype.py
```

The results will be saved in `pole_correction_review_report.xlsx`.

## Current limitations

This is an early prototype. It works with Excel files and does not connect directly to SIGOD or make changes to its records. The generated report is intended to be checked manually.

## Future plans

- Improve the comparison rules.
- Handle missing or invalid data.
- Make the script easier to use.
- Explore direct integration if approved by the company.

## Author Luan Silva

Developed as a personal learning project focused on Python, Excel automation, and improving the pole correction workflow.
