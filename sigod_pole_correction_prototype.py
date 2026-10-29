"""
SIGOD Pole Correction Review Prototype
--------------------------------------
Reads two worksheets from an Excel workbook: correction requests and an
AUTHORIZED SIGOD export. Compares codes and creates a review report.

IMPORTANT:
- This prototype does NOT access Citrix/SIGOD, log in, or modify records.
- Use only files and data you are authorized to access.
- Adjust worksheet and column names to match your workbook.
- Review all results before making any official correction.

Requirements:
    pip install pandas openpyxl
"""

from pathlib import Path
import pandas as pd

# Update these names to match your actual workbook.
INPUT_WORKBOOK = Path("pole_correction_template.xlsx")
REQUESTS_SHEET = "Requests"
SIGOD_EXPORT_SHEET = "SIGOD_Export"
OUTPUT_REPORT = Path("pole_correction_review_report.xlsx")

REQUEST_CODE_COLUMN = "code"
EXPECTED_POLE_COLUMN = "expected_pole"
REQUEST_STATE_COLUMN = "state"  # e.g., MS

SIGOD_CODE_COLUMN = "code"
CURRENT_POLE_COLUMN = "current_pole"
SIGOD_STATE_COLUMN = "state"  # e.g., MS


def normalize_code(value):
    """Normalize codes as text, removing whitespace and Excel's trailing .0."""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text.upper()


def normalize_text(value):
    """Normalize text for comparison."""
    if pd.isna(value):
        return ""
    return " ".join(str(value).strip().upper().split())


def require_columns(dataframe, required_columns, source_name):
    missing = [column for column in required_columns if column not in dataframe.columns]
    if missing:
        raise ValueError(
            f"Missing columns in '{source_name}': {missing}. "
            f"Available columns: {list(dataframe.columns)}"
        )


def main():
    if not INPUT_WORKBOOK.exists():
        raise FileNotFoundError(f"Workbook not found: {INPUT_WORKBOOK.resolve()}")

    requests = pd.read_excel(
        INPUT_WORKBOOK, sheet_name=REQUESTS_SHEET, dtype=str
    ).fillna("")
    sigod = pd.read_excel(
        INPUT_WORKBOOK, sheet_name=SIGOD_EXPORT_SHEET, dtype=str
    ).fillna("")

    require_columns(
        requests,
        [REQUEST_CODE_COLUMN, EXPECTED_POLE_COLUMN, REQUEST_STATE_COLUMN],
        f"{INPUT_WORKBOOK.name} / sheet {REQUESTS_SHEET}",
    )
    require_columns(
        sigod,
        [SIGOD_CODE_COLUMN, CURRENT_POLE_COLUMN, SIGOD_STATE_COLUMN],
        f"{INPUT_WORKBOOK.name} / sheet {SIGOD_EXPORT_SHEET}",
    )

    # Limit the initial scope to Mato Grosso do Sul (MS).
    requests = requests[
        requests[REQUEST_STATE_COLUMN].map(normalize_text).isin(["MS", "MATO GROSSO DO SUL"])
    ].copy()
    sigod = sigod[
        sigod[SIGOD_STATE_COLUMN].map(normalize_text).isin(["MS", "MATO GROSSO DO SUL"])
    ].copy()

    requests["_normalized_code"] = requests[REQUEST_CODE_COLUMN].map(normalize_code)
    sigod["_normalized_code"] = sigod[SIGOD_CODE_COLUMN].map(normalize_code)
    requests["_expected_pole_normalized"] = requests[EXPECTED_POLE_COLUMN].map(normalize_text)
    sigod["_current_pole_normalized"] = sigod[CURRENT_POLE_COLUMN].map(normalize_text)

    # Do not silently choose a record when the SIGOD export contains duplicate codes.
    code_counts = sigod["_normalized_code"].value_counts()
    duplicate_codes = set(code_counts[code_counts > 1].index)

    report_rows = []
    for _, request in requests.iterrows():
        code = request["_normalized_code"]
        expected_pole = request[EXPECTED_POLE_COLUMN]

        if not code:
            report_rows.append({
                "code": request[REQUEST_CODE_COLUMN],
                "expected_pole": expected_pole,
                "current_sigod_pole": "",
                "result": "REVIEW: empty code",
                "suggested_action": "Check the source spreadsheet",
            })
            continue

        if code in duplicate_codes:
            current_poles = sigod.loc[
                sigod["_normalized_code"] == code, CURRENT_POLE_COLUMN
            ].astype(str).tolist()
            report_rows.append({
                "code": request[REQUEST_CODE_COLUMN],
                "expected_pole": expected_pole,
                "current_sigod_pole": " | ".join(current_poles),
                "result": "REVIEW: duplicate code in SIGOD export",
                "suggested_action": "Manual review required",
            })
            continue

        matches = sigod[sigod["_normalized_code"] == code]
        if matches.empty:
            report_rows.append({
                "code": request[REQUEST_CODE_COLUMN],
                "expected_pole": expected_pole,
                "current_sigod_pole": "",
                "result": "NOT FOUND in SIGOD export",
                "suggested_action": "Check the code and the export file",
            })
            continue

        current_pole = matches.iloc[0][CURRENT_POLE_COLUMN]
        if normalize_text(current_pole) == normalize_text(expected_pole):
            result = "MATCH"
            action = "No correction suggested"
        else:
            result = "MISMATCH: review required"
            action = "Validate the expected pole before any correction"

        report_rows.append({
            "code": request[REQUEST_CODE_COLUMN],
            "expected_pole": expected_pole,
            "current_sigod_pole": current_pole,
            "result": result,
            "suggested_action": action,
        })

    report = pd.DataFrame(report_rows)
    report.to_excel(OUTPUT_REPORT, index=False)

    print(f"Done. Review report saved to: {OUTPUT_REPORT.resolve()}")
    print("This report is for review only; no changes were made in SIGOD.")
    if not report.empty:
        print("\\nSummary:")
        print(report["result"].value_counts().to_string())


if __name__ == "__main__":
    main()
