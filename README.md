# CQ091 Report Verification Tool

## Overview

This Python automation tool verifies the **CQ091 report** against design specifications, checking for:

- Version consistency
- Column alignment
- Data completeness
- Business rule compliance

The tool provides detailed error reporting to help development teams quickly identify and fix issues.

## Features

- **Cover Page Validation**: Checks title spelling, version number, and ETL dates
- **Column Comparison**: Compares design spec columns with actual report columns across all three standards
- **Specific Case Validation**: Verifies specific case numbers and their date fields
- **Summary Report Validation**: Ensures summary fields match design specifications
- **Business Rule Compliance**: Validates sensitivity settings and business formulas
- **Detailed Error Analysis**: Provides specific error types (spelling, spacing, content differences)

## Requirements

- Python 3.6+

Required packages:

- `pandas`
- `openpyxl`

Install dependencies with:

```bash
pip install pandas openpyxl
```

File Structure
```nginx
text
├── cq091_verification.py    # Main verification script
├── design_spec.xlsx         # Design specification file
├── verification_report.xlsx # Report to be verified
└── README.md                # This file
```

## Usage

- Update the file paths in the main execution section if needed:

    ```python
    design_spec_path = r"CQ091 - Design Spec - QIP9.KS2 - Seven Day Visit 2025.xlsx"
    
    verification_path =r"Final verification-CQ091 - QIP 9, 11 - KS2 - Private Visits - Kinship Service_Children in Care.xlsx"
    ```

- Set the expected version:

    ```python
    expected_version = "1.3"
    ```

- Run the script:

    ```bash
    python cq091_verification.py
    ```

## Test Categories

1. Cover Page Tests
    
    - Title spelling verification
    - Version number validation
    - ETL date validation (checks if start date is before completion date)

2. Standard Report Tests

    Compares column headers between design specification and actual report for:
    - Standard 1 Report
    - Standard 2 Report
    - Standard 3 Report

3. Specific Cases Test

    Validates specific case numbers and their:
    - 30 Day Private Visit Due Date - 2025
    - 30 Day Private Visit Contact Log Start Date - Extension

4. Summary Report Test

    Compares summary field names between design spec and actual report.

5. Business Rule Tests

    - Sensitivity level validation
    - Business formula verification

## Error Types Detected

The tool categorizes errors into:

- Space differences: Extra or missing spaces in column names
- Case differences: Upper/lower case inconsistencies
- Spelling errors: Similarity-based spelling issues (with similarity score)
- Word order differences: Same words in different order
- Content differences: Completely different content
- Column count mismatches: Different number of columns

## Sample Output

The tool provides detailed output showing:

- Overall pass/fail status for each test category
- Specific error details with column/row numbers
- Design vs. verification comparisons
- Detailed error analysis for debugging

## Customization

To modify the specific cases being tested, update the `case_numbers_to_check` list in the `test_specific_cases_dates()` function:
```python
case_numbers_to_check = ['12891050', '13141575', '11739608', '13038729', '13155126']
```

## Troubleshooting

- File Not Found Errors: Ensure the file paths are correct and accessible.
- Sheet Name Errors: Verify that sheet names match between design spec and verification report.
- Date Format Issues: The tool expects ETL dates in `"dd-MMM-yyyy hh:mm:ss AM/PM"` format.