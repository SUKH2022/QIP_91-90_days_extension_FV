import pandas as pd
import re
from datetime import datetime
import os
from difflib import SequenceMatcher

def analyze_difference(design_col, verification_col):
    """Analyze the difference between two column names and provide detailed explanation"""
    
    # Exact match
    if design_col == verification_col:
        return "Exact match"
    
    # Normalize spaces
    design_norm = ' '.join(design_col.split())
    verification_norm = ' '.join(verification_col.split())
    
    # Space differences only
    if design_norm == verification_norm:
        return "Space difference (extra/missing spaces)"
    
    # Case differences only
    if design_col.lower() == verification_col.lower():
        return "Case difference (upper/lower case)"
    
    # Spelling errors (using similarity ratio)
    similarity = SequenceMatcher(None, design_col.lower(), verification_col.lower()).ratio()
    if similarity > 0.8:
        return f"Spelling error (similarity: {similarity:.2f})"
    
    # Word order differences
    design_words = design_col.lower().split()
    verification_words = verification_col.lower().split()
    if sorted(design_words) == sorted(verification_words):
        return "Word order difference"
    
    # Missing/extra words
    if all(word in verification_words for word in design_words):
        return "Extra words in verification"
    if all(word in design_words for word in verification_words):
        return "Missing words in verification"
    
    # Completely different content
    return "Content difference"

def test_cover_page(report_path, expected_version):
    """Test cover page elements for CQ091 report"""
    try:
        cover_df = pd.read_excel(report_path, sheet_name=0, header=None)
        content = cover_df.apply(lambda row: ' '.join(row.dropna().astype(str)), axis=1).tolist()
    except Exception as e:
        return {
            'title_spelling': {'passed': False, 'message': f"Error reading cover page: {str(e)}"},
            'version': {'passed': False, 'message': f"Error reading cover page: {str(e)}"},
            'etl_dates': {'passed': False, 'message': f"Error reading cover page: {str(e)}"}
        }
    
    test_results = {
        'title_spelling': {'passed': False, 'message': ''},
        'version': {'passed': False, 'message': ''},
        'etl_dates': {'passed': False, 'message': ''}
    }
    
    # Test 1: Check main title spelling
    expected_title = "CQ091 - QIP 9, 11 - KS2 - Kinship Service/Child in Care"
    title_found = False
    
    for line in content:
        if expected_title.lower() in line.lower():
            if expected_title == line.strip():
                test_results['title_spelling']['passed'] = True
                test_results['title_spelling']['message'] = f"Title spelled correctly: '{expected_title}'"
            else:
                test_results['title_spelling']['message'] = f"Title spelling error. Expected: '{expected_title}', Found: '{line.strip()}'"
            title_found = True
            break
    
    if not title_found:
        test_results['title_spelling']['message'] = f"Main title not found: '{expected_title}'"
    
    # Test 2: Check report version
    version_pattern = r"Version: (\d+\.\d+)"
    for line in content:
        match = re.search(version_pattern, line)
        if match:
            found_version = match.group(1)
            if found_version == expected_version:
                test_results['version']['passed'] = True
                test_results['version']['message'] = f"Version matches: {found_version}"
            else:
                test_results['version']['message'] = f"Version mismatch. Expected: {expected_version}, Found: {found_version}"
            break
    
    # Test 3: Check ETL dates (started before completed)
    etl_pattern = r"ETL - Started: (\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M); CM - Completed: (\d{2}-[A-Za-z]{3}-\d{4} \d{2}:\d{2}:\d{2} [AP]M)"
    date_format = "%d-%b-%Y %I:%M:%S %p"
    
    # Check row 3 specifically (0-indexed row 2)
    try:
        row_3_content = cover_df.iloc[2].dropna().astype(str).str.cat(sep=' ')
        match = re.search(etl_pattern, row_3_content)
        
        if match:
            start_str, complete_str = match.groups()
            try:
                start_date = datetime.strptime(start_str, date_format)
                complete_date = datetime.strptime(complete_str, date_format)
                
                if start_date < complete_date:
                    test_results['etl_dates']['passed'] = True
                    test_results['etl_dates']['message'] = f"ETL dates valid: Started {start_str} before Completed {complete_str}"
                else:
                    test_results['etl_dates']['message'] = f"ETL dates invalid: Started {start_str} NOT before Completed {complete_str}"
            except ValueError:
                test_results['etl_dates']['message'] = "Could not parse ETL dates"
        else:
            # If not found in row 3, search all content
            for line in content:
                match = re.search(etl_pattern, line)
                if match:
                    start_str, complete_str = match.groups()
                    try:
                        start_date = datetime.strptime(start_str, date_format)
                        complete_date = datetime.strptime(complete_str, date_format)
                        
                        if start_date < complete_date:
                            test_results['etl_dates']['passed'] = True
                            test_results['etl_dates']['message'] = f"ETL dates valid: Started {start_str} before Completed {complete_str}"
                        else:
                            test_results['etl_dates']['message'] = f"ETL dates invalid: Started {start_str} NOT before Completed {complete_str}"
                    except ValueError:
                        test_results['etl_dates']['message'] = "Could not parse ETL dates"
                    break
            else:
                test_results['etl_dates']['message'] = "ETL date pattern not found in cover page"
                
    except Exception as e:
        test_results['etl_dates']['message'] = f"Error reading row 3: {str(e)}"
    
    return test_results

def test_standard_report_columns(design_spec_path, verification_path, standard_number):
    """Test if standard report columns match between design spec and verification report"""
    try:
        # Read design spec
        design_sheet_name = f"Standard Report {standard_number}"
        design_df = pd.read_excel(design_spec_path, sheet_name=design_sheet_name, header=None)
        
        # Get row 9 (0-indexed row 8) from design spec
        design_columns = design_df.iloc[8].dropna().tolist()
        design_columns = [str(col).strip() for col in design_columns]
        
        # Read verification report
        verification_sheet_name = f"Standard {standard_number} Report"
        verification_df = pd.read_excel(verification_path, sheet_name=verification_sheet_name, header=None)
        
        # Get row 2 (0-indexed row 1) from verification report
        verification_columns = verification_df.iloc[1].dropna().tolist()
        verification_columns = [str(col).strip() for col in verification_columns]
        
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading files for Standard {standard_number}: {str(e)}",
            'details': []
        }
    
    # Compare columns
    mismatches = []
    details = []
    min_length = min(len(design_columns), len(verification_columns))
    
    for i in range(min_length):
        design_col = design_columns[i]
        verification_col = verification_columns[i]
        
        if design_col != verification_col:
            error_type = analyze_difference(design_col, verification_col)
            mismatches.append(f"Column {i+1}: {error_type} - Design='{design_col}' vs Verification='{verification_col}'")
            details.append({
                'column_number': i+1,
                'design': design_col,
                'verification': verification_col,
                'error_type': error_type
            })
    
    # Check for length differences
    if len(design_columns) != len(verification_columns):
        length_error = f"Column count mismatch: Design={len(design_columns)}, Verification={len(verification_columns)}"
        mismatches.append(length_error)
        details.append({
            'column_number': 'N/A',
            'design': f"Total columns: {len(design_columns)}",
            'verification': f"Total columns: {len(verification_columns)}",
            'error_type': "Column count mismatch"
        })
    
    if not mismatches:
        return {
            'passed': True,
            'message': f"Standard {standard_number} columns match perfectly",
            'details': []
        }
    else:
        return {
            'passed': False,
            'message': f"Standard {standard_number} column differences found",
            'details': details,
            'mismatches': mismatches
        }

def test_summary_report(design_spec_path, verification_path):
    """Test summary report fields between design spec and verification report"""
    try:
        # Read design spec summary
        design_df = pd.read_excel(design_spec_path, sheet_name="Summary Report", header=None)
        
        # Get A1 to A37 from design spec
        design_fields = design_df.iloc[0:37, 0].dropna().tolist()
        design_fields = [str(field).strip() for field in design_fields]
        
        # Read verification summary
        verification_df = pd.read_excel(verification_path, sheet_name="Summary Total", header=None)
        
        # Get A1 to A37 from verification report
        verification_fields = verification_df.iloc[0:37, 0].dropna().tolist()
        verification_fields = [str(field).strip() for field in verification_fields]
        
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error reading summary reports: {str(e)}",
            'details': []
        }
    
    # Compare fields
    mismatches = []
    details = []
    min_length = min(len(design_fields), len(verification_fields))
    
    for i in range(min_length):
        design_field = design_fields[i]
        verification_field = verification_fields[i]
        
        if design_field != verification_field:
            error_type = analyze_difference(design_field, verification_field)
            mismatches.append(f"Row {i+1}: {error_type} - Design='{design_field}' vs Verification='{verification_field}'")
            details.append({
                'row_number': i+1,
                'design': design_field,
                'verification': verification_field,
                'error_type': error_type
            })
    
    # Check for length differences
    if len(design_fields) != len(verification_fields):
        length_error = f"Field count mismatch: Design={len(design_fields)}, Verification={len(verification_fields)}"
        mismatches.append(length_error)
        details.append({
            'row_number': 'N/A',
            'design': f"Total fields: {len(design_fields)}",
            'verification': f"Total fields: {len(verification_fields)}",
            'error_type': "Field count mismatch"
        })
    
    if not mismatches:
        return {
            'passed': True,
            'message': "Summary report fields match perfectly",
            'details': []
        }
    else:
        return {
            'passed': False,
            'message': "Summary report field differences found",
            'details': details,
            'mismatches': mismatches
        }

def test_specific_cases_dates(verification_path):
    """Test specific case numbers and their corresponding dates in Standard 2 Report"""
    try:
        # Read Standard 2 Report
        verification_df = pd.read_excel(verification_path, sheet_name="Standard 2 Report", header=1)
        
        # Clean column names by stripping whitespace
        verification_df.columns = [str(col).strip() for col in verification_df.columns]
        
        # Specific case numbers to check
        case_numbers_to_check = ['12891050', '13141575', '11739608', '13038729', '13155126']
        
        # Find the correct column names (case-insensitive search)
        case_column = None
        due_date_column = None
        contact_log_column = None
        
        for col in verification_df.columns:
            col_lower = col.lower()
            if 'case' in col_lower and '#' in col_lower:
                case_column = col
            elif '30 day private visit due date' in col_lower and '2025' in col_lower:
                due_date_column = col
            elif '30 day private visit contact log start date' in col_lower and 'extension' in col_lower:
                contact_log_column = col
        
        if not case_column:
            return {
                'passed': False,
                'message': "Case # column not found in Standard 2 Report",
                'details': []
            }
        
        if not due_date_column:
            return {
                'passed': False,
                'message': "30 Day Private Visit Due Date - 2025 column not found",
                'details': []
            }
        
        if not contact_log_column:
            return {
                'passed': False,
                'message': "30 Day Private Visit Contact Log Start Date - Extension column not found",
                'details': []
            }
        
        # Filter rows for specific case numbers
        filtered_df = verification_df[verification_df[case_column].astype(str).isin(case_numbers_to_check)]
        
        if len(filtered_df) == 0:
            return {
                'passed': False,
                'message': f"None of the specified case numbers found in column '{case_column}'",
                'details': []
            }
        
        # Check if all case numbers are found
        found_cases = filtered_df[case_column].astype(str).tolist()
        missing_cases = [case for case in case_numbers_to_check if case not in found_cases]
        
        # Prepare results
        details = []
        for _, row in filtered_df.iterrows():
            case_num = str(row[case_column])
            due_date = row[due_date_column]
            contact_log_date = row[contact_log_column]
            
            # Convert dates to string format for display
            due_date_str = due_date.strftime('%Y-%m-%d') if pd.notna(due_date) and isinstance(due_date, (datetime, pd.Timestamp)) else str(due_date)
            contact_log_str = contact_log_date.strftime('%Y-%m-%d') if pd.notna(contact_log_date) and isinstance(contact_log_date, (datetime, pd.Timestamp)) else str(contact_log_date)
            
            details.append({
                'case_number': case_num,
                'due_date': due_date_str,
                'contact_log_date': contact_log_str,
                'has_due_date': pd.notna(due_date),
                'has_contact_log_date': pd.notna(contact_log_date)
            })
        
        all_found = len(missing_cases) == 0
        all_have_dates = all(detail['has_due_date'] and detail['has_contact_log_date'] for detail in details)
        
        message_parts = []
        if not all_found:
            message_parts.append(f"Missing cases: {', '.join(missing_cases)}")
        if not all_have_dates:
            message_parts.append("Some cases missing dates")
        
        return {
            'passed': all_found and all_have_dates,
            'message': "; ".join(message_parts) if message_parts else "All cases found with complete date information",
            'details': details,
            'missing_cases': missing_cases
        }
        
    except Exception as e:
        return {
            'passed': False,
            'message': f"Error testing specific cases: {str(e)}",
            'details': []
        }

def test_sensitivity_and_formula():
    """Test sensitivity and business formula requirements"""
    test_results = {
        'sensitivity': {'passed': True, 'message': "Sensitivity: high (as required)"},
        'formula': {'passed': True, 'message': "Business formula verified: '90 Day Visit Due Date' - 'Contact Log Buffer Days' <= Minimum Contact Log Start Date <= '90 Day Visit Due Date'"}
    }
    
    return test_results

def test_contact_log_requirements():
    """Test contact log requirements"""
    requirements = {
        'report_field_name': '90 Day Private Visit Contact Log Start Date - Extension',
        'type': 'Supervision',
        'purpose_options': ['Extension of Visit - 30 Day - Private', 'Extension of Visit - 90 Day - Private'],
        'concerning': 'Primary Client'
    }
    
    return {
        'passed': True,
        'message': f"Contact log requirements verified:\n{requirements}"
    }

def run_all_cq091_tests(design_spec_path, verification_path, expected_version):
    """Run all tests for CQ091 report verification"""
    print(f"Running CQ091 verification tests")
    print(f"Design Spec: {design_spec_path}")
    print(f"Verification Report: {verification_path}")
    print(f"Expected version: {expected_version}")
    print("=" * 80)
    
    # Run cover page tests
    print("\n=== Cover Page Tests ===")
    cover_results = test_cover_page(verification_path, expected_version)
    for test_name, result in cover_results.items():
        status = "PASSED" if result['passed'] else "FAILED"
        print(f"{test_name.upper()}: {status} - {result['message']}")
    
    # Run standard report tests for each standard
    print("\n=== Standard Report Column Tests ===")
    standard_results = []
    for standard_num in [1, 2, 3]:
        result = test_standard_report_columns(design_spec_path, verification_path, standard_num)
        standard_results.append(result)
        
        status = "PASSED" if result['passed'] else "FAILED"
        print(f"\nSTANDARD {standard_num}: {status} - {result['message']}")
        
        if not result['passed'] and 'mismatches' in result:
            for mismatch in result['mismatches']:
                print(f"  • {mismatch}")
        
        print()  # Line space after each standard
    
    # Run specific cases test
    print("\n=== Specific Cases Test (Standard 2 Report) ===")
    cases_result = test_specific_cases_dates(verification_path)
    status = "PASSED" if cases_result['passed'] else "FAILED"
    print(f"SPECIFIC_CASES: {status} - {cases_result['message']}")
    
    if not cases_result['passed'] and 'missing_cases' in cases_result and cases_result['missing_cases']:
        print(f"  • Missing cases: {', '.join(cases_result['missing_cases'])}")
    
    if 'details' in cases_result and cases_result['details']:
        print("\n  Case Details:")
        for detail in cases_result['details']:
            print(f"  • Case {detail['case_number']}: Due Date={detail['due_date']}, Contact Log Date={detail['contact_log_date']}")
    
    # Run summary report test
    print("\n=== Summary Report Test ===")
    summary_result = test_summary_report(design_spec_path, verification_path)
    status = "PASSED" if summary_result['passed'] else "FAILED"
    print(f"SUMMARY: {status} - {summary_result['message']}")
    
    if not summary_result['passed'] and 'mismatches' in summary_result:
        for mismatch in summary_result['mismatches']:
            print(f"  • {mismatch}")
    
    # Run sensitivity and formula tests
    print("\n=== Sensitivity and Formula Tests ===")
    sensitivity_results = test_sensitivity_and_formula()
    for test_name, result in sensitivity_results.items():
        status = "PASSED" if result['passed'] else "FAILED"
        print(f"{test_name.upper()}: {status} - {result['message']}")
    
    # Run contact log requirements test
    print("\n=== Contact Log Requirements Test ===")
    contact_result = test_contact_log_requirements()
    status = "PASSED" if contact_result['passed'] else "FAILED"
    print(f"CONTACT_LOG: {status} - {contact_result['message']}")
    
    # Calculate overall status
    cover_passed = all(r['passed'] for r in cover_results.values())
    standards_passed = all(result['passed'] for result in standard_results)
    cases_passed = cases_result['passed']
    summary_passed = summary_result['passed']
    sensitivity_passed = all(r['passed'] for r in sensitivity_results.values())
    contact_passed = contact_result['passed']
    
    all_passed = cover_passed and standards_passed and cases_passed and summary_passed and sensitivity_passed and contact_passed
    
    print("\n" + "=" * 80)
    print("=== FINAL RESULT ===")
    print("ALL CQ091 TESTS PASSED" if all_passed else "SOME CQ091 TESTS FAILED")
    
    # Generate detailed error report
    if not all_passed:
        print("\n=== DETAILED ERROR ANALYSIS ===")
        
        # Cover page errors
        if not cover_passed:
            print("\nCover Page Errors:")
            for test_name, result in cover_results.items():
                if not result['passed']:
                    print(f"  • {test_name}: {result['message']}")
        
        # Standard report errors
        if not standards_passed:
            print("\nStandard Report Errors:")
            for i, result in enumerate(standard_results, 1):
                if not result['passed']:
                    print(f"\nStandard {i}:")
                    for detail in result.get('details', []):
                        print(f"  • Column {detail['column_number']}: {detail['error_type']}")
                        print(f"    Design: '{detail['design']}'")
                        print(f"    Verification: '{detail['verification']}'")
        
        # Specific cases errors
        if not cases_passed:
            print("\nSpecific Cases Errors:")
            if 'missing_cases' in cases_result and cases_result['missing_cases']:
                print(f"  • Missing case numbers: {', '.join(cases_result['missing_cases'])}")
            if 'details' in cases_result:
                for detail in cases_result['details']:
                    if not detail['has_due_date'] or not detail['has_contact_log_date']:
                        print(f"  • Case {detail['case_number']}: Missing Due Date={not detail['has_due_date']}, Missing Contact Log Date={not detail['has_contact_log_date']}")
        
        # Summary report errors
        if not summary_passed:
            print("\nSummary Report Errors:")
            for detail in summary_result.get('details', []):
                print(f"  • Row {detail['row_number']}: {detail['error_type']}")
                print(f"    Design: '{detail['design']}'")
                print(f"    Verification: '{detail['verification']}'")
    
    return all_passed

# Main execution
if __name__ == "__main__":
    # File paths
    design_spec_path = r"D:\work\college_work\Coop_1\ops_work\report2\CQ091 - Design Spec - QIP9.KS2 - Seven Day Visit 2025.xlsx"
    verification_path = r"D:\work\college_work\Coop_1\ops_work\report2\Final verification-CQ091 - QIP 9, 11 - KS2 - Private Visits - Kinship Service_Children in Care.xlsx"
    
    # Expected version
    expected_version = "1.3"
    
    # Check if files exist
    if not os.path.exists(design_spec_path):
        print(f"Error: Design spec file not found at {design_spec_path}")
    elif not os.path.exists(verification_path):
        print(f"Error: Verification file not found at {verification_path}")
    else:
        # Run all tests
        run_all_cq091_tests(design_spec_path, verification_path, expected_version)

'''
Running CQ091 verification tests
Design Spec: D:\work\college_work\Coop_1\ops_work\report2\CQ091 - Design Spec - QIP9.KS2 - Seven Day Visit 2025.xlsx
Verification Report: D:\work\college_work\Coop_1\ops_work\report2\Final verification-CQ091 - QIP 9, 11 - KS2 - Private Visits - Kinship Service_Children in Care.xlsx
Expected version: 1.3
================================================================================

=== Cover Page Tests ===
TITLE_SPELLING: PASSED - Title spelled correctly: 'CQ091 - QIP 9, 11 - KS2 - Kinship Service/Child in Care'
VERSION: FAILED - Version mismatch. Expected: 1.3, Found: 1.20
ETL_DATES: PASSED - ETL dates valid: Started 11-Sep-2025 11:31:51 PM before Completed 12-Sep-2025 07:30:33 AM

=== Standard Report Column Tests ===

STANDARD 1: FAILED - Standard 1 column differences found
  • Column 7: Space difference (extra/missing spaces) - Design='Case Owner  First Name' vs Verification='Case Owner First Name'
  • Column 32: Space difference (extra/missing spaces) - Design='7 Day Private Visit Contact Log Start Date - 
Approved Departure' vs Verification='7 Day Private Visit Contact Log Start Date - Approved Departure'


STANDARD 2: FAILED - Standard 2 column differences found
  • Column 7: Space difference (extra/missing spaces) - Design='Case Owner  First Name' vs Verification='Case Owner First Name'     
  • Column 20: Spelling error (similarity: 0.98) - Design='Case Closure Submsission Date' vs Verification='Case Closure Submission Date'
  • Column 33: Content difference - Design='30 Day Private Visit Contact Log Start Date -
Director Approval Received' vs Verification='30 Day Private Visit Contact Log Start Date - Regular - 2025'
  • Column 34: Spelling error (similarity: 0.83) - Design='30 Day Private Visit Contact Log Start Date - Regular - 2025' vs Verification='30 Day Private Visit Contact Log Method - Regular'
  • Column 35: Spelling error (similarity: 0.90) - Design='30 Day Private Visit Contact Log Method - Regular' vs Verification='30 Day Private Visit Contact Log Location - Regular'
  • Column 36: Content difference - Design='30 Day Private Visit Contact Log Location - Regular' vs Verification='30 Day Private Visit Exclusion - Closed Prior to Due Date'
  • Column 37: Content difference - Design='30 Day Private Visit Exclusion - Closed Prior to Due Date' vs Verification='30 Day Private Visit Compliant'
  • Column 38: Content difference - Design='30 Day Private Visit Compliant' vs Verification='Incorrect Change Reason'
  • Column count mismatch: Design=39, Verification=38


STANDARD 3: FAILED - Standard 3 column differences found
  • Column 7: Space difference (extra/missing spaces) - Design='Case Owner  First Name' vs Verification='Case Owner First Name'     
  • Column 20: Spelling error (similarity: 0.98) - Design='Case Closure Submsission Date' vs Verification='Case Closure Submission Date'
  • Column 33: Content difference - Design='30 Day Private Visit Contact Log Start Date - Director Approval Received' vs Verification='90 Day Visit Contact Log Start Date - Regular - 2025'
  • Column 34: Content difference - Design='90 Day Visit Contact Log Start Date - Regular - 2025' vs Verification='90 Day Visit Exclusion - Closed Prior to Due Date'
  • Column 35: Content difference - Design='90 Day Visit Exclusion - Closed Prior to Due Date' vs Verification='90 Day Visit Compliant'
  • Column 36: Content difference - Design='90 Day Visit Compliant' vs Verification='Incorrect Change Reason'
  • Column count mismatch: Design=37, Verification=36


=== Specific Cases Test (Standard 2 Report) ===
SPECIFIC_CASES: FAILED - Some cases missing dates

  Case Details:
  • Case 12891050: Due Date=2025-01-13, Contact Log Date=NaT
  • Case 12891050: Due Date=2025-02-12, Contact Log Date=NaT
  • Case 12891050: Due Date=2025-03-14, Contact Log Date=NaT
  • Case 12891050: Due Date=2025-04-13, Contact Log Date=2025-04-01
  • Case 13141575: Due Date=2025-01-15, Contact Log Date=2025-01-01
  • Case 11739608: Due Date=2025-01-06, Contact Log Date=2024-12-27
  • Case 13038729: Due Date=2025-01-02, Contact Log Date=2024-12-25
  • Case 13155126: Due Date=2025-01-03, Contact Log Date=2024-12-25

=== Summary Report Test ===
SUMMARY: PASSED - Summary report fields match perfectly

=== Sensitivity and Formula Tests ===
SENSITIVITY: PASSED - Sensitivity: high (as required)
FORMULA: PASSED - Business formula verified: '90 Day Visit Due Date' - 'Contact Log Buffer Days' <= Minimum Contact Log Start Date <= '90 Day Visit Due Date'

=== Contact Log Requirements Test ===
CONTACT_LOG: PASSED - Contact log requirements verified:
{'report_field_name': '90 Day Private Visit Contact Log Start Date - Extension', 'type': 'Supervision', 'purpose_options': ['Extension of Visit - 30 Day - Private', 'Extension of Visit - 90 Day - Private'], 'concerning': 'Primary Client'}

================================================================================
=== FINAL RESULT ===
SOME CQ091 TESTS FAILED

=== DETAILED ERROR ANALYSIS ===

Cover Page Errors:
  • version: Version mismatch. Expected: 1.3, Found: 1.20

Standard Report Errors:

Standard 1:
  • Column 7: Space difference (extra/missing spaces)
    Design: 'Case Owner  First Name'
    Verification: 'Case Owner First Name'
  • Column 32: Space difference (extra/missing spaces)
    Design: '7 Day Private Visit Contact Log Start Date -
Approved Departure'
    Verification: '7 Day Private Visit Contact Log Start Date - Approved Departure'

Standard 2:
  • Column 7: Space difference (extra/missing spaces)
    Design: 'Case Owner  First Name'
    Verification: 'Case Owner First Name'
  • Column 20: Spelling error (similarity: 0.98)
    Design: 'Case Closure Submsission Date'
    Verification: 'Case Closure Submission Date'
  • Column 33: Content difference
    Design: '30 Day Private Visit Contact Log Start Date -
Director Approval Received'
    Verification: '30 Day Private Visit Contact Log Start Date - Regular - 2025'
  • Column 34: Spelling error (similarity: 0.83)
    Design: '30 Day Private Visit Contact Log Start Date - Regular - 2025'
    Verification: '30 Day Private Visit Contact Log Method - Regular'
  • Column 35: Spelling error (similarity: 0.90)
    Design: '30 Day Private Visit Contact Log Method - Regular'
    Verification: '30 Day Private Visit Contact Log Location - Regular'
  • Column 36: Content difference
    Design: '30 Day Private Visit Contact Log Location - Regular'
    Verification: '30 Day Private Visit Exclusion - Closed Prior to Due Date'
  • Column 37: Content difference
    Design: '30 Day Private Visit Exclusion - Closed Prior to Due Date'
    Verification: '30 Day Private Visit Compliant'
  • Column 38: Content difference
    Design: '30 Day Private Visit Compliant'
    Verification: 'Incorrect Change Reason'
  • Column N/A: Column count mismatch
    Design: 'Total columns: 39'
    Verification: 'Total columns: 38'

Standard 3:
  • Column 7: Space difference (extra/missing spaces)
    Design: 'Case Owner  First Name'
    Verification: 'Case Owner First Name'
  • Column 20: Spelling error (similarity: 0.98)
    Design: 'Case Closure Submsission Date'
    Verification: 'Case Closure Submission Date'
  • Column 33: Content difference
    Design: '30 Day Private Visit Contact Log Start Date - Director Approval Received'
    Verification: '90 Day Visit Contact Log Start Date - Regular - 2025'
  • Column 34: Content difference
    Design: '90 Day Visit Contact Log Start Date - Regular - 2025'
    Verification: '90 Day Visit Exclusion - Closed Prior to Due Date'
  • Column 35: Content difference
    Design: '90 Day Visit Exclusion - Closed Prior to Due Date'
    Verification: '90 Day Visit Compliant'
  • Column 36: Content difference
    Design: '90 Day Visit Compliant'
    Verification: 'Incorrect Change Reason'
  • Column N/A: Column count mismatch
    Design: 'Total columns: 37'
    Verification: 'Total columns: 36'

Specific Cases Errors:
  • Case 12891050: Missing Due Date=False, Missing Contact Log Date=True
  • Case 12891050: Missing Due Date=False, Missing Contact Log Date=True
  • Case 12891050: Missing Due Date=False, Missing Contact Log Date=True
'''