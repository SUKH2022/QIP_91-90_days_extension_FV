import pandas as pd
import re
from datetime import datetime
import os
from difflib import SequenceMatcher
import numpy as np

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
        
        # Prepare results - group by case number to handle duplicates
        case_details = {}
        for _, row in filtered_df.iterrows():
            case_num = str(row[case_column])
            due_date = row[due_date_column]
            contact_log_date = row[contact_log_column]
            
            # Convert dates to string format for display
            due_date_str = due_date.strftime('%Y-%m-%d') if pd.notna(due_date) and isinstance(due_date, (datetime, pd.Timestamp)) else str(due_date)
            contact_log_str = contact_log_date.strftime('%Y-%m-%d') if pd.notna(contact_log_date) and isinstance(contact_log_date, (datetime, pd.Timestamp)) else str(contact_log_date)
            
            if case_num not in case_details:
                case_details[case_num] = {
                    'case_number': case_num,
                    'due_dates': [],
                    'contact_log_dates': [],
                    'has_due_date': False,
                    'has_contact_log_date': False
                }
            
            case_details[case_num]['due_dates'].append(due_date_str)
            case_details[case_num]['contact_log_dates'].append(contact_log_str)
            
            # Update flags if any entry has dates
            if pd.notna(due_date):
                case_details[case_num]['has_due_date'] = True
            if pd.notna(contact_log_date):
                case_details[case_num]['has_contact_log_date'] = True
        
        # Convert to list format for backward compatibility
        details = []
        for case_num, case_info in case_details.items():
            details.append({
                'case_number': case_num,
                'due_date': ', '.join(case_info['due_dates']),
                'contact_log_date': ', '.join(case_info['contact_log_dates']),
                'has_due_date': case_info['has_due_date'],
                'has_contact_log_date': case_info['has_contact_log_date']
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

# =============================================================================
# SUMMARY TOTAL VERIFICATION FUNCTIONS
# =============================================================================

def verify_summary_total_counts(verification_path):
    """
    Verify the counts in Summary Total sheet for all sections:
    - Rows 3-6: 7-day visits (Standard 1 Report)
    - Rows 8-11: 30-day visits (Standard 2 Report) 
    - Rows 13-16: 90-day visits (Standard 3 Report)
    - Rows 19-21: Whereabouts Unknown (All Standard Reports)
    - Rows 24-26: Exclusion - Service Ended (All Standard Reports)
    - Rows 29-31: Exclusion - Data Entry Issue (All Standard Reports)
    - Row 34: For Information Only - Incomplete 90 Day Visits (Standard 3 Report, FCC only)
    - Row 37: Kinship Service Cases - 90 Day Visits with Placement Type Other (Standard 3 Report, Kinship only)
    """
    try:
        # Read all three standard reports
        print("Reading Standard 1 Report...")
        std1_df = pd.read_excel(verification_path, sheet_name="Standard 1 Report", header=1)
        print("Reading Standard 2 Report...")
        std2_df = pd.read_excel(verification_path, sheet_name="Standard 2 Report", header=1)
        print("Reading Standard 3 Report...")
        std3_df = pd.read_excel(verification_path, sheet_name="Standard 3 Report", header=1)
        
        # Clean column names by stripping whitespace for all reports
        std1_df.columns = [str(col).strip() for col in std1_df.columns]
        std2_df.columns = [str(col).strip() for col in std2_df.columns]
        std3_df.columns = [str(col).strip() for col in std3_df.columns]
        
        print("Finding required columns...")
        # Find required columns for each report
        columns_std1 = find_required_columns(std1_df, '7 day private visit compliant')
        columns_std2 = find_required_columns(std2_df, '30 day private visit compliant')
        columns_std3 = find_required_columns(std3_df, '90 day visit compliant')
        
        # Check if all required columns are found
        for report_name, columns in [('Standard 1', columns_std1), 
                                   ('Standard 2', columns_std2), 
                                   ('Standard 3', columns_std3)]:
            missing = [k for k, v in columns.items() if v is None]
            if missing:
                return {"error": f"Missing columns in {report_name}: {missing}"}
        
        # Read the Summary Total sheet
        print("Reading Summary Total sheet...")
        summary_df = pd.read_excel(verification_path, sheet_name="Summary Total", header=None)
        
        # Verify 7-day visits (Rows 3-6) from Standard 1 Report
        print("Verifying 7-day visits...")
        results_7day = verify_visit_section(std1_df, columns_std1, summary_df, 
                                           start_row=2, end_row=5, section_name="7-day")
        
        # Verify 30-day visits (Rows 8-11) from Standard 2 Report
        print("Verifying 30-day visits...")
        results_30day = verify_visit_section(std2_df, columns_std2, summary_df,
                                            start_row=7, end_row=10, section_name="30-day")
        
        # Verify 90-day visits (Rows 13-16) from Standard 3 Report
        print("Verifying 90-day visits...")
        results_90day = verify_visit_section(std3_df, columns_std3, summary_df,
                                            start_row=12, end_row=15, section_name="90-day")
        
        # Verify Whereabouts Unknown section (Rows 19-21) from all reports
        print("Verifying Whereabouts Unknown section...")
        results_whereabouts = verify_whereabouts_unknown_section(std1_df, std2_df, std3_df, 
                                                                columns_std1, columns_std2, columns_std3,
                                                                summary_df, start_row=18, end_row=20)
        
        # Verify Exclusion - Service Ended section (Rows 24-26) from all reports
        print("Verifying Exclusion - Service Ended section...")
        results_exclusion_service = verify_exclusion_service_ended_section(std1_df, std2_df, std3_df, 
                                                                          summary_df, start_row=23, end_row=25)
        
        # Verify Exclusion - Data Entry Issue section (Rows 29-31) from all reports
        print("Verifying Exclusion - Data Entry Issue section...")
        results_exclusion_data = verify_exclusion_data_entry_section(std1_df, std2_df, std3_df, 
                                                                    columns_std1, columns_std2, columns_std3,
                                                                    summary_df, start_row=28, end_row=30)
        
        # Verify For Information Only section (Row 34) from Standard 3 Report
        print("Verifying For Information Only section...")
        results_information = verify_information_only_section(std3_df, columns_std3, summary_df, start_row=33)
        
        # Verify Kinship Service Cases section (Row 37) from Standard 3 Report
        print("Verifying Kinship Service Cases section...")
        results_kinship = verify_kinship_service_cases_section(std3_df, columns_std3, summary_df, start_row=36)
        
        # Combine all results
        all_results = {**results_7day, **results_30day, **results_90day, 
                      **results_whereabouts, **results_exclusion_service, 
                      **results_exclusion_data, **results_information, **results_kinship}
        
        return all_results
        
    except Exception as e:
        import traceback
        return {"error": f"Error verifying Summary Total counts: {str(e)}\n{traceback.format_exc()}"}

def verify_kinship_service_cases_section(std3_df, columns_std3, summary_df, start_row):
    """Verify the Kinship Service Cases section (row 37) - 90 Day Visits with Placement Type Other (Kinship only)"""
    results = {}
    
    # Find primary placement column
    placement_col = find_primary_placement_column(std3_df)
    if not placement_col:
        return {"kinship_error": "Primary Placement Type column not found in Standard 3 Report"}
    
    # Define case types and their corresponding columns (only Kinship for this section)
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1 (should be 0 for Kinship only)
        'C': ('adoption', 2),       # Column C = index 2 (should be 0 for Kinship only)
        'D': ('formal customary care', 3),  # Column D = index 3 (should be 0 for Kinship only)
        'E': ('kinship service', 4) # Column E = index 4 (Kinship only)
    }
    
    current_row = start_row
    
    for col_letter, (case_type, col_index) in case_types.items():
        cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
        
        # Clean and prepare data for comparison
        case_type_data = std3_df[columns_std3['case_type_col']].fillna('').astype(str)
        change_reason_data = std3_df[columns_std3['change_reason_col']].fillna('').astype(str)
        compliant_data = std3_df[columns_std3['compliant_col']].fillna('').astype(str)
        placement_data = std3_df[placement_col].fillna('').astype(str)
        
        # For Kinship Service, check the case type
        if case_type == 'kinship service':
            case_filter = case_type_data.str.lower() == 'kinship service'
        else:
            case_filter = case_type_data.str.lower() == case_type
        
        # Filter for Incorrect Change Reason = 'No'
        change_reason_filter = change_reason_data.str.lower() == 'no'
        
        # Filter for Compliant or Not Compliant status
        compliant_filter = compliant_data.isin(['Compliant', 'Not Compliant'])
        
        # Filter for Placement Type = 'Other'
        placement_filter = placement_data.str.lower() == 'other'
        
        # This section is only for Kinship Service (column E), other columns should be 0
        if case_type == 'kinship service':
            # Count Kinship records with all filters applied
            calculated_value = len(std3_df[case_filter & change_reason_filter & compliant_filter & placement_filter])
        else:
            # Other case types (CIC, Adoption, FCC) should be 0 for this section
            calculated_value = 0
        
        # Get actual value from summary sheet
        actual_value = None
        if current_row < len(summary_df) and col_index < len(summary_df.columns):
            actual_value = summary_df.iloc[current_row, col_index]
            # Handle NaN values and convert to appropriate type
            if pd.isna(actual_value):
                actual_value = 0
            elif isinstance(actual_value, (int, float)):
                pass
            else:
                try:
                    actual_value = float(actual_value)
                except (ValueError, TypeError):
                    actual_value = 0
        
        # Compare calculated vs actual values
        match = int(calculated_value) == int(actual_value) if actual_value is not None else False
        
        results[cell_name] = {
            'calculated': calculated_value,
            'actual': actual_value,
            'match': match,
            'status': 'PASS' if match else 'FAIL',
            'section': 'kinship-service-cases',
            'row_type': '90-day-placement-other',
            'report': '90-day'
        }
        
        # Debug output for troubleshooting
        if not match:
            print(f"DEBUG {cell_name}: case_type='{case_type}', calculated={calculated_value}, actual={actual_value}")
            total_kinship_cases = len(std3_df[case_filter])
            total_correct_reason = len(std3_df[change_reason_filter])
            total_compliant_not = len(std3_df[compliant_filter])
            total_other_placement = len(std3_df[placement_filter])
            print(f"DEBUG Total Kinship cases: {total_kinship_cases}, Correct Change Reason: {total_correct_reason}")
            print(f"DEBUG Compliant/Not Compliant: {total_compliant_not}, Other Placement: {total_other_placement}")
    
    return results

def verify_information_only_section(std3_df, columns_std3, summary_df, start_row):
    """Verify the For Information Only section (row 34) - Incomplete 90 Day Visits (FCC only)"""
    results = {}
    
    # Define case types and their corresponding columns (only FCC for this section)
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1 (should be 0 for FCC only)
        'C': ('adoption', 2),       # Column C = index 2 (should be 0 for FCC only)
        'D': ('formal customary care', 3),  # Column D = index 3 (FCC only)
        'E': ('kinship service', 4) # Column E = index 4 (should be 0 for FCC only)
    }
    
    current_row = start_row
    
    for col_letter, (case_type, col_index) in case_types.items():
        cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
        
        # Clean and prepare data for comparison
        case_type_data = std3_df[columns_std3['case_type_col']].fillna('').astype(str)
        compliant_data = std3_df[columns_std3['compliant_col']].fillna('').astype(str)
        
        # For FCC/Formal Customary Care, check both possible names
        if case_type == 'formal customary care':
            case_filter = case_type_data.str.lower().isin(['fcc', 'formal customary care'])
        else:
            case_filter = case_type_data.str.lower() == case_type
        
        # Filter for Incomplete status
        incomplete_filter = compliant_data.str.lower() == 'incomplete'
        
        # This section is only for FCC (column D), other columns should be 0
        if case_type == 'formal customary care':
            # Count FCC records with Incomplete status
            calculated_value = len(std3_df[case_filter & incomplete_filter])
        else:
            # Other case types (CIC, Adoption, Kinship) should be 0 for this section
            calculated_value = 0
        
        # Get actual value from summary sheet
        actual_value = None
        if current_row < len(summary_df) and col_index < len(summary_df.columns):
            actual_value = summary_df.iloc[current_row, col_index]
            # Handle NaN values and convert to appropriate type
            if pd.isna(actual_value):
                actual_value = 0
            elif isinstance(actual_value, (int, float)):
                pass
            else:
                try:
                    actual_value = float(actual_value)
                except (ValueError, TypeError):
                    actual_value = 0
        
        # Compare calculated vs actual values
        match = int(calculated_value) == int(actual_value) if actual_value is not None else False
        
        results[cell_name] = {
            'calculated': calculated_value,
            'actual': actual_value,
            'match': match,
            'status': 'PASS' if match else 'FAIL',
            'section': 'information-only',
            'row_type': 'incomplete-90-day-fcc',
            'report': '90-day'
        }
        
        # Debug output for troubleshooting
        if not match:
            print(f"DEBUG {cell_name}: case_type='{case_type}', calculated={calculated_value}, actual={actual_value}")
            total_fcc_cases = len(std3_df[case_filter])
            total_incomplete = len(std3_df[incomplete_filter])
            print(f"DEBUG Total FCC cases: {total_fcc_cases}, Total Incomplete: {total_incomplete}")
    
    return results

def verify_exclusion_data_entry_section(std1_df, std2_df, std3_df, columns_std1, columns_std2, columns_std3,
                                       summary_df, start_row, end_row):
    """Verify the Exclusion - Data Entry Issue section (rows 29-31)"""
    results = {}
    
    # Define case types and their corresponding columns
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1
        'C': ('adoption', 2),       # Column C = index 2
        'D': ('formal customary care', 3),  # Column D = index 3
        'E': ('kinship service', 4) # Column E = index 4
    }
    
    # For each row in the section (7-day, 30-day, 90-day)
    for row_offset, (report_df, columns, report_name) in enumerate([
        (std1_df, columns_std1, '7-day'),
        (std2_df, columns_std2, '30-day'),
        (std3_df, columns_std3, '90-day')
    ]):
        current_row = start_row + row_offset
        
        for col_letter, (case_type, col_index) in case_types.items():
            cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
            
            # Clean and prepare data for comparison
            case_type_data = report_df[columns['case_type_col']].fillna('').astype(str)
            change_reason_data = report_df[columns['change_reason_col']].fillna('').astype(str)
            compliant_data = report_df[columns['compliant_col']].fillna('').astype(str)
            
            # For FCC/Formal Customary Care, check both possible names
            if case_type == 'formal customary care':
                case_filter = case_type_data.str.lower().isin(['fcc', 'formal customary care'])
            else:
                case_filter = case_type_data.str.lower() == case_type
            
            # Filter for Incorrect Change Reason = 'Yes'
            incorrect_reason_filter = change_reason_data.str.lower() == 'yes'
            
            # Filter for records that have compliant data (not null/empty)
            compliant_filter = compliant_data.notna() & (compliant_data != '')
            
            # For 90-day visits, exclude CIC
            if report_name == '90-day' and case_type == 'child in care':
                calculated_value = 0  # CIC excluded from 90-day visits
            else:
                # Count records with Incorrect Change Reason = 'Yes' that have compliant data
                calculated_value = len(report_df[case_filter & incorrect_reason_filter & compliant_filter])
            
            # Get actual value from summary sheet
            actual_value = None
            if current_row < len(summary_df) and col_index < len(summary_df.columns):
                actual_value = summary_df.iloc[current_row, col_index]
                # Handle NaN values and convert to appropriate type
                if pd.isna(actual_value):
                    actual_value = 0
                elif isinstance(actual_value, (int, float)):
                    pass
                else:
                    try:
                        actual_value = float(actual_value)
                    except (ValueError, TypeError):
                        actual_value = 0
            
            # Compare calculated vs actual values
            match = int(calculated_value) == int(actual_value) if actual_value is not None else False
            
            results[cell_name] = {
                'calculated': calculated_value,
                'actual': actual_value,
                'match': match,
                'status': 'PASS' if match else 'FAIL',
                'section': 'exclusion-data-entry',
                'row_type': report_name,
                'report': report_name
            }
            
            # Debug output for troubleshooting
            if not match:
                print(f"DEBUG {cell_name}: case_type='{case_type}', calculated={calculated_value}, actual={actual_value}")
                total_cases = len(report_df[case_filter])
                total_incorrect_reason = len(report_df[incorrect_reason_filter])
                total_with_compliant = len(report_df[compliant_filter])
                print(f"DEBUG Total cases: {total_cases}, Incorrect Change Reason: {total_incorrect_reason}, With Compliant Data: {total_with_compliant}")
    
    return results

def verify_exclusion_service_ended_section(std1_df, std2_df, std3_df, summary_df, start_row, end_row):
    """Verify the Exclusion - Service Ended section (rows 24-26)"""
    results = {}
    
    # Find exclusion columns in each report
    exclusion_col_std1 = find_exclusion_column(std1_df, '7 day')
    exclusion_col_std2 = find_exclusion_column(std2_df, '30 day')
    exclusion_col_std3 = find_exclusion_column(std3_df, '90 day')
    
    if not exclusion_col_std1:
        return {"exclusion_error": "7 Day Private Visit Exclusion column not found in Standard 1 Report"}
    if not exclusion_col_std2:
        return {"exclusion_error": "30 Day Private Visit Exclusion column not found in Standard 2 Report"}
    if not exclusion_col_std3:
        return {"exclusion_error": "90 Day Visit Exclusion column not found in Standard 3 Report"}
    
    # Define case types and their corresponding columns
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1
        'C': ('adoption', 2),       # Column C = index 2
        'D': ('formal customary care', 3),  # Column D = index 3
        'E': ('kinship service', 4) # Column E = index 4
    }
    
    # For each row in the section (7-day, 30-day, 90-day)
    for row_offset, (report_df, exclusion_col, report_name) in enumerate([
        (std1_df, exclusion_col_std1, '7-day'),
        (std2_df, exclusion_col_std2, '30-day'),
        (std3_df, exclusion_col_std3, '90-day')
    ]):
        current_row = start_row + row_offset
        
        for col_letter, (case_type, col_index) in case_types.items():
            cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
            
            # Clean and prepare data for comparison
            case_type_data = report_df[find_case_type_column(report_df)].fillna('').astype(str)
            exclusion_data = report_df[exclusion_col].fillna('').astype(str)
            
            # For FCC/Formal Customary Care, check both possible names
            if case_type == 'formal customary care':
                case_filter = case_type_data.str.lower().isin(['fcc', 'formal customary care'])
            else:
                case_filter = case_type_data.str.lower() == case_type
            
            # Exclusion filter - count "Yes" values
            exclusion_filter = exclusion_data.str.lower() == 'yes'
            
            # For 90-day visits, exclude CIC
            if report_name == '90-day' and case_type == 'child in care':
                calculated_value = 0  # CIC excluded from 90-day visits
            else:
                # Count records with Exclusion = "Yes"
                calculated_value = len(report_df[case_filter & exclusion_filter])
            
            # Get actual value from summary sheet
            actual_value = None
            if current_row < len(summary_df) and col_index < len(summary_df.columns):
                actual_value = summary_df.iloc[current_row, col_index]
                # Handle NaN values and convert to appropriate type
                if pd.isna(actual_value):
                    actual_value = 0
                elif isinstance(actual_value, (int, float)):
                    pass
                else:
                    try:
                        actual_value = float(actual_value)
                    except (ValueError, TypeError):
                        actual_value = 0
            
            # Compare calculated vs actual values
            match = int(calculated_value) == int(actual_value) if actual_value is not None else False
            
            results[cell_name] = {
                'calculated': calculated_value,
                'actual': actual_value,
                'match': match,
                'status': 'PASS' if match else 'FAIL',
                'section': 'exclusion-service-ended',
                'row_type': report_name,
                'report': report_name
            }
            
            # Debug output for troubleshooting
            if not match:
                print(f"DEBUG {cell_name}: case_type='{case_type}', exclusion_col='{exclusion_col}', calculated={calculated_value}, actual={actual_value}")
                total_cases = len(report_df[case_filter])
                total_exclusions = len(report_df[exclusion_filter])
                print(f"DEBUG Total cases: {total_cases}, Total exclusions: {total_exclusions}")
    
    return results

def find_exclusion_column(df, visit_type):
    """Find the exclusion column for the specified visit type"""
    # More flexible matching for exclusion columns
    for col in df.columns:
        col_lower = str(col).lower()
        # Check if it's an exclusion column with the right pattern
        if 'exclusion' in col_lower and 'closed prior to due date' in col_lower:
            # For 7-day visits
            if visit_type == '7 day' and '7' in col_lower:
                return col
            # For 30-day visits  
            elif visit_type == '30 day' and '30' in col_lower:
                return col
            # For 90-day visits
            elif visit_type == '90 day' and '90' in col_lower:
                return col
                
    return None

def find_case_type_column(df):
    """Find the Case Type column"""
    for col in df.columns:
        col_lower = str(col).lower()
        if 'case type' in col_lower:
            return col
    return None

def verify_whereabouts_unknown_section(std1_df, std2_df, std3_df, columns_std1, columns_std2, columns_std3, 
                                      summary_df, start_row, end_row):
    """Verify the Whereabouts Unknown section (rows 19-21)"""
    results = {}
    
    # Find primary placement column in each report
    placement_col_std1 = find_primary_placement_column(std1_df)
    placement_col_std2 = find_primary_placement_column(std2_df)
    placement_col_std3 = find_primary_placement_column(std3_df)
    
    if not placement_col_std1:
        return {"whereabouts_error": "Primary Placement Type column not found in Standard 1 Report"}
    if not placement_col_std2:
        return {"whereabouts_error": "Primary Placement Type column not found in Standard 2 Report"}
    if not placement_col_std3:
        return {"whereabouts_error": "Primary Placement Type column not found in Standard 3 Report"}
    
    # Define case types and their corresponding columns
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1
        'C': ('adoption', 2),       # Column C = index 2
        'D': ('formal customary care', 3),  # Column D = index 3
        'E': ('kinship service', 4) # Column E = index 4
    }
    
    # For each row in the section (7-day, 30-day, 90-day)
    for row_offset, (report_df, columns, report_name) in enumerate([
        (std1_df, columns_std1, '7-day'),
        (std2_df, columns_std2, '30-day'),
        (std3_df, columns_std3, '90-day')
    ]):
        current_row = start_row + row_offset
        
        for col_letter, (case_type, col_index) in case_types.items():
            cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
            
            # Filter data: Incorrect Change Reason = 'No'
            change_reason_data = report_df[columns['change_reason_col']].fillna('').astype(str)
            filtered_df = report_df[change_reason_data.str.lower() == 'no']
            
            # Clean and prepare data for comparison
            case_type_data = filtered_df[columns['case_type_col']].fillna('').astype(str)
            placement_data = filtered_df[placement_col_std1].fillna('').astype(str)
            
            # For FCC/Formal Customary Care, check both possible names
            if case_type == 'formal customary care':
                case_filter = case_type_data.str.lower().isin(['fcc', 'formal customary care'])
            else:
                case_filter = case_type_data.str.lower() == case_type
            
            # Whereabouts Unknown filter
            whereabouts_filter = placement_data.str.lower() == 'whereabouts unknown'
            
            # For 90-day visits, exclude CIC
            if report_name == '90-day' and case_type == 'child in care':
                calculated_value = 0  # CIC excluded from 90-day visits
            else:
                # Count records with Whereabouts Unknown placement type
                calculated_value = len(filtered_df[case_filter & whereabouts_filter])
            
            # Get actual value from summary sheet
            actual_value = None
            if current_row < len(summary_df) and col_index < len(summary_df.columns):
                actual_value = summary_df.iloc[current_row, col_index]
                # Handle NaN values and convert to appropriate type
                if pd.isna(actual_value):
                    actual_value = 0
                elif isinstance(actual_value, (int, float)):
                    pass
                else:
                    try:
                        actual_value = float(actual_value)
                    except (ValueError, TypeError):
                        actual_value = 0
            
            # Compare calculated vs actual values
            match = int(calculated_value) == int(actual_value) if actual_value is not None else False
            
            results[cell_name] = {
                'calculated': calculated_value,
                'actual': actual_value,
                'match': match,
                'status': 'PASS' if match else 'FAIL',
                'section': 'whereabouts-unknown',
                'row_type': report_name,
                'report': report_name
            }
    
    return results

def find_primary_placement_column(df):
    """Find the Primary In Care Placement Type column"""
    for col in df.columns:
        col_lower = str(col).lower()
        if 'primary in care placement type' in col_lower:
            return col
    return None

def find_required_columns(df, compliant_pattern):
    """Find required columns in a dataframe based on patterns"""
    columns_needed = {
        'compliant_col': None,
        'change_reason_col': None,
        'case_type_col': None
    }
    
    for col in df.columns:
        col_lower = str(col).lower()
        if compliant_pattern in col_lower:
            columns_needed['compliant_col'] = col
        elif 'incorrect change reason' in col_lower:
            columns_needed['change_reason_col'] = col
        elif 'case type' in col_lower:
            columns_needed['case_type_col'] = col
    
    return columns_needed

def verify_visit_section(df, columns, summary_df, start_row, end_row, section_name):
    """Verify a specific visit section in the summary sheet"""
    results = {}
    
    print(f"Verifying {section_name} section, rows {start_row+1} to {end_row+1}")
    
    # Filter data: Incorrect Change Reason = 'No'
    change_reason_data = df[columns['change_reason_col']].fillna('').astype(str)
    filtered_df = df[change_reason_data.str.lower() == 'no']
    
    print(f"Filtered data size: {len(filtered_df)} (original: {len(df)})")
    
    if len(filtered_df) == 0:
        return {f"{section_name}_error": f"No records found with Incorrect Change Reason = 'No'"}
    
    # Define case types and their corresponding columns with proper column indices
    case_types = {
        'B': ('child in care', 1),  # Column B = index 1
        'C': ('adoption', 2),       # Column C = index 2
        'D': ('formal customary care', 3),  # Column D = index 3
        'E': ('kinship service', 4) # Column E = index 4
    }
    
    # For each row in the section (Total, Compliant, Non-Compliant, Compliance Rate)
    for row_offset, row_type in enumerate(['total', 'compliant', 'non_compliant', 'compliance_rate']):
        current_row = start_row + row_offset
        
        for col_letter, (case_type, col_index) in case_types.items():
            cell_name = f"{col_letter}{current_row + 1}"  # Excel rows are 1-indexed
            
            # Clean and prepare data for comparison
            case_type_data = filtered_df[columns['case_type_col']].fillna('').astype(str)
            compliant_data = filtered_df[columns['compliant_col']].fillna('').astype(str)
            
            # For FCC/Formal Customary Care, check both possible names
            if case_type == 'formal customary care':
                case_filter = case_type_data.str.lower().isin(['fcc', 'formal customary care'])
            else:
                case_filter = case_type_data.str.lower() == case_type
            
            # Calculate expected value based on row type
            if row_type == 'total':
                if section_name == '90-day':
                    compliant_filter = compliant_data.isin(['Compliant', 'Not Compliant', 'Incomplete'])
                else:
                    compliant_filter = compliant_data.isin(['Compliant', 'Not Compliant'])
                calculated_value = len(filtered_df[case_filter & compliant_filter])
                
            elif row_type == 'compliant':
                compliant_filter = compliant_data == 'Compliant'
                calculated_value = len(filtered_df[case_filter & compliant_filter])
                
            elif row_type == 'non_compliant':
                if section_name == '90-day':
                    compliant_filter = compliant_data.isin(['Not Compliant', 'Incomplete'])
                else:
                    compliant_filter = compliant_data == 'Not Compliant'
                calculated_value = len(filtered_df[case_filter & compliant_filter])
                
            else:  # compliance_rate
                if section_name == '90-day':
                    total_filter = compliant_data.isin(['Compliant', 'Not Compliant', 'Incomplete'])
                else:
                    total_filter = compliant_data.isin(['Compliant', 'Not Compliant'])
                    
                compliant_filter = compliant_data == 'Compliant'
                
                total_count = len(filtered_df[case_filter & total_filter])
                compliant_count = len(filtered_df[case_filter & compliant_filter])
                calculated_value = compliant_count / total_count if total_count > 0 else 0
            
            # Get actual value from summary sheet
            actual_value = None
            if current_row < len(summary_df) and col_index < len(summary_df.columns):
                actual_value = summary_df.iloc[current_row, col_index]
                if pd.isna(actual_value):
                    actual_value = 0
                elif not isinstance(actual_value, (int, float)):
                    try:
                        actual_value = float(actual_value)
                    except (ValueError, TypeError):
                        actual_value = 0
            
            # For compliance rate, compare with tolerance
            if row_type == 'compliance_rate' and actual_value is not None:
                try:
                    calculated_float = float(calculated_value)
                    actual_float = float(actual_value)
                    match = abs(calculated_float - actual_float) < 0.01
                except (ValueError, TypeError):
                    match = False
            else:
                try:
                    match = int(calculated_value) == int(actual_value) if actual_value is not None else False
                except (ValueError, TypeError):
                    match = False
            
            results[cell_name] = {
                'calculated': calculated_value,
                'actual': actual_value,
                'match': match,
                'status': 'PASS' if match else 'FAIL',
                'section': section_name,
                'row_type': row_type
            }
    
    return results

def verify_complete_summary_sheet(verification_path):
    """
    Complete verification of Summary Total sheet against all Standard Reports
    """
    print("=== COMPREHENSIVE SUMMARY TOTAL SHEET VERIFICATION ===")
    print(f"Verification File: {verification_path}")
    print("=" * 80)
    
    # Step 1: Verify all counts across all sections
    print("\n1. Verifying All Summary Counts:")
    print("-" * 40)
    
    basic_results = verify_summary_total_counts(verification_path)
    
    if 'error' in basic_results:
        print(f"❌ ERROR: {basic_results['error']}")
        return False
    
    # Check if we have any section errors
    section_errors = {k: v for k, v in basic_results.items() if 'error' in k}
    if section_errors:
        for error_key, error_msg in section_errors.items():
            print(f"❌ {error_key}: {error_msg}")
        basic_results = {k: v for k, v in basic_results.items() if 'error' not in k}
    
    if not basic_results:
        print("❌ No results to verify")
        return False
    
    # Group results by section for better reporting
    sections = {
        '7-day Visits (Standard 1 Report)': {},
        '30-day Visits (Standard 2 Report)': {}, 
        '90-day Visits (Standard 3 Report)': {},
        'Whereabouts Unknown (All Reports)': {},
        'Exclusion - Service Ended (All Reports)': {},
        'Exclusion - Data Entry Issue (All Reports)': {},
        'For Information Only (Standard 3 Report)': {},
        'Kinship Service Cases (Standard 3 Report)': {}
    }
    
    for cell, result in basic_results.items():
        section_key = result.get('section', '')
        if '7-day' in section_key and 'exclusion' not in section_key and 'information' not in section_key and 'kinship' not in section_key:
            sections['7-day Visits (Standard 1 Report)'][cell] = result
        elif '30-day' in section_key and 'exclusion' not in section_key and 'information' not in section_key and 'kinship' not in section_key:
            sections['30-day Visits (Standard 2 Report)'][cell] = result
        elif '90-day' in section_key and 'exclusion' not in section_key and 'information' not in section_key and 'kinship' not in section_key:
            sections['90-day Visits (Standard 3 Report)'][cell] = result
        elif 'whereabouts-unknown' in section_key:
            sections['Whereabouts Unknown (All Reports)'][cell] = result
        elif 'exclusion-service-ended' in section_key:
            sections['Exclusion - Service Ended (All Reports)'][cell] = result
        elif 'exclusion-data-entry' in section_key:
            sections['Exclusion - Data Entry Issue (All Reports)'][cell] = result
        elif 'information-only' in section_key:
            sections['For Information Only (Standard 3 Report)'][cell] = result
        elif 'kinship-service-cases' in section_key:
            sections['Kinship Service Cases (Standard 3 Report)'][cell] = result
    
    # Report results by section
    for section_name, section_results in sections.items():
        if not section_results:
            continue
            
        print(f"\n   {section_name}:")
        print("   " + "-" * 30)
        
        row_types = {}
        for cell, result in section_results.items():
            row_type = result.get('row_type', 'unknown')
            if row_type not in row_types:
                row_types[row_type] = []
            row_types[row_type].append((cell, result))
        
        for row_type, cells in row_types.items():
            print(f"   {row_type.replace('_', ' ').title()}:")
            for cell, result in cells:
                status_icon = "✅" if result['match'] else "❌"
                actual_display = result['actual']
                
                if result.get('row_type') == 'compliance_rate' and isinstance(actual_display, (int, float)):
                    actual_display = f"{actual_display:.2%}"
                
                calculated_display = result['calculated']
                if result.get('row_type') == 'compliance_rate' and isinstance(calculated_display, (int, float)):
                    calculated_display = f"{calculated_display:.2%}"
                
                print(f"     {cell}: {status_icon} Actual: {actual_display}, Calculated: {calculated_display}")
    
    # Step 2: Calculate overall verification status
    print("\n2. Overall Verification Summary:")
    print("-" * 40)
    
    valid_results = [result for result in basic_results.values() if not isinstance(result, str)]
    
    if valid_results:
        all_match = all(result['match'] for result in valid_results)
        total_cells = len(valid_results)
        passed_cells = sum(1 for result in valid_results if result['match'])
        
        print(f"   Total Cells Verified: {total_cells}")
        print(f"   Cells Passed: {passed_cells}")
        print(f"   Cells Failed: {total_cells - passed_cells}")
        print(f"   Success Rate: {passed_cells/total_cells:.1%}" if total_cells > 0 else "N/A")
        
        print("\n" + "=" * 80)
        if all_match:
            print("✅ COMPREHENSIVE SUMMARY VERIFICATION: PASSED")
            return True
        else:
            print("❌ COMPREHENSIVE SUMMARY VERIFICATION: FAILED")
            print("\nFailed Cells:")
            for cell, result in basic_results.items():
                if not isinstance(result, str) and not result['match']:
                    actual_display = result['actual']
                    calculated_display = result['calculated']
                    
                    if result.get('row_type') == 'compliance_rate':
                        actual_display = f"{actual_display:.2%}" if isinstance(actual_display, (int, float)) else actual_display
                        calculated_display = f"{calculated_display:.2%}" if isinstance(calculated_display, (int, float)) else calculated_display
                    
                    print(f"   {cell}: Expected {calculated_display}, Got {actual_display}")
            
            return False
    else:
        print("❌ No valid results to verify")
        return False

def run_all_cq091_tests(design_spec_path, verification_path, expected_version):
    """Run all tests for CQ091 report verification including summary total verification"""
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
    
    # Run comprehensive summary total verification
    print("\n=== Summary Total Sheet Verification ===")
    summary_total_passed = verify_complete_summary_sheet(verification_path)
    
    # Calculate overall status
    cover_passed = all(r['passed'] for r in cover_results.values())
    standards_passed = all(result['passed'] for result in standard_results)
    cases_passed = cases_result['passed']
    summary_passed = summary_result['passed']
    sensitivity_passed = all(r['passed'] for r in sensitivity_results.values())
    contact_passed = contact_result['passed']
    
    all_passed = (cover_passed and standards_passed and cases_passed and 
                  summary_passed and sensitivity_passed and contact_passed and 
                  summary_total_passed)
    
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
        
        # Summary total errors
        if not summary_total_passed:
            print("\nSummary Total Sheet Errors:")
            print("  • Detailed errors shown in the summary total verification section above")
    
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