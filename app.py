import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta
import numpy as np
import os
import zipfile

st.set_page_config(page_title="Excel File Validator", layout="wide")


def main():
    st.title("Excel File Validator")

    # File Upload Section
    st.header("Upload Files")

    col1, col2 = st.columns(2)

    with col1:
        template_file = st.file_uploader(
            "Upload Template Excel File", type=["xlsx", "xls"])

    with col2:
        data_files = st.file_uploader(
            "Upload Data Excel File to Validate", type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="You can upload multiple data files to validate against the same template")

    if template_file and data_files:
        # Store uploaded files in session state for tab access
        if 'template_file' not in st.session_state:
            st.session_state.template_file = template_file
            st.session_state.data_files = data_files

        # Show uploaded files info
        st.info(f"📁 Template: {template_file.name}")
        st.info(f"📊 Data Files: {len(data_files)} file(s) uploaded")

        # File selector if multiple files
        if len(data_files) > 1:
            st.write("### Select Data File to Validate")
            file_options = [f.name for f in data_files]
            selected_file_name = st.selectbox(
                "Choose a data file:", file_options)
            selected_data_file = next(
                f for f in data_files if f.name == selected_file_name)
        else:
            selected_data_file = data_files[0]
            st.success(f"✓ Validating: {selected_data_file.name}")

        # Create tabs for each validation step
        tabs = st.tabs([
            "1. Sheet Validation",
            "2. Column & Cell Validation",
            "3. Row Count Validation",
            "4. Data Validation",
            "5. Find Missing Data",
            "6. Export Data to CSV"
        ])

        # Load Excel files
        try:
            template_data = load_excel_file(template_file)
            data_to_validate = load_excel_file(selected_data_file)

            # Tab 1: Sheet Validation
            with tabs[0]:
                validate_sheets(template_data, data_to_validate)

            # Tab 2: Column & Cell Validation
            with tabs[1]:
                validate_columns_and_cells(template_data, data_to_validate)

            # Tab 3: Row Count Validation (placeholder for now)
            with tabs[2]:
                validate_row_counts(template_data, data_to_validate)

            # Tab 4: Data Validation (placeholder for now)
            with tabs[3]:
                validate_data(template_data, data_to_validate)

            # Tab 5: Find Missing Data (placeholder for now)
            with tabs[4]:
                find_missing_data(data_to_validate)

            # Tab 6: Export Data to CSV
            with tabs[5]:
                export_data_to_csv_tab(template_data, data_to_validate)

        except Exception as e:
            st.error(f"Error loading Excel files: {e}")
    else:
        st.info("Please upload both template and data file(s) to begin validation")


def load_excel_file(file):
    """Load all sheets from an Excel file into a dictionary of dataframes"""
    excel_file = pd.ExcelFile(file)
    sheet_names = excel_file.sheet_names

    sheet_data = {}
    for sheet in sheet_names:
        sheet_data[sheet] = pd.read_excel(excel_file, sheet_name=sheet)

    return {
        'file_name': file.name,
        'sheet_names': sheet_names,
        'sheet_data': sheet_data
    }


def validate_sheets(template_data, data_to_validate):
    """Validate that all sheets in template exist in the data file"""
    st.subheader("Sheet Validation")

    template_sheets = set(template_data['sheet_names'])
    data_sheets = set(data_to_validate['sheet_names'])

    # Check if all template sheets exist in data file
    missing_sheets = template_sheets - data_sheets
    extra_sheets = data_sheets - template_sheets

    col1, col2 = st.columns(2)

    with col1:
        st.write(f"Template File: **{template_data['file_name']}**")
        st.write(
            f"Template Sheets: {', '.join([str(s) for s in template_data['sheet_names']])}")

    with col2:
        st.write(f"Data File: **{data_to_validate['file_name']}**")
        st.write(
            f"Data Sheets: {', '.join([str(s) for s in data_to_validate['sheet_names']])}")

    if not missing_sheets and not extra_sheets:
        st.success(
            "All sheets match! The data file contains all the sheets from the template file.")
    else:
        if missing_sheets:
            st.error(
                f"Missing sheets in data file: {', '.join([str(s) for s in missing_sheets])}")
        if extra_sheets:
            st.warning(
                f"Extra sheets in data file (not in template): {', '.join([str(s) for s in extra_sheets])}")

    # Display sheet comparison table
    comparison_data = []
    for sheet in sorted(template_sheets.union(data_sheets)):
        comparison_data.append({
            "Sheet Name": sheet,
            "In Template": "Y" if sheet in template_sheets else "N",
            "In Data File": "Y" if sheet in data_sheets else "N",
            "Status": "Match" if (sheet in template_sheets and sheet in data_sheets) else "Missing"
        })

    st.subheader("Sheet Comparison")
    st.dataframe(pd.DataFrame(comparison_data), use_container_width=True)


def validate_columns_and_cells(template_data, data_to_validate):
    """
    Validate cell values row by row across columns.
    For rows 5,6,7: checks A5-Z5, A6-Z6, A7-Z7 (where Z is last column with data)
    """
    st.subheader("Column & Cell Validation")

    # Add UI for selecting which rows to validate
    st.write("### Configure Validation")

    col1, col2 = st.columns(2)

    with col1:
        validation_mode = st.radio(
            "Validation Mode",
            ["Specific Rows", "All Rows"],
            index=0,
            help="Choose whether to validate specific rows or all rows"
        )

    rows_to_validate = []
    if validation_mode == "Specific Rows":
        with col2:
            row_input = st.text_input(
                "Rows to Validate (1-based, e.g. 5,6,7)",
                value="5,6,7",
                help="Enter the row numbers to validate (1-based indexing, where row 1 is the first row)"
            )

            try:
                # Convert to 0-based indexing
                rows_to_validate = [
                    int(r.strip()) - 1 for r in row_input.split(",") if r.strip()]
                # Display in 1-based for user clarity
                st.info(
                    f"Will validate rows: {', '.join([str(r+1) for r in rows_to_validate])}")
            except ValueError:
                st.error(
                    "Invalid row numbers. Please enter comma-separated integers.")
                return

    template_sheets = set(template_data['sheet_names'])
    data_sheets = set(data_to_validate['sheet_names'])

    # Find common sheets to validate
    common_sheets = template_sheets.intersection(data_sheets)

    if not common_sheets:
        st.error("No common sheets found to validate!")
        return

    all_valid = True

    for sheet in sorted(common_sheets):
        template_df = template_data['sheet_data'][sheet]
        data_df = data_to_validate['sheet_data'][sheet]

        # Check if number of columns match
        template_cols = list(template_df.columns)
        data_cols = list(data_df.columns)
        columns_match = len(template_cols) == len(data_cols)

        # Track mismatches
        text_cell_mismatches = []
        validated_data = []

        # Determine which rows to validate
        if validation_mode == "All Rows":
            rows_to_check = range(len(template_df))
        else:
            rows_to_check = rows_to_validate

        # Only validate rows that exist in both dataframes
        valid_rows = [row_idx for row_idx in rows_to_check
                      if row_idx < len(template_df) and row_idx < len(data_df)]

        if columns_match:
            # For each row, check all columns horizontally
            for row_idx in valid_rows:
                # Find the last column with data in this row
                last_col_with_data = -1
                for col_idx in range(len(template_cols)):
                    template_value = template_df.iloc[row_idx, col_idx]
                    if not pd.isna(template_value) and str(template_value).strip() != "":
                        last_col_with_data = col_idx

                # Only process if there's data in this row
                if last_col_with_data >= 0:
                    # Check all columns up to the last one with data
                    for col_idx in range(last_col_with_data + 1):
                        col_name = template_cols[col_idx]

                        # Get cell values from both files
                        template_value = template_df.iloc[row_idx, col_idx]
                        data_value = data_df.iloc[row_idx, col_idx]

                        # Convert both to strings for comparison
                        template_str = str(template_value).strip(
                        ) if not pd.isna(template_value) else ""
                        data_str = str(data_value).strip(
                        ) if not pd.isna(data_value) else ""

                        # Store validated data
                        validated_data.append({
                            "Row": row_idx + 1,  # 1-based row number
                            "Column": col_name,
                            "Template Value": template_str,
                            "Data Value": data_str,
                            "Match": "Yes" if template_str == data_str else "No"
                        })

                        # Check if values match
                        if template_str != data_str:
                            # Convert to Excel cell reference (A1, B2, etc.)
                            cell_addr = get_excel_column_name(
                                col_idx) + str(row_idx + 1)

                            text_cell_mismatches.append({
                                "Cell": cell_addr,
                                "Row": row_idx + 1,
                                "Column": col_name,
                                "Template Value": template_str,
                                "Data Value": data_str
                            })

        # Determine sheet validation status
        sheet_valid = columns_match and not text_cell_mismatches
        all_valid = all_valid and sheet_valid

        # Create expandable section for sheet details
        with st.expander(f"Sheet: {sheet} - {'✓ Valid' if sheet_valid else '✗ Invalid'}"):
            if not columns_match:
                st.error(
                    f"Column count mismatch: Template has {len(template_cols)} columns, but data file has {len(data_cols)} columns")
            else:
                st.success(
                    f"✓ Column count matches: Both files have {len(template_cols)} columns")

            # Show validated data
            if validated_data:
                st.write(
                    f"**Validated Data ({len(validated_data)} cells checked):**")
                validated_df = pd.DataFrame(validated_data)

                # Apply styling to show matches/mismatches
                def highlight_match(row):
                    if row['Match'] == 'No':
                        return ['background-color: #ffc7ce; color: #9c0006'] * len(row)
                    else:
                        return ['background-color: #c6efce; color: #006100'] * len(row)

                styled_df = validated_df.style.apply(highlight_match, axis=1)
                st.dataframe(styled_df, use_container_width=True)

            # Text cell comparison
            if text_cell_mismatches:
                st.error(
                    f"✗ Found {len(text_cell_mismatches)} cell mismatches")
                mismatches_df = pd.DataFrame(text_cell_mismatches)
                st.dataframe(mismatches_df, use_container_width=True)
            else:
                st.success("✓ All cell values match")

    # Show overall status
    st.write("---")
    if all_valid:
        st.success("✓ All sheets passed validation!")
    else:
        st.error("✗ Some sheets failed validation. See details above.")

    return all_valid


def get_excel_column_name(col_idx):
    """Convert 0-based column index to Excel column name (A, B, C, ..., Z, AA, AB, ...)"""
    result = ""
    while True:
        col_idx, remainder = divmod(col_idx, 26)
        result = chr(65 + remainder) + result
        if col_idx == 0:
            break
        col_idx -= 1
    return result


def validate_row_counts(template_data, data_to_validate):
    """
    Validate that each sheet in the data file has the correct number of rows based on the month
    specified in cell A7 (row with index 6).

    The expected row count is: (minutes in month) + 7 header rows
    For example: January = 1 * 60 * 24 * 31 + 7 = 44647 rows
    """
    st.subheader("Row Count Validation")

    # Define days in each month (considering leap years)
    def is_leap_year(year):
        return (year % 4 == 0 and year % 100 != 0) or (year % 400 == 0)

    def get_days_in_month(month, year):
        days_in_month = {
            1: 31,  # January
            2: 29 if is_leap_year(year) else 28,  # February
            3: 31,  # March
            4: 30,  # April
            5: 31,  # May
            6: 30,  # June
            7: 31,  # July
            8: 31,  # August
            9: 30,  # September
            10: 31,  # October
            11: 30,  # November
            12: 31   # December
        }
        # Default to 30 if month is unknown
        return days_in_month.get(month, 30)

    # Get common sheets between template and data file
    template_sheets = set(template_data['sheet_names'])
    data_sheets = set(data_to_validate['sheet_names'])
    common_sheets = template_sheets.intersection(data_sheets)

    if not common_sheets:
        st.error("No common sheets found between template and data file.")
        return

    # Create a comparison table for row counts
    comparison_data = []

    # Check each common sheet
    for sheet_name in sorted(common_sheets):
        template_sheet = template_data['sheet_data'][sheet_name]
        data_sheet = data_to_validate['sheet_data'][sheet_name]

        try:
            # Get month from cell A7 (row with index 6)
            if len(data_sheet) >= 7:  # Make sure sheet has at least 7 rows
                # A7 cell (zero-indexed)
                date_cell_value = data_sheet.iloc[6, 0]

                # Handle different date formats
                if pd.isna(date_cell_value):
                    # If A7 is empty, show warning
                    st.warning(f"Sheet '{sheet_name}': Date cell A7 is empty.")
                    month_days = None
                    expected_rows = None
                else:
                    # Convert the date cell to datetime if it's not already
                    if isinstance(date_cell_value, str):
                        try:
                            date_value = pd.to_datetime(date_cell_value)
                        except:
                            st.warning(
                                f"Sheet '{sheet_name}': Unable to parse date from A7: {date_cell_value}")
                            date_value = None
                    else:
                        date_value = pd.to_datetime(date_cell_value) if not pd.isna(
                            date_cell_value) else None

                    if date_value:
                        month = date_value.month
                        year = date_value.year
                        month_days = get_days_in_month(month, year)
                        # Expected rows = minutes in month + 7 header rows
                        expected_rows = month * 60 * 24 * month_days + 7
                    else:
                        month_days = None
                        expected_rows = None
            else:
                st.warning(
                    f"Sheet '{sheet_name}': Sheet has fewer than 7 rows.")
                month_days = None
                expected_rows = None

            # Get actual row counts
            template_rows = len(template_sheet)
            data_rows = len(data_sheet)

            # Determine status
            if expected_rows is None:
                status = "Unknown"
            elif data_rows == expected_rows:
                status = "Correct"
            else:
                status = "Incorrect"

            # Add to comparison data
            comparison_data.append({
                "Sheet Name": sheet_name,
                "Template Rows": template_rows,
                "Data File Rows": data_rows,
                "Expected Rows": expected_rows if expected_rows is not None else "Unknown",
                "Month Days": month_days if month_days is not None else "Unknown",
                "Status": status
            })

        except Exception as e:
            st.error(f"Error processing sheet '{sheet_name}': {str(e)}")
            comparison_data.append({
                "Sheet Name": sheet_name,
                "Template Rows": len(template_sheet),
                "Data File Rows": len(data_sheet),
                "Expected Rows": "Error",
                "Month Days": "Error",
                "Status": "Error"
            })

    # Display comparison table
    st.dataframe(pd.DataFrame(comparison_data), use_container_width=True)

    # Summary
    correct_sheets = sum(
        1 for item in comparison_data if "Y" in item["Status"])
    incorrect_sheets = sum(
        1 for item in comparison_data if "N" in item["Status"])
    unknown_sheets = sum(
        1 for item in comparison_data if "Unknown" in item["Status"])

    st.subheader("Summary")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Correct Row Counts", correct_sheets)
    with col2:
        st.metric("Incorrect Row Counts", incorrect_sheets)
    with col3:
        st.metric("Unknown Row Counts", unknown_sheets)

    if incorrect_sheets == 0 and unknown_sheets == 0:
        st.success("All sheets have the correct number of rows!")
    elif incorrect_sheets > 0:
        st.error(f"{incorrect_sheets} sheet(s) have incorrect row counts.")
    elif unknown_sheets > 0:
        st.warning(f"{unknown_sheets} sheet(s) have unknown row counts.")


def validate_data(template_data, data_to_validate):
    """
    Validate data in the Excel file:
    - Column A (Date) and Column B (Time) from specified row - validates format consistency and sequence
    - All other columns (starting from specified row and column) should have only numerical values
    - No null values, special characters (except decimal), alphabetical values, or empty cells allowed

    Parameters:
    -----------
    template_data : dict
        Dictionary containing template file data (not used for validation)
    data_to_validate : dict
        Dictionary containing data file to validate

    Returns:
    --------
    None, displays validation results in Streamlit UI
    """
    st.subheader("Data Validation")

    # Add configuration section
    st.write("### Configure Data Validation")

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Date & Time Configuration**")
        datetime_start_row = st.number_input(
            "Date & Time Starting Row (1-based)",
            min_value=1,
            value=8,
            step=1,
            help="Row number from where Date (Col A) and Time (Col B) validation should start"
        )
        date_col = st.text_input(
            "Date Column",
            value="A",
            help="Column letter for Date (usually A)"
        ).upper().strip()

        time_col = st.text_input(
            "Time Column",
            value="B",
            help="Column letter for Time (usually B)"
        ).upper().strip()

    with col2:
        st.write("**Numeric Data Configuration**")
        numeric_start_row = st.number_input(
            "Numeric Data Starting Row (1-based)",
            min_value=1,
            value=9,
            step=1,
            help="Row number from where numeric data validation should start"
        )

        numeric_start_col = st.text_input(
            "Numeric Data Starting Column",
            value="C",
            help="Column letter from where numeric validation should start (e.g., C)"
        ).upper().strip()

    # Convert column letters to indices
    try:
        date_col_idx = column_letter_to_index(date_col)
        time_col_idx = column_letter_to_index(time_col)
        numeric_start_col_idx = column_letter_to_index(numeric_start_col)
    except:
        st.error("Invalid column letter. Please use A, B, C, etc.")
        return

    validation_results = {}
    all_valid = True
    validation_logs = []

    # Process each sheet in the data file
    for sheet_name in data_to_validate['sheet_names']:
        df = data_to_validate['sheet_data'][sheet_name]
        validation_results[sheet_name] = {
            'datetime_issues': [],
            'datetime_format_issues': [],
            'datetime_sequence_issues': [],
            'numeric_issues': [],
            'null_issues': [],
            'special_char_issues': [],
            'alphabetical_issues': [],
            'empty_cell_issues': [],
            'is_valid': True,
            'total_rows_validated': 0,
            'total_cells_validated': 0,
            'columns_validated': [],
            'detected_date_format': None,
            'detected_time_format': None,
            'detected_interval': None
        }

        # Convert to 0-based index
        datetime_start_row_idx = datetime_start_row - 1
        numeric_start_row_idx = numeric_start_row - 1

        # Validate Date and Time columns
        if len(df) > datetime_start_row_idx and date_col_idx < len(df.columns) and time_col_idx < len(df.columns):
            date_col_name = df.columns[date_col_idx]
            time_col_name = df.columns[time_col_idx]

            validation_results[sheet_name]['columns_validated'].append(
                f"Column {date_col} ({date_col_name}): Date validation & sequence check")
            validation_results[sheet_name]['columns_validated'].append(
                f"Column {time_col} ({time_col_name}): Time validation & sequence check")

            # Extract date and time data
            date_data = df.iloc[datetime_start_row_idx:, date_col_idx]
            time_data = df.iloc[datetime_start_row_idx:, time_col_idx]

            # Detect date format from first valid entry
            detected_date_format = None
            detected_time_format = None
            date_formats_to_try = [
                '%d.%m.%Y',      # 09.01.2026
                '%d/%m/%Y',      # 09/01/2026
                '%d-%m-%Y',      # 09-01-2026
                '%Y.%m.%d',      # 2026.01.09
                '%Y/%m/%d',      # 2026/01/09
                '%Y-%m-%d',      # 2026-01-09
                '%m.%d.%Y',      # 01.09.2026
                '%m/%d/%Y',      # 01/09/2026
                '%m-%d-%Y',      # 01-09-2026
            ]

            time_formats_to_try = [
                '%H:%M:%S',      # 23:00:00
                '%H:%M',         # 23:00
                '%I:%M:%S %p',   # 11:00:00 PM
                '%I:%M %p',      # 11:00 PM
            ]

            # Find first valid date to detect format
            for i in range(len(date_data)):
                date_val = date_data.iloc[i]

                if pd.isna(date_val):
                    continue

                # Handle pandas Timestamp objects
                if isinstance(date_val, (pd.Timestamp, datetime)):
                    detected_date_format = 'datetime_object'
                    break

                date_str = str(date_val).strip()

                # Try to parse with each format
                for fmt in date_formats_to_try:
                    try:
                        datetime.strptime(date_str, fmt)
                        detected_date_format = fmt
                        break
                    except:
                        continue
                if detected_date_format:
                    break

            # Find first valid time to detect format
            for i in range(len(time_data)):
                time_val = time_data.iloc[i]

                if pd.isna(time_val):
                    continue

                # Handle pandas Timestamp objects
                if isinstance(time_val, (pd.Timestamp, datetime)):
                    detected_time_format = 'datetime_object'
                    break

                time_str = str(time_val).strip()

                for fmt in time_formats_to_try:
                    try:
                        datetime.strptime(time_str, fmt)
                        detected_time_format = fmt
                        break
                    except:
                        continue
                if detected_time_format:
                    break

            validation_results[sheet_name]['detected_date_format'] = detected_date_format
            validation_results[sheet_name]['detected_time_format'] = detected_time_format

            if detected_date_format:
                validation_logs.append(
                    f"Sheet '{sheet_name}': Detected date format: {detected_date_format}")
            else:
                validation_logs.append(
                    f"Sheet '{sheet_name}': Could not detect date format - first few values: {[str(date_data.iloc[i]) for i in range(min(3, len(date_data)))]}")

            if detected_time_format:
                validation_logs.append(
                    f"Sheet '{sheet_name}': Detected time format: {detected_time_format}")
            else:
                validation_logs.append(
                    f"Sheet '{sheet_name}': Could not detect time format - first few values: {[str(time_data.iloc[i]) for i in range(min(3, len(time_data)))]}")

            # Validate each date and time entry
            combined_datetimes = []

            for i in range(len(date_data)):
                actual_row_num = datetime_start_row_idx + i + 1
                date_val = date_data.iloc[i]
                time_val = time_data.iloc[i]

                # Check for null/empty dates
                if pd.isna(date_val):
                    validation_results[sheet_name]['datetime_issues'].append(
                        f"Row {actual_row_num}, Column {date_col}: Empty/Null date value"
                    )
                    validation_results[sheet_name]['is_valid'] = False
                    all_valid = False
                    continue

                # Check for null/empty times
                if pd.isna(time_val):
                    validation_results[sheet_name]['datetime_issues'].append(
                        f"Row {actual_row_num}, Column {time_col}: Empty/Null time value"
                    )
                    validation_results[sheet_name]['is_valid'] = False
                    all_valid = False
                    continue

                # Parse date
                parsed_date = None
                if isinstance(date_val, (pd.Timestamp, datetime)):
                    parsed_date = date_val if isinstance(
                        date_val, datetime) else date_val.to_pydatetime()
                else:
                    date_str = str(date_val).strip()
                    if detected_date_format and detected_date_format != 'datetime_object':
                        try:
                            parsed_date = datetime.strptime(
                                date_str, detected_date_format)
                        except Exception as e:
                            validation_results[sheet_name]['datetime_format_issues'].append(
                                f"Row {actual_row_num}, Column {date_col}: Date '{date_str}' doesn't match format {detected_date_format}"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue
                    elif not detected_date_format:
                        # Try all formats as fallback
                        date_parsed = False
                        for fmt in date_formats_to_try:
                            try:
                                parsed_date = datetime.strptime(date_str, fmt)
                                date_parsed = True
                                break
                            except:
                                continue

                        if not date_parsed:
                            validation_results[sheet_name]['datetime_format_issues'].append(
                                f"Row {actual_row_num}, Column {date_col}: Could not parse date '{date_str}' with any known format"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue

                # Parse time
                parsed_time = None
                if isinstance(time_val, (pd.Timestamp, datetime)):
                    parsed_time = time_val if isinstance(
                        time_val, datetime) else time_val.to_pydatetime()
                else:
                    time_str = str(time_val).strip()
                    if detected_time_format and detected_time_format != 'datetime_object':
                        try:
                            parsed_time = datetime.strptime(
                                time_str, detected_time_format)
                        except Exception as e:
                            validation_results[sheet_name]['datetime_format_issues'].append(
                                f"Row {actual_row_num}, Column {time_col}: Time '{time_str}' doesn't match format {detected_time_format}"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue
                    elif not detected_time_format:
                        # Try all formats as fallback
                        time_parsed = False
                        for fmt in time_formats_to_try:
                            try:
                                parsed_time = datetime.strptime(time_str, fmt)
                                time_parsed = True
                                break
                            except:
                                continue

                        if not time_parsed:
                            validation_results[sheet_name]['datetime_format_issues'].append(
                                f"Row {actual_row_num}, Column {time_col}: Could not parse time '{time_str}' with any known format"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue

                # Combine date and time
                if parsed_date and parsed_time:
                    try:
                        combined_dt = datetime.combine(
                            parsed_date.date(), parsed_time.time())
                        combined_datetimes.append({
                            'row': actual_row_num,
                            'datetime': combined_dt
                        })
                    except Exception as e:
                        validation_results[sheet_name]['datetime_issues'].append(
                            f"Row {actual_row_num}: Could not combine date and time. Error: {str(e)}"
                        )
                        validation_results[sheet_name]['is_valid'] = False
                        all_valid = False

            # Check datetime sequence and detect interval
            if len(combined_datetimes) >= 2:
                # Calculate intervals between consecutive datetime entries
                intervals = []
                for i in range(len(combined_datetimes) - 1):
                    dt1 = combined_datetimes[i]['datetime']
                    dt2 = combined_datetimes[i + 1]['datetime']
                    interval = (dt2 - dt1).total_seconds()
                    intervals.append(interval)

                # Detect the most common interval (in seconds)
                from collections import Counter
                interval_counts = Counter(intervals)
                most_common_interval = interval_counts.most_common(1)[0][0]

                # Determine interval type
                if most_common_interval == 60:
                    interval_type = "1 minute"
                elif most_common_interval == 3600:
                    interval_type = "1 hour"
                elif most_common_interval == 86400:
                    interval_type = "1 day"
                else:
                    interval_minutes = most_common_interval / 60
                    interval_hours = most_common_interval / 3600
                    if interval_minutes < 60:
                        interval_type = f"{int(interval_minutes)} minutes"
                    else:
                        interval_type = f"{interval_hours:.2f} hours"

                validation_results[sheet_name]['detected_interval'] = interval_type
                validation_logs.append(
                    f"Sheet '{sheet_name}': Detected time interval: {interval_type}")

                # Validate sequence with detected interval
                for i in range(len(combined_datetimes) - 1):
                    dt1 = combined_datetimes[i]['datetime']
                    dt2 = combined_datetimes[i + 1]['datetime']
                    row1 = combined_datetimes[i]['row']
                    row2 = combined_datetimes[i + 1]['row']

                    actual_interval = (dt2 - dt1).total_seconds()

                    # Allow small tolerance (1 second)
                    if abs(actual_interval - most_common_interval) > 1:
                        expected_dt = dt1 + \
                            timedelta(seconds=most_common_interval)
                        validation_results[sheet_name]['datetime_sequence_issues'].append(
                            f"Rows {row1}-{row2}: DateTime sequence broken. Expected {expected_dt.strftime('%d.%m.%Y %H:%M:%S')}, got {dt2.strftime('%d.%m.%Y %H:%M:%S')}"
                        )
                        validation_results[sheet_name]['is_valid'] = False
                        all_valid = False

            validation_results[sheet_name]['total_rows_validated'] = len(
                date_data)

        # Validate numeric columns starting from specified column
        if len(df) > numeric_start_row_idx:
            # Find last column with data
            last_col_with_data = -1
            for col_idx in range(len(df.columns)):
                if df.iloc[:, col_idx].notna().any():
                    last_col_with_data = col_idx

            if numeric_start_col_idx < len(df.columns):
                data_rows = df.iloc[numeric_start_row_idx:]

                # Validate from numeric_start_col_idx to last_col_with_data
                for col_idx in range(numeric_start_col_idx, last_col_with_data + 1):
                    col_name = df.columns[col_idx]
                    col_letter = get_excel_column_name(col_idx)

                    validation_results[sheet_name]['columns_validated'].append(
                        f"Column {col_letter} ({col_name}): Numeric validation")
                    validation_logs.append(
                        f"Sheet '{sheet_name}': Validating Column {col_letter} ({col_name}) for numeric values")

                    col_data = data_rows[col_name]

                    for idx, value in col_data.items():
                        row_num = idx + 1
                        validation_results[sheet_name]['total_cells_validated'] += 1
                        cell_ref = f"{col_letter}{row_num}"

                        # Check for null/empty values
                        if pd.isna(value):
                            validation_results[sheet_name]['null_issues'].append(
                                f"Cell {cell_ref} (Row {row_num}, Column {col_name}): Null/Empty value found"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue

                        # Convert to string for validation
                        value_str = str(value).strip()

                        # Check for empty string
                        if value_str == "":
                            validation_results[sheet_name]['empty_cell_issues'].append(
                                f"Cell {cell_ref} (Row {row_num}, Column {col_name}): Empty cell found"
                            )
                            validation_results[sheet_name]['is_valid'] = False
                            all_valid = False
                            continue

                        # Check if the value is numeric
                        if not pd.api.types.is_numeric_dtype(type(value)) and not isinstance(value, (int, float, np.number)):
                            try:
                                # Try to convert to float
                                float_value = float(value_str)
                                validation_logs.append(
                                    f"Sheet '{sheet_name}', Cell {cell_ref}: Value '{value_str}' successfully converted to numeric")
                            except ValueError:
                                # Check if it contains alphabetical characters
                                if any(c.isalpha() for c in value_str):
                                    validation_results[sheet_name]['alphabetical_issues'].append(
                                        f"Cell {cell_ref} (Row {row_num}, Column {col_name}): Contains alphabetical characters '{value_str}'"
                                    )
                                    validation_results[sheet_name]['is_valid'] = False
                                    all_valid = False
                                else:
                                    # Check for special characters (excluding decimal point and negative sign)
                                    allowed_chars = set('0123456789.-')
                                    if not set(value_str).issubset(allowed_chars):
                                        validation_results[sheet_name]['special_char_issues'].append(
                                            f"Cell {cell_ref} (Row {row_num}, Column {col_name}): Contains special characters '{value_str}'"
                                        )
                                        validation_results[sheet_name]['is_valid'] = False
                                        all_valid = False
                                    else:
                                        validation_results[sheet_name]['numeric_issues'].append(
                                            f"Cell {cell_ref} (Row {row_num}, Column {col_name}): Value '{value_str}' is not numeric"
                                        )
                                        validation_results[sheet_name]['is_valid'] = False
                                        all_valid = False

                                validation_logs.append(
                                    f"Sheet '{sheet_name}', Cell {cell_ref}: Value '{value_str}' failed numeric validation")

    # Display validation summary with sheet selector
    st.subheader("Validation Summary")

    # Show overall status
    if all_valid:
        st.success("✓ All data validation checks passed!")
    else:
        st.error("✗ Data validation found issues. Please check the details below.")

    # Sheet selector for summary
    st.write("### Select Sheet to View Summary")
    sheet_options = ["All Sheets"] + list(validation_results.keys())
    selected_sheet = st.selectbox("Choose a sheet:", sheet_options)

    # Display statistics
    if selected_sheet == "All Sheets":
        # Show overall statistics
        total_sheets = len(validation_results)
        valid_sheets = sum(
            1 for r in validation_results.values() if r['is_valid'])
        invalid_sheets = total_sheets - valid_sheets

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Sheets", total_sheets)
        with col2:
            st.metric("✓ Valid Sheets", valid_sheets)
        with col3:
            st.metric("✗ Invalid Sheets", invalid_sheets)

        # Show summary for all sheets
        for sheet_name, results in validation_results.items():
            with st.expander(f"Sheet: {sheet_name} - {'✓ Valid' if results['is_valid'] else '✗ Invalid'}"):
                display_sheet_summary(sheet_name, results)
    else:
        # Show summary for selected sheet only
        if selected_sheet in validation_results:
            results = validation_results[selected_sheet]

            # Display metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Rows Validated", results['total_rows_validated'])
            with col2:
                st.metric("Cells Validated", results['total_cells_validated'])
            with col3:
                st.metric("Columns Checked", len(results['columns_validated']))
            with col4:
                status = "✓ Valid" if results['is_valid'] else "✗ Invalid"
                st.metric("Status", status)

            # Show detected formats and intervals
            if results['detected_date_format'] or results['detected_time_format'] or results['detected_interval']:
                st.write("---")
                st.write("**Detected Patterns:**")
                col1, col2, col3 = st.columns(3)
                with col1:
                    if results['detected_date_format']:
                        st.info(
                            f"📅 Date Format: `{results['detected_date_format']}`")
                with col2:
                    if results['detected_time_format']:
                        st.info(
                            f"🕐 Time Format: `{results['detected_time_format']}`")
                with col3:
                    if results['detected_interval']:
                        st.info(
                            f"⏱️ Interval: `{results['detected_interval']}`")

            display_sheet_summary(selected_sheet, results)

    # Create expandable section for detailed logs
    with st.expander("View Detailed Validation Logs"):
        st.write("**Validation Process Logs:**")
        for log in validation_logs:
            st.write(f"- {log}")

        st.write("**Columns Validated per Sheet:**")
        for sheet_name, results in validation_results.items():
            with st.container():
                st.write(f"**Sheet: {sheet_name}**")
                st.write(
                    f"- Rows validated: {results['total_rows_validated']}")
                st.write(
                    f"- Cells validated: {results['total_cells_validated']}")
                if results['detected_date_format']:
                    st.write(
                        f"- Detected date format: {results['detected_date_format']}")
                if results['detected_time_format']:
                    st.write(
                        f"- Detected time format: {results['detected_time_format']}")
                if results['detected_interval']:
                    st.write(
                        f"- Detected time interval: {results['detected_interval']}")
                for col_info in results['columns_validated']:
                    st.write(f"- {col_info}")


def display_sheet_summary(sheet_name, results):
    """Helper function to display validation summary for a sheet"""
    if results['is_valid']:
        st.success(
            f"✓ All data validation checks passed for sheet '{sheet_name}'")
    else:
        st.error(f"✗ Found validation issues in sheet '{sheet_name}'")

        # Count total issues
        total_issues = (
            len(results['datetime_issues']) +
            len(results['datetime_format_issues']) +
            len(results['datetime_sequence_issues']) +
            len(results['numeric_issues']) +
            len(results['null_issues']) +
            len(results['special_char_issues']) +
            len(results['alphabetical_issues']) +
            len(results['empty_cell_issues'])
        )

        st.write(f"**Total Issues Found: {total_issues}**")

        # Display issue breakdown
        col1, col2, col3 = st.columns(3)
        with col1:
            if results['datetime_issues']:
                st.metric("DateTime Issues", len(results['datetime_issues']))
            if results['datetime_format_issues']:
                st.metric("Format Issues", len(
                    results['datetime_format_issues']))
            if results['datetime_sequence_issues']:
                st.metric("Sequence Issues", len(
                    results['datetime_sequence_issues']))

        with col2:
            if results['null_issues']:
                st.metric("Null Values", len(results['null_issues']))
            if results['empty_cell_issues']:
                st.metric("Empty Cells", len(results['empty_cell_issues']))
            if results['numeric_issues']:
                st.metric("Numeric Issues", len(results['numeric_issues']))

        with col3:
            if results['special_char_issues']:
                st.metric("Special Characters", len(
                    results['special_char_issues']))
            if results['alphabetical_issues']:
                st.metric("Alphabetical Values", len(
                    results['alphabetical_issues']))

        # Display detailed issues
        if results['datetime_issues']:
            with st.expander("📅 DateTime Validation Issues"):
                for issue in results['datetime_issues']:
                    st.write(f"- {issue}")

        if results['datetime_format_issues']:
            with st.expander("📝 DateTime Format Issues"):
                for issue in results['datetime_format_issues']:
                    st.write(f"- {issue}")

        if results['datetime_sequence_issues']:
            with st.expander("🔄 DateTime Sequence Issues"):
                for issue in results['datetime_sequence_issues']:
                    st.write(f"- {issue}")

        if results['null_issues']:
            with st.expander("❌ Null Value Issues"):
                for issue in results['null_issues']:
                    st.write(f"- {issue}")

        if results['empty_cell_issues']:
            with st.expander("📭 Empty Cell Issues"):
                for issue in results['empty_cell_issues']:
                    st.write(f"- {issue}")

        if results['numeric_issues']:
            with st.expander("🔢 Numeric Validation Issues"):
                for issue in results['numeric_issues']:
                    st.write(f"- {issue}")

        if results['alphabetical_issues']:
            with st.expander("🔤 Alphabetical Character Issues"):
                for issue in results['alphabetical_issues']:
                    st.write(f"- {issue}")

        if results['special_char_issues']:
            with st.expander("⚠️ Special Character Issues"):
                for issue in results['special_char_issues']:
                    st.write(f"- {issue}")


def column_letter_to_index(col_letter):
    """Convert Excel column letter to 0-based index (A=0, B=1, C=2, ..., Z=25, AA=26, etc.)"""
    col_letter = col_letter.upper()
    result = 0
    for i, char in enumerate(reversed(col_letter)):
        result += (ord(char) - ord('A') + 1) * (26 ** i)
    return result - 1


def get_excel_column_name(col_idx):
    """Convert 0-based column index to Excel column name (A, B, C, ..., Z, AA, AB, ...)"""
    result = ""
    while True:
        col_idx, remainder = divmod(col_idx, 26)
        result = chr(65 + remainder) + result
        if col_idx == 0:
            break
        col_idx -= 1
    return result


def find_missing_data(data_to_validate):
    """
    Find missing datetime data in the Excel sheets.

    This function analyzes each sheet in the data file and identifies minutes missing from
    the datetime sequence in Column A, starting from row 7, assuming each row should
    increment by 1 minute.

    Args:
        data_to_validate (dict): Dictionary containing sheet data from the uploaded file

    Returns:
        None: Results are displayed directly in the Streamlit UI
    """
    import streamlit as st
    import pandas as pd
    from datetime import datetime, timedelta
    import numpy as np

    st.subheader("Find Missing Data")

    # Get sheet names and data
    sheet_names = data_to_validate['sheet_names']
    sheet_data = data_to_validate['sheet_data']

    if not sheet_names:
        st.warning("No sheets found in the data file.")
        return

    # Create sheet selector
    selected_sheet = st.selectbox("Select a sheet to analyze:", sheet_names)

    if selected_sheet:
        df = sheet_data[selected_sheet]

        # Check if dataframe is not empty and has enough rows
        if df.empty:
            st.warning(f"Sheet '{selected_sheet}' is empty.")
            return

        if len(df) < 7:
            st.warning(
                f"Sheet '{selected_sheet}' has fewer than 7 rows. Analysis requires at least 7 rows.")
            return

        # Extract datetime column (Column A) starting from row 7
        # In pandas, index is 0-based, so row 7 is at index 6
        datetime_col = df.iloc[6:, 0].copy()

        # Check if the column looks like datetime data
        try:
            # Try to convert to datetime format if not already
            if not pd.api.types.is_datetime64_any_dtype(datetime_col):
                datetime_col = pd.to_datetime(datetime_col, errors='coerce')

            # Remove NaT values (could be non-datetime entries)
            datetime_col = datetime_col.dropna()

            if datetime_col.empty:
                st.error(
                    f"No valid datetime data found in Column A of sheet '{selected_sheet}'.")
                return

            # Sort the datetime values (preserve the original index for reference)
            datetime_col = datetime_col.sort_values()

            # Automatic time interval detection
            time_diffs = datetime_col.diff().dropna()

            # Find the most common time difference (in minutes)
            if len(time_diffs) > 0:
                # Convert to minutes for analysis
                diff_minutes = time_diffs.dt.total_seconds() / 60
                expected_interval = diff_minutes.value_counts().idxmax()

                # Round to nearest integer if close
                if abs(expected_interval - round(expected_interval)) < 0.1:
                    expected_interval = round(expected_interval)

                # Display the detected interval
                if expected_interval == 1:
                    interval_str = "1 minute"
                elif expected_interval < 1:
                    interval_str = f"{expected_interval*60:.0f} seconds"
                else:
                    interval_str = f"{expected_interval:.0f} minutes"

                st.info(f"Detected time interval between rows: {interval_str}")

                # Option to override detected interval
                use_custom_interval = st.checkbox(
                    "Use custom interval instead")

                if use_custom_interval:
                    interval_unit = st.selectbox(
                        "Interval unit:", ["Minutes", "Seconds", "Hours"])

                    if interval_unit == "Minutes":
                        custom_interval = st.number_input("Minutes between each row:",
                                                          min_value=1, value=1, step=1)
                        expected_interval = custom_interval
                    elif interval_unit == "Seconds":
                        custom_interval = st.number_input("Seconds between each row:",
                                                          min_value=1, value=60, step=1)
                        expected_interval = custom_interval / 60
                    else:  # Hours
                        custom_interval = st.number_input("Hours between each row:",
                                                          min_value=1, value=1, step=1)
                        expected_interval = custom_interval * 60

                    st.info(
                        f"Using custom interval: {expected_interval:.1f} minutes")

                # Find missing timestamps
                missing_timestamps = []
                gap_ranges = []

                for i in range(len(datetime_col) - 1):
                    current_dt = datetime_col.iloc[i]
                    next_dt = datetime_col.iloc[i+1]

                    # Calculate difference in minutes
                    diff = (next_dt - current_dt).total_seconds() / 60

                    # Check if there's a gap (more than the expected interval)
                    if diff > expected_interval * 1.1:  # 10% tolerance
                        # Calculate how many rows should be in between
                        expected_rows = int(diff / expected_interval) - 1

                        # Generate the missing timestamps
                        gap_start = current_dt
                        gap_end = next_dt

                        # Store the gap range
                        gap_ranges.append({
                            "Gap Start": gap_start,
                            "Gap End": gap_end,
                            "Missing Minutes": expected_rows * expected_interval,
                            "Missing Rows": expected_rows
                        })

                        # Generate individual missing timestamps
                        for j in range(1, expected_rows + 1):
                            missing_dt = current_dt + \
                                timedelta(minutes=j * expected_interval)
                            missing_timestamps.append(missing_dt)

                if not gap_ranges:
                    st.success(
                        "? No missing data found! All data points follow the expected time sequence.")
                else:
                    # Summarize the findings
                    total_missing = len(missing_timestamps)
                    st.error(
                        f"? Found {len(gap_ranges)} gaps with {total_missing} missing data points.")

                    # Display gap summary
                    gap_df = pd.DataFrame(gap_ranges)
                    st.subheader("Gap Summary")
                    st.dataframe(gap_df, use_container_width=True)

                    # Option to view detailed missing timestamps
                    if st.checkbox("Show all missing timestamps"):
                        timestamps_df = pd.DataFrame(
                            {"Missing Timestamp": missing_timestamps})
                        timestamps_df = timestamps_df.sort_values(
                            "Missing Timestamp")
                        st.dataframe(timestamps_df, use_container_width=True)

                    # Visual representation of gaps
                    st.subheader("Gap Visualization")

                    import plotly.graph_objects as go

                    # Create a continuous range from min to max datetime
                    if expected_interval >= 1:
                        # Use minutes for larger intervals
                        freq = f"{int(expected_interval)}min"
                    else:
                        # Use seconds for smaller intervals
                        freq = f"{int(expected_interval*60)}s"

                    full_range = pd.date_range(
                        start=datetime_col.min(),
                        end=datetime_col.max(),
                        freq=freq
                    )

                    # Ensure visualization is manageable
                    max_points = 10000
                    if len(full_range) > max_points:
                        # Sample points for visualization
                        step = len(full_range) // max_points + 1
                        full_range = full_range[::step]

                    # Create data availability series (1 for present, 0 for missing)
                    availability = pd.Series(0, index=full_range)

                    # Mark existing data points
                    for dt in datetime_col:
                        # Find closest point in our reference range
                        idx = full_range.get_indexer([dt], method='nearest')[0]
                        if 0 <= idx < len(full_range):
                            availability.iloc[idx] = 1

                    # Create figure
                    fig = go.Figure()

                    # Add data availability trace
                    fig.add_trace(go.Scatter(
                        x=availability.index,
                        y=availability.values,
                        mode='lines',
                        name='Data Present',
                        line=dict(color='green', width=2),
                        fill='tozeroy'
                    ))

                    # Add gap annotations
                    for idx, gap in enumerate(gap_ranges):
                        fig.add_vrect(
                            x0=gap['Gap Start'],
                            x1=gap['Gap End'],
                            fillcolor="red",
                            opacity=0.2,
                            layer="below",
                            line_width=0,
                            annotation_text=f"Gap {idx+1}",
                            annotation_position="top left"
                        )

                    fig.update_layout(
                        title=f"Data Availability for Sheet '{selected_sheet}'",
                        xaxis_title="Date/Time",
                        yaxis_title="Data Present (1) / Missing (0)",
                        height=500,
                        yaxis=dict(tickvals=[0, 1]),
                        hovermode="x"
                    )

                    st.plotly_chart(fig, use_container_width=True)

                    # Export options
                    st.subheader("Export Results")

                    # Gaps summary export
                    gaps_csv = gap_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        "Download Gap Summary CSV",
                        gaps_csv,
                        f"gap_summary_{selected_sheet}.csv",
                        "text/csv",
                        key='download-gaps'
                    )

                    # Missing timestamps export
                    if missing_timestamps:
                        timestamps_df = pd.DataFrame(
                            {"Missing Timestamp": missing_timestamps})
                        timestamps_csv = timestamps_df.to_csv(
                            index=False).encode('utf-8')
                        st.download_button(
                            "Download All Missing Timestamps CSV",
                            timestamps_csv,
                            f"missing_timestamps_{selected_sheet}.csv",
                            "text/csv",
                            key='download-timestamps'
                        )
            else:
                st.warning("Not enough data points to analyze time intervals.")

        except Exception as e:
            st.error(f"Error analyzing datetime data: {e}")
            st.write("Please ensure Column A contains valid datetime values.")

            # Show a sample of data for debugging
            st.subheader("Sample Data (First 10 rows)")
            st.dataframe(df.head(10))
    else:
        st.info("Please select a sheet to analyze.")


def export_data_to_csv_tab(template_data, data_to_validate):
    """
    Tab function for exporting sensor data to individual CSV files
    """
    st.subheader("Export Data to CSV")

    # User inputs
    col1, col2 = st.columns(2)

    with col1:
        # Sheet selection
        available_sheets = list(data_to_validate['sheet_data'].keys())
        selected_sheet = st.selectbox(
            "Select Sheet to Export", available_sheets)

        # Header row input
        header_row = st.number_input(
            "Header Row (1-based, 0 for no header)", min_value=0, value=1, step=1)

        # Starting row input
        start_row = st.number_input(
            "Data Starting Row (1-based)", min_value=1, value=8, step=1)

    with col2:
        # Asset and hierarchy row references
        asset_row = st.number_input(
            "Asset Row in Template (1-based)", min_value=1, value=7, step=1)
        hierarchy_row = st.number_input(
            "Hierarchy Row in Template (1-based)", min_value=1, value=8, step=1)

    if st.button("Generate CSV Files"):
        try:
            # Get the selected sheet data
            if selected_sheet not in data_to_validate['sheet_data']:
                st.error(f"Sheet '{selected_sheet}' not found in data file")
                return

            if selected_sheet not in template_data['sheet_data']:
                st.error(
                    f"Sheet '{selected_sheet}' not found in template file")
                return

            data_df = data_to_validate['sheet_data'][selected_sheet].copy()
            template_df = template_data['sheet_data'][selected_sheet].copy()

            # Handle case where column names might not be proper strings
            if header_row == 0:
                # Create generic column names
                data_df.columns = [
                    f'Timestamp' if i == 0 else f'Sensor_{i}' for i in range(len(data_df.columns))]
                template_df.columns = data_df.columns  # Use same column names
            else:
                # Ensure all column names are strings
                data_df.columns = [str(col) for col in data_df.columns]
                template_df.columns = [str(col) for col in template_df.columns]

            # Extract data starting from the specified row (convert to 0-based index)
            data_subset = data_df.iloc[start_row-1:].copy()

            if data_subset.empty:
                st.error("No data found starting from the specified row")
                return

            # Reset index for clean processing
            data_subset = data_subset.reset_index(drop=True)

            # Get column names (assuming first column is timestamp)
            columns = data_subset.columns.tolist()

            if len(columns) < 2:
                st.error("Need at least timestamp column and one sensor column")
                return

            timestamp_col = columns[0]
            sensor_cols = columns[1:]  # All columns except timestamp

            st.info(f"Found {len(sensor_cols)} sensor columns to process")

            # Process each sensor column
            csv_files_created = []

            for i, sensor_col in enumerate(sensor_cols):
                try:
                    # Create individual CSV for this sensor
                    csv_content = create_sensor_csv(
                        data_subset,
                        timestamp_col,
                        sensor_col,
                        template_df,
                        # Column index in template (1-based, B=1, C=2, etc.)
                        i + 1,
                        asset_row - 1,  # Convert to 0-based
                        hierarchy_row - 1  # Convert to 0-based
                    )

                    if csv_content:
                        # Create download button for this CSV
                        # Handle non-string column names
                        sensor_name = str(sensor_col).replace(
                            ' ', '_').replace('/', '_').replace('.', '_')
                        csv_filename = f"sensor_{i+1}_{sensor_name}_data.csv"

                        st.download_button(
                            label=f"Download Sensor {i+1} ({str(sensor_col)[:20]}...) CSV" if len(
                                str(sensor_col)) > 20 else f"Download Sensor {i+1} ({sensor_col}) CSV",
                            data=csv_content,
                            file_name=csv_filename,
                            mime="text/csv",
                            key=f"download_{i}"
                        )

                        csv_files_created.append(csv_filename)

                except Exception as e:
                    st.error(
                        f"Error processing sensor '{sensor_col}': {str(e)}")
                    continue

            if csv_files_created:
                st.success(
                    f"Successfully created {len(csv_files_created)} CSV files")
                st.info(
                    "Click the download buttons above to save individual sensor CSV files")
            else:
                st.error("No CSV files were created")

        except Exception as e:
            st.error(f"Error during CSV export: {str(e)}")


def create_sensor_csv(data_df, timestamp_col, sensor_col, template_df, col_index, asset_row, hierarchy_row):
    """
    Create CSV content for a single sensor with timestamp, asset, hierarchy, sensor value, and hierarchy details
    """
    try:
        # Get asset and hierarchy values from template
        # Column index: A=0, B=1, C=2, etc.
        try:
            asset_value = str(template_df.iloc[asset_row, col_index]).strip()
            hierarchy_value = str(
                template_df.iloc[hierarchy_row, col_index]).strip()
        except (IndexError, KeyError):
            st.warning(
                f"Could not find asset/hierarchy values for column {col_index+1} in template")
            asset_value = "Unknown"
            hierarchy_value = "Unknown"

        # Handle NaN values
        if asset_value.lower() in ['nan', 'none', '']:
            asset_value = "Unknown"
        if hierarchy_value.lower() in ['nan', 'none', '']:
            hierarchy_value = "Unknown"

        # Split asset and hierarchy values
        asset_parts = [part.strip() for part in asset_value.split(',')]
        hierarchy_parts = [part.strip() for part in hierarchy_value.split(',')]

        # Create CSV content
        csv_rows = []

        # Process each row of data
        for _, row in data_df.iterrows():
            try:
                # Format timestamp
                timestamp = format_timestamp(row[timestamp_col])
                sensor_value = row[sensor_col]

                # Handle NaN sensor values
                if pd.isna(sensor_value):
                    sensor_value = ""

                # Create row: timestamp, asset_parts, sensor_value, hierarchy_parts
                csv_row = [timestamp] + asset_parts + \
                    [str(sensor_value)] + hierarchy_parts
                csv_rows.append(csv_row)

            except Exception as e:
                st.warning(f"Skipping row due to error: {str(e)}")
                continue

        if not csv_rows:
            return None

        # Convert to CSV string
        import io
        output = io.StringIO()

        # Write header (optional, you can modify this based on requirements)
        header = ['Timestamp'] + [f'Asset_{i+1}' for i in range(len(asset_parts))] + [
            'Value'] + [f'Hierarchy_{i+1}' for i in range(len(hierarchy_parts))]
        output.write(','.join(header) + '\n')

        # Write data rows
        for row in csv_rows:
            # Escape commas in values by quoting them
            escaped_row = []
            for value in row:
                value_str = str(value)
                if ',' in value_str or '"' in value_str:
                    # Escape quotes and wrap in quotes
                    value_str = '"' + value_str.replace('"', '""') + '"'
                escaped_row.append(value_str)

            output.write(','.join(escaped_row) + '\n')

        return output.getvalue()

    except Exception as e:
        st.error(f"Error creating CSV for sensor {sensor_col}: {str(e)}")
        return None


def format_timestamp(timestamp_value):
    """
    Convert timestamp to yyyy/mm/dd hh:mm:ss format (24-hour format)
    """
    try:
        # Handle different timestamp formats
        if pd.isna(timestamp_value):
            return ""

        # If it's already a datetime object
        if isinstance(timestamp_value, (pd.Timestamp, datetime)):
            return timestamp_value.strftime('%Y/%m/%d %H:%M:%S')

        # If it's a string, try to parse it
        if isinstance(timestamp_value, str):
            # Clean the string first
            timestamp_str = str(timestamp_value).strip()

            # Try common formats including AM/PM formats
            formats_to_try = [
                '%m/%d/%Y %I:%M:%S %p',   # 7/1/2024 12:00:00 AM/PM
                # 7/1/2024  12:00:00 AM/PM (double space)
                '%m/%d/%Y  %I:%M:%S %p',
                '%m/%d/%Y %H:%M:%S',      # 7/1/2024 00:00:00 (24-hour)
                '%Y-%m-%d %H:%M:%S',      # 2024-07-01 00:00:00
                '%m/%d/%Y %I:%M %p',      # 7/1/2024 12:00 AM/PM
                # 7/1/2024  12:00 AM/PM (double space)
                '%m/%d/%Y  %I:%M %p',
                '%m/%d/%Y %H:%M',         # 7/1/2024 00:00
                '%Y-%m-%d %H:%M',         # 2024-07-01 00:00
                '%m/%d/%Y',               # 7/1/2024
                '%Y-%m-%d',               # 2024-07-01
            ]

            for fmt in formats_to_try:
                try:
                    dt = datetime.strptime(timestamp_str, fmt)
                    return dt.strftime('%Y/%m/%d %H:%M:%S')
                except ValueError:
                    continue

        # If it's a number (Excel date serial)
        if isinstance(timestamp_value, (int, float)):
            # Excel date serial number (days since 1900-01-01)
            excel_epoch = datetime(1900, 1, 1)
            # -2 for Excel's leap year bug
            dt = excel_epoch + timedelta(days=timestamp_value - 2)
            return dt.strftime('%Y/%m/%d %H:%M:%S')

        # Last resort: try pandas to_datetime
        try:
            dt = pd.to_datetime(timestamp_value)
            return dt.strftime('%Y/%m/%d %H:%M:%S')
        except:
            pass

        # Final fallback: convert to string
        return str(timestamp_value)

    except Exception as e:
        return str(timestamp_value)  # Return as-is if conversion fails


if __name__ == "__main__":
    main()
