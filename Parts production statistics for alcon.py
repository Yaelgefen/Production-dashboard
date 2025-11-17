import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import plotly.express as px
import plotly.figure_factory as ff
import datetime
from datetime import datetime
import os
import warnings
from dateutil import parser
import numpy as np

warnings.filterwarnings('ignore')


def extract_month_from_filename(filename):
    """
    Extract month from filename format PM2507 (July 2025)
    PM = Parts Manufacturing (or similar)
    25 = year (2025)
    07 = month (July)
    Returns a datetime object for the first day of that month
    """
    import re

    # Look for pattern like PM2507, PM2508, etc.
    match = re.search(r'PM(\d{2})(\d{2})', filename)

    if match:
        year_short = match.group(1)  # "25"
        month = match.group(2)  # "07"

        # Convert to full year
        year = 2000 + int(year_short)  # 25 -> 2025
        month_int = int(month)  # 07 -> 7

        # Return first day of that month
        return pd.to_datetime(f"{year}-{month_int:02d}-01")

    # If no match found, return None
    return None

def standardize_rejection_reasons(df):
    """
    Standardize rejection reasons based on approved list.
    Any reason not in the list becomes 'N/A'.
    Normalize surface irregularities variations.
    """
    if df.empty or 'Reason' not in df.columns:
        return df

    df = df.copy()

    # Define the approved rejection reasons list
    approved_reasons = {
        # Standard reasons (keep as-is, case-insensitive)
        'bubbles and voids': 'Bubbles and voids',
        'embedded particle': 'Embedded particle',
        'tears': 'Tears',
        'not cured': 'Not cured',
        'other': 'Other',
        'embedded metals': 'Embedded metals',
        'hole in sm': 'Hole in SM',
        'lack of uniformity in pigment': 'Lack of uniformity in pigment',
        'surface irregularity': 'Surface irregularity',

        # Keep N/A as-is for already processed data
        'n/a': 'N/A',
        'na': 'N/A',
        'NA': 'N/A'
    }

    # Function to standardize individual reason
    def standardize_reason(reason):
        if pd.isna(reason) or str(reason).strip() == '':
            return 'N/A'

        # Convert to string and clean
        reason_str = str(reason).strip().lower()

        # Check if it's in approved list
        if reason_str in approved_reasons:
            return approved_reasons[reason_str]

        # If not found, return N/A
        return 'N/A'

    # Apply standardization only to rejected parts
    rejected_mask = df['Status'] == 'rejected'
    df.loc[rejected_mask, 'Reason'] = df.loc[rejected_mask, 'Reason'].apply(standardize_reason)

    # For accepted parts, ensure reason is empty
    df.loc[df['Status'] == 'accepted', 'Reason'] = ''

    return df


def get_rejection_reason_colors():
    """
    Define consistent colors for each rejection reason.
    Each reason will always have the same color across all charts.
    """
    color_mapping = {
        # Updated merged reasons
        'Tears/ Voids': '#b3e5fc',  # Very light blue (for FTA)
        'Embedded particle': '#ffcc99',  # Light orange (merged for all)
        'Tears': '#99ff99',  # Light green (for Optic/PL/Wings)
        'Bubbles and voids': '#b3e5fc',  # Very light blue (for Optic/PL)
        'Other': '#ff99cc',  # Light pink (merged for all)
        'Hole in SM': '#d3d3d3',  # Light gray
        'Lack of uniformity in pigment': '#66b3ff',  # Light blue
        'Surface irregularity': '#fffacd',  # Light yellow
        'Not cured': '#c2c2f0',  # Light purple
        'N/A': '#17becf',  # Cyan
        'Accepted': '#28a745',  # Green for accepted parts
    }
    return color_mapping

def create_color_sequence_for_reasons(reasons_list):
    """
    Create a color sequence based on the reasons present in the data.
    This ensures consistent colors across all charts.
    """
    color_mapping = get_rejection_reason_colors()

    # Create color sequence in the order of reasons_list
    colors = []
    for reason in reasons_list:
        if reason in color_mapping:
            colors.append(color_mapping[reason])
        else:
            # Fallback color for any unmapped reason
            colors.append('#95a5a6')  # Light gray

    return colors

# Page configuration
st.set_page_config(page_title="Parts Production Statistics", page_icon=":bar_chart:", layout="wide")
st.title(" :bar_chart: Parts Production Statistics")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)

# Add custom CSS for Segoe UI Semilight font
st.markdown("""
<style>
    .stApp {
        font-family: 'Segoe UI Semilight (Body)', 'Segoe UI', sans-serif;
    }
    * {
        font-family: 'Segoe UI Semilight (Body)', 'Segoe UI', sans-serif;
    }
</style>
""", unsafe_allow_html=True)

# Multiple file upload
st.subheader("📁 Upload Excel Files")
uploaded_files = st.file_uploader(
    ":file_folder: Upload Excel files",
    type=(["xlsx", "xls"]),
    accept_multiple_files=True,
    help="You can upload multiple Excel files. Each should contain QC sheets."
)


def find_data_start_row(df, default_start=8):
    """Find the row where actual data starts by looking for accept/reject values"""
    for row_idx in range(min(15, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.lower()
        if any('accept' in val or 'reject' in val for val in row_values):
            return max(0, row_idx - 1)
    return default_start


def process_optic_qc(df):
    """Process Optic QC sheet with specific structure"""
    try:
        st.write("🔍 **DEBUG: Starting Optic QC Processing**")
        st.write(f"Raw dataframe shape: {df.shape}")

        # Show first few rows of raw data
        with st.expander("🔍 Debug: First 10 rows of raw Optic QC data"):
            st.write(df.head(10))

        # data starts at index 8
        data = df.iloc[8:].copy()
        st.write(f"After selecting from row 8: {data.shape[0]} rows")

        # Check what's in the status column (index 6)
        st.write(f"🔍 Checking column 6 (status column):")
        status_col_values = data.iloc[:, 6].dropna().unique()
        st.write(f"Unique values in column 6: {status_col_values[:10]}")  # Show first 10

        # Check if any contain 'accept' or 'reject'
        status_containing_accept_reject = [val for val in status_col_values if
                                           'accept' in str(val).lower() or 'reject' in str(val).lower()]
        st.write(f"Values containing 'accept' or 'reject': {status_containing_accept_reject}")

        # Create a new DataFrame with the columns we want
        processed_data = pd.DataFrame({
            'Date': data.iloc[:, 0],
            'Status': data.iloc[:, 6],
            'Employee': data.iloc[:, 1].astype(str),
            'Mold': data.iloc[:, 3],
            'Reason': data.iloc[:, 7]
        })

        st.write(f"🔍 After creating processed_data: {len(processed_data)} rows")

        # Show sample of processed data before filtering
        with st.expander("🔍 Debug: Sample processed data (before filtering)"):
            st.write(processed_data.head(10))

        # Handle merged cells
        processed_data['Date'] = processed_data['Date'].fillna(method='ffill')
        processed_data['Mold'] = processed_data['Mold'].fillna(method='ffill')

        # Check status values before filtering
        st.write(f"🔍 Status values before filtering:")
        st.write(processed_data['Status'].value_counts())

        # Remove any rows where Status is not 'accepted' or 'rejected'
        before_filter = len(processed_data)
        processed_data = processed_data[
            processed_data['Status'].str.lower().str.contains('accept|reject', na=False)
        ].copy()
        after_filter = len(processed_data)

        st.write(
            f"🔍 Filtering results: {before_filter} rows → {after_filter} rows (removed {before_filter - after_filter})")

        if after_filter == 0:
            st.error("⚠️ NO DATA after filtering! Check if status values match 'accept' or 'reject'")
            st.write("Available status values were:",
                     processed_data['Status'].unique() if before_filter > 0 else "None")

        # Clean up status values
        processed_data['Status'] = processed_data['Status'].str.lower()
        processed_data['Status'] = processed_data['Status'].replace('acceppted', 'accepted')
        processed_data['Status'] = processed_data['Status'].replace('accept', 'accepted')
        processed_data['Status'] = processed_data['Status'].replace('reject', 'rejected')

        # Standardize status values
        processed_data.loc[processed_data['Status'].str.lower().str.contains('accept'), 'Status'] = 'accepted'
        processed_data.loc[processed_data['Status'].str.lower().str.contains('reject'), 'Status'] = 'rejected'

        # Clean up rejection reasons
        processed_data['Reason'] = processed_data['Reason'].astype(str).str.strip()
        processed_data.loc[processed_data['Status'] == 'accepted', 'Reason'] = ''
        processed_data.loc[(processed_data['Status'] != 'accepted') &
                           ((processed_data['Reason'] == '') |
                            (processed_data['Reason'] == 'nan') |
                            (processed_data['Reason'] == 'None') |
                            (processed_data['Reason'].isna())), 'Reason'] = 'Unknown'

        # Convert Date column to datetime
        processed_data['Date'] = pd.to_datetime(processed_data['Date'], errors='coerce')

        st.write(f"✅ **FINAL: Optic QC processed {len(processed_data)} rows**")
        if len(processed_data) > 0:
            st.write(f"Status distribution: {processed_data['Status'].value_counts().to_dict()}")

        return processed_data

    except Exception as e:
        st.error(f"Error processing Optic QC data: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])

def process_fta_qc(df):
    """Process FTA QC sheet with specific structure"""
    try:
        start_row = find_data_start_row(df)

        if len(df) <= start_row:
            return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason', 'Configuration'])

        data = df.iloc[start_row:].copy()

        # Find status column
        status_col_index = None
        for col_idx in range(data.shape[1]):
            col_values = data.iloc[:, col_idx].dropna().astype(str).str.lower()
            if any(val in ['accept', 'accepted', 'reject', 'rejected', 'acceppted'] for val in col_values):
                status_col_index = col_idx
                break

        if status_col_index is None:
            status_col_index = 6

        # Define column indices for FTA
        date_col_index = 0
        employee_col_index = 1
        mold_col_index = 2  # Different for FTA
        configuration_col_index = 3  # Added for FTA configuration
        reason_col_index = min(status_col_index + 1, data.shape[1] - 1)

        # Create processed dataframe
        processed_data = pd.DataFrame({
            'Date': data.iloc[:, date_col_index],
            'Status': data.iloc[:, status_col_index],
            'Employee': data.iloc[:, employee_col_index].astype(str),
            'Mold': data.iloc[:, mold_col_index] if mold_col_index < data.shape[1] else None,
            'Configuration': data.iloc[:, configuration_col_index] if configuration_col_index < data.shape[1] else None,
            'Reason': data.iloc[:, reason_col_index] if reason_col_index < data.shape[1] else None
        })

        # Handle merged cells
        processed_data['Date'] = processed_data['Date'].fillna(method='ffill')
        processed_data['Mold'] = processed_data['Mold'].fillna(method='ffill')
        processed_data['Configuration'] = processed_data['Configuration'].fillna(method='ffill')

        # Filter for accept/reject only
        processed_data = processed_data[
            processed_data['Status'].str.lower().str.contains('accept|reject', na=False)
        ].copy()

        # Clean up status values
        processed_data['Status'] = processed_data['Status'].str.lower()
        processed_data['Status'] = processed_data['Status'].replace('acceppted', 'accepted')
        processed_data['Status'] = processed_data['Status'].replace('accept', 'accepted')
        processed_data['Status'] = processed_data['Status'].replace('reject', 'rejected')

        # Clean up rejection reasons
        processed_data['Reason'] = processed_data['Reason'].astype(str).str.strip()
        processed_data.loc[processed_data['Status'] == 'accepted', 'Reason'] = ''
        processed_data.loc[(processed_data['Status'] != 'accepted') &
                           ((processed_data['Reason'] == '') |
                            (processed_data['Reason'] == 'nan') |
                            (processed_data['Reason'] == 'None') |
                            (processed_data['Reason'].isna())), 'Reason'] = 'Unknown'

        # Convert Date column to datetime
        processed_data['Date'] = pd.to_datetime(processed_data['Date'], errors='coerce')

        return processed_data

    except Exception as e:
        st.error(f"Error processing FTA QC data: {str(e)}")
        return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason', 'Configuration'])


def process_pl_qc(df, filename=None):
    """Process PL QC sheet - headers in rows 4&5, data starts row 6"""
    try:
        if len(df) <= 6:
            return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])

        # Extract month from filename for PL data
        file_date = None
        if filename:
            file_date = extract_month_from_filename(filename)
            if file_date:
                st.write(f"📅 PL QC: Using date from filename: {file_date.strftime('%B %Y')}")

        # Data starts at row 6 (0-indexed), headers at row 5
        data = df.iloc[6:].copy()

        # Define exact column indices based on analysis
        date_col = 0  # Date (we'll ignore this and use filename date)
        employee_col = 1  # Inspected by
        status_col = 2  # Status (r=7.71)
        approved_col = 3  # Number of approved parts
        rejected_col = 4  # Number of rejected parts

        # Rejection reason columns start at column 5
        reason_columns = {
            5: 'Bubbles and voids',
            6: 'Embedded Particle',
            7: 'Tears',
            8: 'Surface Irregularity',
            9: 'Not cured',
            10: 'Other'
        }

        processed_rows = []

        for idx, row in data.iterrows():
            try:
                # Use filename date if available, otherwise try to get from sheet
                if file_date:
                    current_date = file_date
                else:
                    # Fallback to sheet date if filename parsing failed
                    if pd.notna(row.iloc[date_col]):
                        current_date = pd.to_datetime(row.iloc[date_col], errors='coerce')
                    else:
                        continue

                if pd.isna(current_date):
                    continue

                employee = str(row.iloc[employee_col]).strip().title() if pd.notna(
                    row.iloc[employee_col]) else 'Unknown'
                if employee in ['', 'Nan', 'None']:
                    continue

                status = str(row.iloc[status_col]) if pd.notna(row.iloc[status_col]) else ''
                if status in ['', 'nan', 'None']:
                    continue

                # Extract mold from status (r=7.71 -> 7.71)
                mold = status.replace('r=', '').replace('R=', '').strip()

                # Get counts
                approved = int(float(row.iloc[approved_col])) if pd.notna(row.iloc[approved_col]) and row.iloc[
                    approved_col] != '' else 0
                total_rejected = int(float(row.iloc[rejected_col])) if pd.notna(row.iloc[rejected_col]) and row.iloc[
                    rejected_col] != '' else 0

                # Add accepted entries
                for _ in range(approved):
                    processed_rows.append({
                        'Date': current_date,
                        'Employee': employee,
                        'Status': 'accepted',
                        'Mold': mold,
                        'Reason': ''
                    })

                # Count specific rejection reasons
                specific_rejections = 0
                for col_idx, reason_name in reason_columns.items():
                    if col_idx < len(row):
                        count = int(float(row.iloc[col_idx])) if pd.notna(row.iloc[col_idx]) and row.iloc[
                            col_idx] != '' else 0
                        specific_rejections += count
                        for _ in range(count):
                            processed_rows.append({
                                'Date': current_date,
                                'Employee': employee,
                                'Status': 'rejected',
                                'Mold': mold,
                                'Reason': reason_name
                            })

                # If total rejected > specific rejections, add the difference as "N/A"
                unspecified_rejections = max(0, total_rejected - specific_rejections)
                for _ in range(unspecified_rejections):
                    processed_rows.append({
                        'Date': current_date,
                        'Employee': employee,
                        'Status': 'rejected',
                        'Mold': mold,
                        'Reason': 'N/A'
                    })

            except Exception as e:
                print(f"Error processing row {idx}: {e}")
                continue

        return pd.DataFrame(processed_rows) if processed_rows else pd.DataFrame(
            columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])

    except Exception as e:
        st.error(f"Error processing PL QC: {e}")
        return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])


def process_wings_qc(df, filename=None):
    """Process Wings QC sheet - headers in rows 4&5, data starts row 6"""
    try:
        if len(df) <= 6:
            return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])

        # Extract month from filename for Wings data
        file_date = None
        if filename:
            file_date = extract_month_from_filename(filename)
            if file_date:
                st.write(f"📅 Wings QC: Using date from filename: {file_date.strftime('%B %Y')}")

        # Data starts at row 6 (0-indexed), headers at row 5
        data = df.iloc[6:].copy()

        # Define exact column indices based on analysis
        date_col = 0  # Date (we'll ignore this and use filename date)
        employee_col = 1  # Inspected by
        approved_col = 2  # Number of approved parts
        rejected_col = 3  # Number of rejected parts

        # Rejection reason columns start at column 4
        reason_columns = {
            4: 'Bubbles and voids',
            5: 'Embedded Particle',
            6: 'Tears',
            7: 'lack of uniformity in pigment',
            8: 'Not cured',
            9: 'Embedded Metals',
            10: 'Other'
        }

        processed_rows = []

        for idx, row in data.iterrows():
            try:
                # Use filename date if available, otherwise try to get from sheet
                if file_date:
                    date = file_date
                else:
                    # Fallback to sheet date if filename parsing failed
                    date = pd.to_datetime(row.iloc[date_col], errors='coerce')

                if pd.isna(date):
                    continue

                employee = str(row.iloc[employee_col]).strip().title() if pd.notna(
                    row.iloc[employee_col]) else 'Unknown'

                # Get counts
                approved = int(float(row.iloc[approved_col])) if pd.notna(row.iloc[approved_col]) else 0

                # Add accepted entries
                for _ in range(approved):
                    processed_rows.append({
                        'Date': date,
                        'Employee': employee,
                        'Status': 'accepted',
                        'Mold': 'N/A',
                        'Reason': ''
                    })

                # Add rejected entries
                for col_idx, reason_name in reason_columns.items():
                    if col_idx < len(row):
                        count = int(float(row.iloc[col_idx])) if pd.notna(row.iloc[col_idx]) else 0
                        for _ in range(count):
                            processed_rows.append({
                                'Date': date,
                                'Employee': employee,
                                'Status': 'rejected',
                                'Mold': 'N/A',
                                'Reason': reason_name
                            })

            except Exception as e:
                continue

        return pd.DataFrame(processed_rows) if processed_rows else pd.DataFrame(
            columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])

    except Exception as e:
        st.error(f"Error processing Wings QC: {e}")
        return pd.DataFrame(columns=['Date', 'Status', 'Employee', 'Mold', 'Reason'])


def merge_rejection_reasons(df, component_name):
    """
    Merge rejection reasons based on component-specific rules.
    This is applied BEFORE any chart generation.
    """
    if df.empty or 'Reason' not in df.columns:
        return df

    df = df.copy()

    # Only process rejected parts
    rejected_mask = df['Status'] == 'rejected'

    # Create a normalized version for comparison (lowercase, stripped)
    df['Reason_normalized'] = df['Reason'].astype(str).str.lower().str.strip()

    if component_name == "FTA":
        # FTA merging rules:
        # 1. 'bubbles and voids' + 'tears' + 'hole in sm' = 'Tears/ Voids'
        fta_merge_1 = ['bubbles and voids', 'hole in sm']
        df.loc[rejected_mask & df['Reason_normalized'].isin(fta_merge_1), 'Reason'] = 'bubbles and voids'

        # 2. 'embedded metals' + 'embedded particle' = 'Embedded particle'
        fta_merge_2 = ['embedded metals', 'embedded particle']
        df.loc[rejected_mask & df['Reason_normalized'].isin(fta_merge_2), 'Reason'] = 'Embedded particle'

        # 3. 'surface irregularity' + 'not cured' + 'other' + 'n/a' = 'Other'
        fta_merge_3 = ['surface irregularity', 'not cured', 'other', 'n/a', 'na']
        df.loc[rejected_mask & df['Reason_normalized'].isin(fta_merge_3), 'Reason'] = 'Other'

    elif component_name in ["Optic", "PL"]:
        # Optic and PL merging rules:
        # 1. 'not cured' + 'other' + 'N/A' = 'Other'
        optic_pl_merge_1 = ['not cured', 'other', 'n/a', 'na']
        df.loc[rejected_mask & df['Reason_normalized'].isin(optic_pl_merge_1), 'Reason'] = 'Other'

        # 2. 'surface irregularity' + 'bubbles and voids' = 'Bubbles and voids'
        optic_pl_merge_2 = ['surface irregularity', 'bubbles and voids']
        df.loc[rejected_mask & df['Reason_normalized'].isin(optic_pl_merge_2), 'Reason'] = 'Bubbles and voids'

    elif component_name == "Wings":
        # Wings merging rules:
        # 1. 'embedded metals' + 'embedded particle' = 'Embedded particle'
        wings_merge_1 = ['embedded metals', 'embedded particle']
        df.loc[rejected_mask & df['Reason_normalized'].isin(wings_merge_1), 'Reason'] = 'Embedded particle'

    # Drop the temporary normalized column
    df = df.drop(columns=['Reason_normalized'])

    return df

def normalize_employee_names(df):
    """Normalize employee names to handle case sensitivity and spacing issues"""
    if not df.empty and 'Employee' in df.columns:
        df['Employee'] = df['Employee'].astype(str).str.lower().str.strip()
        df['Employee'] = df['Employee'].str.replace(r'\s+', ' ', regex=True)
        df['Employee'] = df['Employee'].str.title()
    return df


# Main processing starts here
if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully!")

    # Process all uploaded files
    all_optic_data = []
    all_fta_data = []
    all_pl_data = []
    all_wings_data = []

    for i, file in enumerate(uploaded_files):
        st.write(f"📊 Processing file {i + 1}: {file.name}")

        try:
            # Read Excel file to check available sheets
            excel_file = pd.ExcelFile(file, engine='openpyxl')
            available_sheets = excel_file.sheet_names
            st.write(f"Available sheets: {available_sheets}")

            # Process Optic QC if available
            if 'Optic QC' in available_sheets:
                optic_df = pd.read_excel(file, sheet_name='Optic QC', header=None, engine='openpyxl')
                optic_processed = process_optic_qc(optic_df)
                if not optic_processed.empty:
                    optic_processed = standardize_rejection_reasons(optic_processed)
                    optic_processed['File'] = file.name
                    all_optic_data.append(optic_processed)
                    st.write(f"✅ Processed {len(optic_processed)} Optic QC records")

            # Process FTA QC if available
            if 'FTA QC' in available_sheets:
                fta_df = pd.read_excel(file, sheet_name='FTA QC', header=None, engine='openpyxl')
                fta_processed = process_fta_qc(fta_df)
                if not fta_processed.empty:
                    fta_processed = standardize_rejection_reasons(fta_processed)
                    fta_processed['File'] = file.name
                    all_fta_data.append(fta_processed)
                    st.write(f"✅ Processed {len(fta_processed)} FTA QC records")

            # Process Optic QC if available
            if 'Optic QC' in available_sheets:
                optic_df = pd.read_excel(file, sheet_name='Optic QC', header=None, engine='openpyxl')
                optic_processed = process_optic_qc(optic_df)
                if not optic_processed.empty:
                    optic_processed = standardize_rejection_reasons(optic_processed)
                    optic_processed = merge_rejection_reasons(optic_processed, "Optic")
                    # ADD THIS DEBUG CODE:
                    st.write("🔍 DEBUG: Optic reasons after merging:")
                    st.write(optic_processed[optic_processed['Status'] == 'rejected']['Reason'].value_counts())
                    optic_processed['File'] = file.name
                    all_optic_data.append(optic_processed)
                    st.write(f"✅ Processed {len(optic_processed)} Optic QC records")

            # Process FTA QC if available
            if 'FTA QC' in available_sheets:
                fta_df = pd.read_excel(file, sheet_name='FTA QC', header=None, engine='openpyxl')
                fta_processed = process_fta_qc(fta_df)
                if not fta_processed.empty:
                    fta_processed = standardize_rejection_reasons(fta_processed)
                    fta_processed = merge_rejection_reasons(fta_processed, "FTA")
                    # ADD THIS DEBUG CODE:
                    st.write("🔍 DEBUG: FTA reasons after merging:")
                    st.write(fta_processed[fta_processed['Status'] == 'rejected']['Reason'].value_counts())
                    fta_processed['File'] = file.name
                    all_fta_data.append(fta_processed)
                    st.write(f"✅ Processed {len(fta_processed)} FTA QC records")

            # Process PL QC if available
            if 'PL QC' in available_sheets:
                pl_df = pd.read_excel(file, sheet_name='PL QC', header=None, engine='openpyxl')
                pl_processed = process_pl_qc(pl_df, filename=file.name)  # PASS FILENAME
                if not pl_processed.empty:
                    pl_processed = standardize_rejection_reasons(pl_processed)
                    pl_processed = merge_rejection_reasons(pl_processed, "PL")
                    # ADD THIS DEBUG CODE:
                    st.write("🔍 DEBUG: pl reasons after merging:")
                    st.write(pl_processed[pl_processed['Status'] == 'rejected']['Reason'].value_counts())
                    pl_processed['File'] = file.name
                    all_pl_data.append(pl_processed)
                    st.write(f"✅ Processed {len(pl_processed)} PL QC records")

            # Process Wings QC if available
            if 'Wings QC' in available_sheets:
                wings_df = pd.read_excel(file, sheet_name='Wings QC', header=None, engine='openpyxl')
                wings_processed = process_wings_qc(wings_df, filename=file.name)  # PASS FILENAME
                if not wings_processed.empty:
                    wings_processed = standardize_rejection_reasons(wings_processed)
                    wings_processed = merge_rejection_reasons(wings_processed, "Wings")
                    # ADD THIS DEBUG CODE:
                    st.write("🔍 DEBUG: wings reasons after merging:")
                    st.write(wings_processed[wings_processed['Status'] == 'rejected']['Reason'].value_counts())
                    wings_processed['File'] = file.name
                    all_wings_data.append(wings_processed)
                    st.write(f"✅ Processed {len(wings_processed)} Wings QC records")

        except Exception as e:
            st.error(f"❌ Error processing {file.name}: {str(e)}")

    # Combine all data
    optic_data = pd.concat(all_optic_data, ignore_index=True) if all_optic_data else pd.DataFrame()
    fta_data = pd.concat(all_fta_data, ignore_index=True) if all_fta_data else pd.DataFrame()
    pl_data = pd.concat(all_pl_data, ignore_index=True) if all_pl_data else pd.DataFrame()
    wings_data = pd.concat(all_wings_data, ignore_index=True) if all_wings_data else pd.DataFrame()

    # Normalize employee names
    optic_data = normalize_employee_names(optic_data)
    fta_data = normalize_employee_names(fta_data)
    pl_data = normalize_employee_names(pl_data)
    wings_data = normalize_employee_names(wings_data)

    # After normalizing employee names, add this for each dataset:
    #optic_data = clean_rejection_reasons(optic_data)
    #fta_data = clean_rejection_reasons(fta_data)
    #pl_data = clean_rejection_reasons(pl_data)
    #wings_data = clean_rejection_reasons(wings_data)

    st.success(
        f"📈 Combined data: {len(optic_data)} Optic QC records, {len(fta_data)} FTA QC records, {len(pl_data)} PL QC records, {len(wings_data)} Wings QC records")

    # Enhanced Time Analysis Controls
    st.subheader("📅 Time Analysis Controls")

    # Get date range from combined data
    all_dates = pd.concat([
        optic_data['Date'] if not optic_data.empty else pd.Series(),
        fta_data['Date'] if not fta_data.empty else pd.Series(),
        pl_data['Date'] if not pl_data.empty else pd.Series(),
        wings_data['Date'] if not wings_data.empty else pd.Series()
    ]).dropna()

    if len(all_dates) > 0:
        startDate = all_dates.min()
        endDate = all_dates.max()
    else:
        startDate = datetime.today()
        endDate = datetime.today()

    # Date and grouping controls
    date_col1, date_col2, date_col3 = st.columns([2, 2, 2])

    with date_col1:
        date1 = st.date_input("Start Date", startDate)

    with date_col2:
        end_date = st.date_input("End Date", endDate)
        date2 = pd.Timestamp(end_date)

    with date_col3:
        time_grouping = st.selectbox(
            "Group Data By:",
            ["Whole Date Range", "Weekly", "Monthly", "Last 3 Months", "Quarterly"],
            index=0,
            help="Choose how to aggregate your data within the selected date range"
        )

    # Show user what they've selected
    date_range_days = (pd.Timestamp(date2) - pd.Timestamp(date1)).days
    st.info(f"📊 Analyzing {date_range_days} days of data, grouped by {time_grouping.lower()}")

    # Convert filter dates
    date1 = pd.to_datetime(date1).date()
    date2 = pd.to_datetime(date2).date()


    # Function to aggregate data based on selected time grouping
    def aggregate_data_by_period(data, grouping_type):
        """Aggregate data by the selected time period"""
        if data.empty or 'Date' not in data.columns:
            return data

        data = data.copy()
        data['Date'] = pd.to_datetime(data['Date'])

        # Create period grouping
        if grouping_type == "Whole Date Range":
            # No grouping - treat all data as one period
            data['Period'] = "Full Range"
            data[
                'Period_Label'] = f"{data['Date'].min().strftime('%Y-%m-%d')} to {data['Date'].max().strftime('%Y-%m-%d')}"
            data['Sort_Date'] = data['Date'].min()
        elif grouping_type == "Weekly":
            data['Period'] = data['Date'].dt.to_period('W')
            data['Period_Label'] = data['Period'].astype(str)
            data['Sort_Date'] = data['Period'].dt.start_time
        elif grouping_type == "Monthly":
            data['Period'] = data['Date'].dt.to_period('M')
            data['Period_Label'] = data['Date'].dt.strftime('%b %Y')  # Format as "Mar 2025"
            data['Sort_Date'] = data['Period'].dt.start_time
        elif grouping_type == "Last 3 Months":
            # Get the latest date in the data
            max_date = data['Date'].max()

            # Filter to only last 3 months from the max date
            three_months_ago = max_date - pd.DateOffset(months=2)  # 2 months back from current = 3 months total
            data = data[data['Date'] >= three_months_ago].copy()

            # Create monthly periods for the last 3 months only
            data['Period'] = data['Date'].dt.to_period('M')
            data['Period_Label'] = data['Date'].dt.strftime('%b %Y')  # Format as "Mar 2025"
            data['Sort_Date'] = data['Period'].dt.start_time

            # Add months back calculation for proper ordering
            data['MonthsBack'] = ((max_date.year - data['Date'].dt.year) * 12 +
                                  (max_date.month - data['Date'].dt.month))

        elif grouping_type == "Quarterly":
            data['Period'] = data['Date'].dt.to_period('Q')
            data['Period_Label'] = data['Period'].astype(str)
            data['Sort_Date'] = data['Period'].dt.start_time

        return data


    # Filter data by date first
    if not optic_data.empty:
        optic_data['Date'] = pd.to_datetime(optic_data['Date']).dt.date
        optic_data = optic_data[(optic_data['Date'] >= date1) & (optic_data['Date'] <= date2)].copy()
        # Apply time grouping
        optic_data = aggregate_data_by_period(optic_data, time_grouping)
        # Apply merging AGAIN after filtering (IMPORTANT!)
        optic_data = merge_rejection_reasons(optic_data, "Optic")

    if not fta_data.empty:
        fta_data['Date'] = pd.to_datetime(fta_data['Date']).dt.date
        fta_data = fta_data[(fta_data['Date'] >= date1) & (fta_data['Date'] <= date2)].copy()
        # Apply time grouping
        fta_data = aggregate_data_by_period(fta_data, time_grouping)
        # Apply merging AGAIN after filtering (IMPORTANT!)
        fta_data = merge_rejection_reasons(fta_data, "FTA")

    if not pl_data.empty:
        pl_data['Date'] = pd.to_datetime(pl_data['Date']).dt.date
        pl_data = pl_data[(pl_data['Date'] >= date1) & (pl_data['Date'] <= date2)].copy()
        # Apply time grouping
        pl_data = aggregate_data_by_period(pl_data, time_grouping)
        # Apply merging AGAIN after filtering (IMPORTANT!)
        pl_data = merge_rejection_reasons(pl_data, "PL")

    if not wings_data.empty:
        wings_data['Date'] = pd.to_datetime(wings_data['Date']).dt.date
        wings_data = wings_data[(wings_data['Date'] >= date1) & (wings_data['Date'] <= date2)].copy()
        # Apply time grouping
        wings_data = aggregate_data_by_period(wings_data, time_grouping)
        # Apply merging AGAIN after filtering (IMPORTANT!)
        wings_data = merge_rejection_reasons(wings_data, "Wings")

    # Sidebar filters
    st.sidebar.header("Choose your filter:")

    # Part filter
    part_options = ["Optic", "FTA", "PL", "Wings"]
    selected_parts = st.sidebar.multiselect("Select Part", part_options, default=part_options)

    # Employee filter
   #all_employees = set()
   #if "Optic" in selected_parts and not optic_data.empty:
   #    all_employees.update(optic_data['Employee'].unique())
   #if "FTA" in selected_parts and not fta_data.empty:
   #    all_employees.update(fta_data['Employee'].unique())
   #if "PL" in selected_parts and not pl_data.empty:
   #    all_employees.update(pl_data['Employee'].unique())
   #if "Wings" in selected_parts and not wings_data.empty:
   #    all_employees.update(wings_data['Employee'].unique())

   #all_employees = sorted([emp for emp in all_employees if pd.notna(emp)])
   #selected_employees = st.sidebar.multiselect("Select Employees", all_employees)

    # Apply employee filter
   #if selected_employees:
   #    if not optic_data.empty:
   #        optic_data = optic_data[optic_data['Employee'].isin(selected_employees)]
   #    if not fta_data.empty:
   #        fta_data = fta_data[fta_data['Employee'].isin(selected_employees)]
   #    if not pl_data.empty:
   #        pl_data = pl_data[pl_data['Employee'].isin(selected_employees)]
   #    if not wings_data.empty:
   #        wings_data = wings_data[wings_data['Employee'].isin(selected_employees)]

    # Filter data based on selected parts
    filtered_data = {}
    if "Optic" in selected_parts:
        filtered_data["Optic"] = optic_data
    if "FTA" in selected_parts:
        filtered_data["FTA"] = fta_data
    if "PL" in selected_parts:
        filtered_data["PL"] = pl_data
    if "Wings" in selected_parts:
        filtered_data["Wings"] = wings_data

    # Show data info
    #date_range_days = (pd.Timestamp(date2) - pd.Timestamp(date1)).days
    #st.info(f"📊 Analyzing {date_range_days} days of data, grouped by {time_grouping.lower()}")

    # SUMMARY STATISTICS (MOVED TO TOP)
    st.subheader("📊 Summary Statistics")

    # Combine all data for the summary
    summary_stats = []
    for component_name, data in filtered_data.items():
        if not data.empty:
            total = len(data)
            accepted = len(data[data['Status'] == 'accepted'])
            rejected = len(data[data['Status'] == 'rejected'])
            inspected = accepted + rejected

            summary_stats.append({
                'Component': component_name,
                'Total Parts': total,
                'Inspected Parts': inspected,
                'Accepted': accepted,
                'Rejected': rejected,
                'Acceptance Rate (%)': round((accepted / inspected * 100)) if inspected > 0 else 0,
                'Rejection Rate (%)': round((rejected / inspected * 100)) if inspected > 0 else 0
            })

    if summary_stats:
        summary_df = pd.DataFrame(summary_stats)
        st.dataframe(summary_df, use_container_width=True)

        # Download option for summary
        csv_summary = summary_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Download Summary Statistics CSV",
            csv_summary,
            "summary_statistics.csv",
            "text/csv"
        )
    else:
        st.info("No data available for summary statistics")

    # TIME-BASED TREND ANALYSIS (Now appears after Summary Statistics)
    st.subheader(f"📈 {time_grouping} Trend Analysis")


    # ... rest of the trend analysis code continues here ...


    def create_period_statistics(data, component_name, grouping_type):
        """Create statistics grouped by time period"""
        if data.empty or 'Period' not in data.columns:
            return pd.DataFrame()

        period_stats = []

        for period in data['Period'].unique():
            period_data = data[data['Period'] == period]

            total = len(period_data)
            accepted = len(period_data[period_data['Status'] == 'accepted'])
            rejected = len(period_data[period_data['Status'] == 'rejected'])
            inspected = accepted + rejected

            yield_pct = (accepted / inspected * 100) if inspected > 0 else 0
            rejection_pct = (rejected / inspected * 100) if inspected > 0 else 0

            period_stats.append({
                'Period': str(period),
                'Period_Label': period_data['Period_Label'].iloc[
                    0] if 'Period_Label' in period_data.columns else str(period),
                'Sort_Date': period_data['Sort_Date'].iloc[0] if 'Sort_Date' in period_data.columns else
                period_data['Date'].iloc[0],
                'Component': component_name,
                'Total': total,
                'Accepted': accepted,
                'Rejected': rejected,
                'Inspected': inspected,
                'Yield': round(yield_pct),
                'Rejection_Rate': round(rejection_pct)
            })

        result_df = pd.DataFrame(period_stats)
        if not result_df.empty:
            result_df = result_df.sort_values('Sort_Date')

        return result_df


    # Generate period statistics
    period_data_parts = []
    for component_name, data in filtered_data.items():
        if not data.empty:
            period_stats = create_period_statistics(data, component_name, time_grouping)
            if not period_stats.empty:
                period_data_parts.append(period_stats)

    if period_data_parts:
        combined_period_data = pd.concat(period_data_parts, ignore_index=True)

        # Sort by Sort_Date for proper chronological order
        combined_period_data = combined_period_data.sort_values('Sort_Date')

        # Create trend visualizations
        col1, col2 = st.columns(2)

        with col1:
            # Yield trend line chart
            fig_yield_trend = px.line(
                combined_period_data,
                x='Period_Label',
                y='Yield',
                color='Component',
                markers=True,
                title=f"{time_grouping} Yield Trends",
                labels={'Period_Label': time_grouping, 'Yield': 'Yield (%)', 'Component': 'Component'}
            )

            # Add target line at 100%
            fig_yield_trend.add_shape(
                type="line", line=dict(dash="dash", width=2, color="gray"),
                y0=100, y1=100, x0=0, x1=1, xref="paper", yref="y"
            )

            fig_yield_trend.update_layout(
                xaxis_tickangle=-45,
                yaxis_range=[0, 105],
                font_family="Segoe UI Semilight (Body)"
            )
            st.plotly_chart(fig_yield_trend, use_container_width=True)

        with col2:
            # Updated Production volume trend - Accepted parts only, separate bars
            fig_volume_trend = go.Figure()

            # Get unique periods and components
            periods = combined_period_data['Period_Label'].unique()
            components = combined_period_data['Component'].unique()

            # Use colors for all four components
            colors = {'Optic': '#1f77b4', 'FTA': '#aec7e8', 'PL': '#ff7f0e', 'Wings': '#ffbb78'}

            # Add bars for each component
            for component in components:
                component_data = combined_period_data[combined_period_data['Component'] == component]

                fig_volume_trend.add_trace(go.Bar(
                    name=f'{component} Accepted',
                    x=component_data['Period_Label'],
                    y=component_data['Accepted'],
                    marker_color=colors.get(component, '#1f77b4'),
                    text=component_data['Accepted'],
                    textposition='inside',
                    textfont=dict(color='white', size=12),
                    showlegend=True
                ))

            fig_volume_trend.update_layout(
                title=f"{time_grouping} Accepted Parts Production",
                xaxis_title=time_grouping,
                yaxis_title='Accepted Parts',
                barmode='group',
                xaxis_tickangle=-45,
                font_family="Segoe UI Semilight (Body)"
            )
            st.plotly_chart(fig_volume_trend, use_container_width=True)

        # Time-based data table
        with st.expander(f"View {time_grouping} Trend Data"):
            display_cols = ['Period_Label', 'Component', 'Total', 'Accepted', 'Rejected', 'Yield', 'Rejection_Rate']
            st.dataframe(combined_period_data[display_cols], use_container_width=True)
            csv_trends = combined_period_data[display_cols].to_csv(index=False).encode('utf-8')
            st.download_button(
                f"📥 Download {time_grouping} Trends CSV",
                csv_trends,
                f"{time_grouping.lower()}_trends.csv",
                "text/csv"
            )
    else:
        st.info(f"No data available for {time_grouping.lower()} trend analysis")

    # PLOT 2: Mold Configuration Statistics (Grouped with spacing) - UPDATED WITH QUANTITIES
    st.subheader("🔧 Mold Configuration Statistics")

    mold_stats_data = []

    for component_name, data in filtered_data.items():
        if not data.empty and 'Mold' in data.columns and 'Status' in data.columns:
            # Handle FTA with Configuration differently
            if component_name == "FTA" and 'Configuration' in data.columns:
                # Create combined mold identifier for FTA
                data = data.copy()
                data['Mold_Config'] = data['Mold'].astype(str) + '-' + data['Configuration'].astype(str)

                # Group by mold configuration
                for mold_config in data['Mold_Config'].unique():
                    if pd.notna(mold_config):
                        mold_data = data[data['Mold_Config'] == mold_config]
                        total = len(mold_data)
                        accepted = len(mold_data[mold_data['Status'] == 'accepted'])
                        rejected = len(mold_data[mold_data['Status'] == 'rejected'])
                        inspected = accepted + rejected

                        if inspected > 0:
                            acceptance_pct = (accepted / inspected * 100)
                            rejection_pct = (rejected / inspected * 100)

                            mold_stats_data.append({
                                'Component': component_name,
                                'Mold_Display': str(mold_config),
                                'Acceptance_Percentage': round(acceptance_pct),
                                'Rejection_Percentage': round(rejection_pct),
                                'Total_Inspected': inspected,
                                'Accepted': accepted,
                                'Rejected': rejected
                            })
            elif component_name == "Wings":
                # Wings has no mold, so create one entry for all Wings data
                total = len(data)
                accepted = len(data[data['Status'] == 'accepted'])
                rejected = len(data[data['Status'] == 'rejected'])
                inspected = accepted + rejected

                if inspected > 0:
                    acceptance_pct = (accepted / inspected * 100)
                    rejection_pct = (rejected / inspected * 100)

                    mold_stats_data.append({
                        'Component': component_name,
                        'Mold_Display': "All",
                        'Acceptance_Percentage': round(acceptance_pct),
                        'Rejection_Percentage': round(rejection_pct),
                        'Total_Inspected': inspected,
                        'Accepted': accepted,
                        'Rejected': rejected
                    })
            else:
                # Handle Optic and PL - group by mold number
                for mold in data['Mold'].unique():
                    if pd.notna(mold) and str(mold) not in ['N/A', 'nan', 'None', '']:
                        mold_data = data[data['Mold'] == mold]
                        total = len(mold_data)
                        accepted = len(mold_data[mold_data['Status'] == 'accepted'])
                        rejected = len(mold_data[mold_data['Status'] == 'rejected'])
                        inspected = accepted + rejected

                        if inspected > 0:
                            acceptance_pct = (accepted / inspected * 100)
                            rejection_pct = (rejected / inspected * 100)

                            mold_stats_data.append({
                                'Component': component_name,
                                'Mold_Display': str(mold),
                                'Acceptance_Percentage': round(acceptance_pct),
                                'Rejection_Percentage': round(rejection_pct),
                                'Total_Inspected': inspected,
                                'Accepted': accepted,
                                'Rejected': rejected
                            })

    if mold_stats_data:
        mold_df = pd.DataFrame(mold_stats_data)

        # Sort by component and mold for better grouping
        mold_df = mold_df.sort_values(['Component', 'Mold_Display'])

        # Create x-axis labels with spacing between component groups
        x_labels = []
        x_positions = []
        current_pos = 0

        component_order = ['FTA', 'Optic', 'PL', 'Wings']  # Define order

        for i, component in enumerate(component_order):
            component_data = mold_df[mold_df['Component'] == component]
            if not component_data.empty:
                # Add spacing between groups (except for first group)
                if i > 0:
                    current_pos += 1  # Add gap between component groups

                for _, row in component_data.iterrows():
                    x_labels.append(f"{component}-{row['Mold_Display']}")
                    x_positions.append(current_pos)
                    current_pos += 1

        # Create the chart
        fig_mold = go.Figure()

        # Add ACCEPTANCE bars (ALL GREEN) with quantities
        acceptance_y = []
        rejection_y = []
        accepted_quantities = []
        rejected_quantities = []

        for label in x_labels:
            component, mold = label.split('-', 1)
            row_data = mold_df[(mold_df['Component'] == component) & (mold_df['Mold_Display'] == mold)]
            if not row_data.empty:
                acceptance_y.append(row_data.iloc[0]['Acceptance_Percentage'])
                rejection_y.append(row_data.iloc[0]['Rejection_Percentage'])
                accepted_quantities.append(row_data.iloc[0]['Accepted'])
                rejected_quantities.append(row_data.iloc[0]['Rejected'])
            else:
                acceptance_y.append(0)
                rejection_y.append(0)
                accepted_quantities.append(0)
                rejected_quantities.append(0)

        # Acceptance bars - ALL GREEN with quantity in text
        fig_mold.add_trace(go.Bar(
            name='Acceptance %',
            x=x_positions,
            y=acceptance_y,
            marker_color='#28a745',  # Green for all acceptance
            text=[f'{val:.0f}% ({qty})' for val, qty in zip(acceptance_y, accepted_quantities)],
            textposition='inside',
            customdata=list(zip(x_labels, accepted_quantities)),
            hovertemplate='<b>%{customdata[0]}</b><br>Acceptance: %{y:.1f}%<br>Quantity: %{customdata[1]}<extra></extra>'
        ))

        # Rejection bars - ALL RED with quantity in text
        fig_mold.add_trace(go.Bar(
            name='Rejection %',
            x=x_positions,
            y=rejection_y,
            marker_color='#dc3545',  # Red for all rejection
            text=[f'{val:.0f}% ({qty})' for val, qty in zip(rejection_y, rejected_quantities)],
            textposition='inside',
            customdata=list(zip(x_labels, rejected_quantities)),
            hovertemplate='<b>%{customdata[0]}</b><br>Rejection: %{y:.1f}%<br>Quantity: %{customdata[1]}<extra></extra>'
        ))

        fig_mold.update_layout(
            title='Mold Configuration Performance by Component (Grouped)',
            xaxis_title='Mold Configuration (Grouped by Component)',
            yaxis_title='Percentage (%)',
            barmode='stack',
            yaxis_range=[0, 100],
            height=500,
            font_family="Segoe UI Semilight (Body)",
            xaxis=dict(
                tickmode='array',
                tickvals=x_positions,
                ticktext=x_labels,
                tickangle=-45
            )
        )

        st.plotly_chart(fig_mold, use_container_width=True)

    # PLOT 3: Rejection Reasons Analysis
    st.subheader("❌ Rejection Reasons Analysis")


    def create_rejection_analysis_plot(component_name, data, time_grouping):
        """Create rejection analysis plot for a specific component with bar and pie charts"""
        if data.empty or 'Status' not in data.columns or 'Reason' not in data.columns:
            return None, None, None, None, None

        # Store production volume data for annotations
        production_volume_data = {}

        # If we have time grouping, process by period
        if time_grouping != "Whole Date Range" and 'Period' in data.columns:
            periods = sorted(data['Period'].unique())

            all_period_data = []
            period_totals = {}

            for period in periods:
                period_data = data[data['Period'] == period]
                total_parts = len(period_data)
                total_accepted = len(period_data[period_data['Status'] == 'accepted'])
                total_rejected = len(period_data[period_data['Status'] == 'rejected'])
                total_rejection_rate = (total_rejected / total_parts * 100) if total_parts > 0 else 0

                period_label = period_data['Period_Label'].iloc[0] if 'Period_Label' in period_data.columns else str(
                    period)

                # ADD THIS: Store accepted data too
                all_period_data.append({
                    'Period': period_label,
                    'Reason': 'Accepted',  # Add accepted as a "reason"
                    'Count': total_accepted,
                    'Percentage': round((total_accepted / total_parts * 100) if total_parts > 0 else 0),
                    'Rejected_Count': 0
                })

                # Get rejection reasons for this period
                rejected_data = period_data[period_data['Status'] == 'rejected']

                if not rejected_data.empty:
                    reason_counts = rejected_data['Reason'].value_counts()

                    for reason, count in reason_counts.items():
                        if pd.notna(reason) and str(reason).strip() != '' and str(reason).strip().lower() != 'unknown':
                            percentage = (count / total_parts * 100) if total_parts > 0 else 0

                            all_period_data.append({
                                'Period': period_label,
                                'Reason': str(reason),
                                'Count': count,
                                'Percentage': round(percentage),
                                'Rejected_Count': count
                            })

            if all_period_data:
                reasons_df = pd.DataFrame(all_period_data)
                reasons_df = reasons_df.sort_values('Period')

                # Create grouped bar chart with consistent colors
                # Get unique reasons and create consistent color mapping
                unique_reasons = reasons_df['Reason'].unique()
                color_sequence = create_color_sequence_for_reasons(unique_reasons)
                fig_bar = px.bar(
                    reasons_df,
                    x='Period',
                    y='Percentage',
                    color='Reason',
                    title=f'{component_name} - Rejection Reasons by {time_grouping} (% of Total Parts)',
                    labels={'Period': time_grouping, 'Percentage': 'Percentage of Total Parts (%)',
                            'Reason': 'Rejection Reason'},
                    color_discrete_sequence=color_sequence,
                    barmode='stack',
                    hover_data={'Rejected_Count': True},
                    text=[f'{row["Percentage"]}% ({row["Count"]})' for _, row in reasons_df.iterrows()],
                )

                # Add sorting function and apply chronological order
                def sort_month_periods(periods):
                    """Sort month-year labels chronologically"""
                    try:
                        period_dates = []
                        for period in periods:
                            try:
                                date_obj = pd.to_datetime(period, format='%b %Y')
                                period_dates.append((date_obj, period))
                            except:
                                period_dates.append((pd.to_datetime('1900-01-01'), period))

                        period_dates.sort(key=lambda x: x[0])
                        return [period for _, period in period_dates]
                    except:
                        return sorted(periods)

                sorted_periods = sort_month_periods(reasons_df['Period'].unique())
                fig_bar.update_xaxes(categoryorder='array', categoryarray=sorted_periods)
                fig_bar.update_traces(
                    hovertemplate='<b>%{fullData.name}</b><br>' +
                                  f'{time_grouping}: %{{x}}<br>' +
                                  'Percentage of Total Parts: %{y}%<br>' +
                                  'Rejected Quantity: %{customdata[0]}<extra></extra>'
                )

                # Add total rejection annotations with production volume
                for period_label, info in period_totals.items():
                    component_data = reasons_df[reasons_df['Period'] == period_label]
                    total_height = component_data['Percentage'].sum()

                    fig_bar.add_annotation(
                        x=period_label,
                        y=total_height + 3,  # Raised higher
                        text=f"{info['rejection_rate']}% ({info['total_produced']})",  # Shortened format
                        showarrow=False,
                        font=dict(size=10, color="black", family="Segoe UI Semilight (Body)"),
                        bgcolor="rgba(255,255,255,0.8)",
                        bordercolor="black",
                        borderwidth=1,
                        xanchor="center"
                    )

                # Calculate max height for y-axis
                max_height = max([sum(reasons_df[reasons_df['Period'] == period]['Percentage'])
                                  for period in reasons_df['Period'].unique()])

                fig_bar.update_layout(
                    yaxis_range=[0, max_height + 10],  # More space for annotations
                    height=400,
                    xaxis_tickangle=-45,
                    font_family="Segoe UI Semilight (Body)"
                )

                # Create pie chart for overall rejection reasons
                overall_rejected = data[data['Status'] == 'rejected']
                overall_accepted = data[data['Status'] == 'accepted']

                if not overall_rejected.empty:
                    overall_reason_counts = overall_rejected['Reason'].value_counts()

                    # Filter out empty/unknown reasons for pie chart
                    filtered_reasons = {k: v for k, v in overall_reason_counts.items()
                                        if pd.notna(k) and str(k).strip() != '' and str(k).strip().lower() != 'unknown'}

                    if filtered_reasons:
                        # Get unique reasons and create consistent color mapping
                        unique_reasons = list(filtered_reasons.keys())
                        color_sequence = create_color_sequence_for_reasons(unique_reasons)

                        fig_pie = px.pie(
                            values=list(filtered_reasons.values()),
                            names=list(filtered_reasons.keys()),
                            title=f'{component_name} - Overall Rejection Reasons Distribution',
                            color_discrete_sequence=color_sequence
                        )
                        fig_pie.update_layout(
                            font_family="Segoe UI Semilight (Body)",
                            height=400
                        )
                        fig_pie.update_traces(
                            textinfo='label+percent',
                            texttemplate='%{label}<br>%{percent:.0%}',
                            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.0%}<extra></extra>'
                        )

                        # NEW: Create pie chart with accepted + rejected breakdown
                        total_accepted = len(overall_accepted)

                        # Combine accepted count with rejection reasons
                        combined_data = {'Accepted': total_accepted}
                        combined_data.update(filtered_reasons)

                        # Get colors for combined chart
                        combined_reasons = list(combined_data.keys())
                        combined_color_sequence = create_color_sequence_for_reasons(combined_reasons)

                        fig_pie_combined = px.pie(
                            values=list(combined_data.values()),
                            names=list(combined_data.keys()),
                            title=f'{component_name} - Accepted vs Rejected Breakdown',
                            color_discrete_sequence=combined_color_sequence
                        )
                        fig_pie_combined.update_layout(
                            font_family="Segoe UI Semilight (Body)",
                            height=400
                        )
                        fig_pie_combined.update_traces(
                            textinfo='label+percent',
                            texttemplate='%{label}<br>%{percent:.0%}',
                            hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.0%}<extra></extra>'
                        )
                    else:
                        fig_pie = None
                        fig_pie_combined = None
                else:
                    fig_pie = None
                    fig_pie_combined = None
                return fig_bar, fig_pie, fig_pie_combined, reasons_df, production_volume_data  # UPDATED

        else:
            # Whole date range - single bar
            total_parts = len(data)
            total_accepted = len(data[data['Status'] == 'accepted'])
            total_rejected = len(data[data['Status'] == 'rejected'])
            total_rejection_rate = (total_rejected / total_parts * 100) if total_parts > 0 else 0
            rejected_data = data[data['Status'] == 'rejected']
            # ADD THIS: Add accepted data first
            reasons_data = [{
                'Component': component_name,
                'Reason': 'Accepted',
                'Count': total_accepted,
                'Percentage': round((total_accepted / total_parts * 100) if total_parts > 0 else 0),
                'Rejected_Count': 0
            }]

            if not rejected_data.empty:
                reason_counts = rejected_data['Reason'].value_counts()

                for reason, count in reason_counts.items():
                    if pd.notna(reason) and str(reason).strip() != '' and str(reason).strip().lower() != 'unknown':
                        percentage = (count / total_parts * 100) if total_parts > 0 else 0

                        reasons_data.append({
                            'Component': component_name,
                            'Reason': str(reason),
                            'Count': count,
                            'Percentage': round(percentage),
                            'Rejected_Count': count
                        })

            # Initialize as None - IMPORTANT!
            fig_bar = None
            fig_pie = None
            fig_pie_combined = None
            reasons_df = None
            production_volume_data = {}

            if not rejected_data.empty:
                reason_counts = rejected_data['Reason'].value_counts()
                reasons_data = []

                for reason, count in reason_counts.items():
                    if pd.notna(reason) and str(reason).strip() != '' and str(reason).strip().lower() != 'unknown':
                        # Calculate percentage of TOTAL PARTS
                        percentage = (count / total_parts * 100) if total_parts > 0 else 0
                        reasons_data.append({
                            'Component': component_name,
                            'Reason': str(reason),
                            'Count': count,
                            'Percentage': round(percentage),
                            'Rejected_Count': count
                        })

                if reasons_data:
                    reasons_df = pd.DataFrame(reasons_data)

                    # Get unique reasons and create consistent color mapping
                    unique_reasons = reasons_df['Reason'].unique()
                    color_sequence = create_color_sequence_for_reasons(unique_reasons)
                    fig_bar = px.bar(
                        reasons_df,
                        x='Component',
                        y='Percentage',
                        color='Reason',
                        title=f'{component_name} - Rejection Reasons (% of Total Parts)',
                        labels={'Component': 'Component', 'Percentage': 'Percentage of Total Parts (%)',
                                'Reason': 'Rejection Reason'},
                        color_discrete_sequence=color_sequence,
                        barmode='stack',
                        hover_data={'Count': True},
                        text=[f'{row["Percentage"]}% ({row["Count"]})' for _, row in reasons_df.iterrows()]
                    )

                    # Update hover template
                    fig_bar.update_traces(
                        hovertemplate='<b>%{fullData.name}</b><br>' +
                                      f'{time_grouping}: %{{x}}<br>' +
                                      'Percentage of Total Parts: %{y}%<br>' +
                                      'Rejected Quantity: %{customdata[0]}<extra></extra>'
                    )

                    # Add total rejection annotation
                    total_height = reasons_df['Percentage'].sum()
                    fig_bar.add_annotation(
                        x=component_name,
                        y=total_height + 3,
                        text=f"Total: {round(total_rejection_rate)}% ({total_parts})",
                        showarrow=False,
                        font=dict(size=11, color="black", family="Segoe UI Semilight (Body)"),
                        bgcolor="rgba(255,255,255,0.8)",
                        bordercolor="black",
                        borderwidth=1,
                        xanchor="center"
                    )

                    fig_bar.update_layout(
                        yaxis_range=[0, total_height + 10],
                        height=400,
                        font_family="Segoe UI Semilight (Body)"
                    )

                    # Create pie chart for rejection reasons only
                    fig_pie = px.pie(
                        reasons_df,
                        values='Count',
                        names='Reason',
                        title=f'{component_name} - Rejection Reasons Distribution',
                        color_discrete_sequence=color_sequence
                    )

                    fig_pie.update_layout(
                        font_family="Segoe UI Semilight (Body)",
                        height=400
                    )

                    fig_pie.update_traces(
                        textinfo='label+percent',
                        texttemplate='%{label}<br>%{percent:.0%}',
                        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.0%}<extra></extra>'
                    )

                    # NEW: Create combined pie chart with accepted + rejected breakdown
                    combined_data = {'Accepted': total_accepted}
                    for _, row in reasons_df.iterrows():
                        combined_data[row['Reason']] = row['Count']

                    # Get colors for combined chart
                    combined_reasons = list(combined_data.keys())
                    combined_color_sequence = create_color_sequence_for_reasons(combined_reasons)
                    fig_pie_combined = px.pie(
                        values=list(combined_data.values()),
                        names=list(combined_data.keys()),
                        title=f'{component_name} - Accepted vs Rejected Breakdown',
                        color_discrete_sequence=combined_color_sequence
                    )

                    fig_pie_combined.update_layout(
                        font_family="Segoe UI Semilight (Body)",
                        height=400
                    )

                    fig_pie_combined.update_traces(
                        textinfo='label+percent',
                        texttemplate='%{label}<br>%{percent:.0%}',
                        hovertemplate='<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent:.0%}<extra></extra>'
                    )
                    production_volume_data = {component_name: total_accepted}

            return fig_bar, fig_pie, fig_pie_combined, reasons_df, production_volume_data

        return None, None, None, None, None


    # Create separate plots for each component
    # OPTIC PLOT
    # OPTIC PLOT
    st.write("### Optic Rejection Analysis")
    if "Optic" in filtered_data and not filtered_data["Optic"].empty:
        fig_optic_bar, fig_optic_pie, fig_optic_pie_combined, data_optic, prod_vol_optic = create_rejection_analysis_plot(
            "Optic",
            filtered_data["Optic"],
            time_grouping)

        if fig_optic_bar:
            # Bar chart first
            st.plotly_chart(fig_optic_bar, use_container_width=True)

            # Create two columns for the pie charts
            col1, col2 = st.columns(2)

            with col1:
                # Rejection reasons pie chart
                if fig_optic_pie:
                    st.plotly_chart(fig_optic_pie, use_container_width=True)
                else:
                    st.info("No valid rejection reasons for pie chart")

            with col2:
                # Accepted vs Rejected combined pie chart
                if fig_optic_pie_combined:
                    st.plotly_chart(fig_optic_pie_combined, use_container_width=True)

            # Optic rejection data table
            with st.expander("View Optic Rejection Data"):
                st.dataframe(data_optic, use_container_width=True)
                csv_optic = data_optic.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download Optic Rejection CSV",
                    csv_optic,
                    "optic_rejection_analysis.csv",
                    "text/csv"
                )
        else:
            st.info("No Optic rejection data available - check if rejection reasons exist and are not empty/Unknown")
    else:
        st.info("No Optic data selected or available")

    # FTA PLOT
    st.write("### FTA Rejection Analysis")
    if "FTA" in filtered_data and not filtered_data["FTA"].empty:
        fig_fta_bar, fig_fta_pie, fig_fta_pie_combined, data_fta, prod_vol_fta = create_rejection_analysis_plot("FTA",
                                                                                                                filtered_data[
                                                                                                                    "FTA"],
                                                                                                                time_grouping)

        if fig_fta_bar:
            # Bar chart first
            st.plotly_chart(fig_fta_bar, use_container_width=True)

            # Create two columns for the pie charts
            col1, col2 = st.columns(2)

            with col1:
                # Rejection reasons pie chart
                if fig_fta_pie:
                    st.plotly_chart(fig_fta_pie, use_container_width=True)
                else:
                    st.info("No valid rejection reasons for pie chart")

            with col2:
                # Accepted vs Rejected combined pie chart
                if fig_fta_pie_combined:
                    st.plotly_chart(fig_fta_pie_combined, use_container_width=True)

            # FTA rejection data table
            with st.expander("View FTA Rejection Data"):
                st.dataframe(data_fta, use_container_width=True)
                csv_fta = data_fta.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download FTA Rejection CSV",
                    csv_fta,
                    "fta_rejection_analysis.csv",
                    "text/csv"
                )
        else:
            st.info("No FTA rejection data available")
    else:
        st.info("No FTA data selected or available")

    # PL PLOT
    st.write("### PL Rejection Analysis")
    if "PL" in filtered_data and not filtered_data["PL"].empty:
        fig_pl_bar, fig_pl_pie, fig_pl_pie_combined, data_pl, prod_vol_pl = create_rejection_analysis_plot("PL",
                                                                                                           filtered_data[
                                                                                                               "PL"],
                                                                                                           time_grouping)

        if fig_pl_bar:
            # Bar chart first
            st.plotly_chart(fig_pl_bar, use_container_width=True)

            # Create two columns for the pie charts
            col1, col2 = st.columns(2)

            with col1:
                # Rejection reasons pie chart
                if fig_pl_pie:
                    st.plotly_chart(fig_pl_pie, use_container_width=True)
                else:
                    st.info("No valid rejection reasons for pie chart")

            with col2:
                # Accepted vs Rejected combined pie chart
                if fig_pl_pie_combined:
                    st.plotly_chart(fig_pl_pie_combined, use_container_width=True)

            # PL rejection data table
            with st.expander("View PL Rejection Data"):
                st.dataframe(data_pl, use_container_width=True)
                csv_pl = data_pl.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download PL Rejection CSV",
                    csv_pl,
                    "pl_rejection_analysis.csv",
                    "text/csv"
                )
        else:
            st.info("No PL rejection data available")
    else:
        st.info("No PL data selected or available")

    # WINGS PLOT
    st.write("### Wings Rejection Analysis")
    if "Wings" in filtered_data and not filtered_data["Wings"].empty:
        fig_wings_bar, fig_wings_pie, fig_wings_pie_combined, data_wings, prod_vol_wings = create_rejection_analysis_plot(
            "Wings",
            filtered_data["Wings"],
            time_grouping)

        if fig_wings_bar:
            # Bar chart first
            st.plotly_chart(fig_wings_bar, use_container_width=True)

            # Create two columns for the pie charts
            col1, col2 = st.columns(2)

            with col1:
                # Rejection reasons pie chart
                if fig_wings_pie:
                    st.plotly_chart(fig_wings_pie, use_container_width=True)
                else:
                    st.info("No valid rejection reasons for pie chart")

            with col2:
                # Accepted vs Rejected combined pie chart
                if fig_wings_pie_combined:
                    st.plotly_chart(fig_wings_pie_combined, use_container_width=True)

            # Wings rejection data table
            with st.expander("View Wings Rejection Data"):
                st.dataframe(data_wings, use_container_width=True)
                csv_wings = data_wings.to_csv(index=False).encode('utf-8')
                st.download_button(
                    "📥 Download Wings Rejection CSV",
                    csv_wings,
                    "wings_rejection_analysis.csv",
                    "text/csv"
                )
        else:
            st.info("No Wings rejection data available")
    else:
        st.info("No Wings data selected or available")

    # DATA TABLE
    st.subheader("📋 Data Table")

    # Combine all data for the table
    table_data = []

    for component_name, data in filtered_data.items():
        if not data.empty:
            data_copy = data.copy()
            data_copy['Component'] = component_name
            table_data.append(data_copy)

    if table_data:
        combined_table = pd.concat(table_data, ignore_index=True)

        # Reorder columns for better display
        column_order = ['Date', 'Component', 'Employee', 'Mold', 'Status', 'Reason']
        if 'Configuration' in combined_table.columns:
            column_order.insert(4, 'Configuration')
        if 'File' in combined_table.columns:
            column_order.append('File')

        # Only include columns that exist
        existing_columns = [col for col in column_order if col in combined_table.columns]
        display_table = combined_table[existing_columns]

        st.dataframe(display_table, use_container_width=True)

        # Download option
        csv = display_table.to_csv(index=False).encode('utf-8')
        st.download_button(
            "📥 Download Data as CSV",
            csv,
            "production_data.csv",
            "text/csv"
        )


    else:
        st.info("No data available to display in the table")

else:
    # Instructions for file upload
    st.info("Please upload one or more Excel files with QC sheets.")
    st.markdown("""
    ### Expected file format:
    Each Excel file can contain one or more of these sheets:
    1. **Optic QC** - Individual part inspection data
    2. **FTA QC** - Individual part inspection data with configuration
    3. **PL QC** - Aggregate count data with mold status (r=7.71 format)
    4. **Wings QC** - Aggregate count data (no mold/status)

    **For Optic/FTA sheets:**
    - Date, Employee, Status (accepted/rejected), Mold, Rejection reason
    - Configuration column (FTA only)

    **For PL/Wings sheets:**
    - Date, Employee, Number of approved parts, Number of rejected parts
    - Multiple rejection reason columns with counts
    - Mold status in format like "r=7.71" (PL only)

    You can upload multiple files and they will be combined for analysis.
    """)
