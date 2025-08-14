#!/usr/bin/env python3
"""
Karen NCB Data Processor - Version 3.0
Working version with Karen 2.0 instruction mapping added
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io
import traceback

st.set_page_config(page_title="Karen NCB v3.0", page_icon="🚀", layout="wide")

def debug_column_info(df, col_name, step_name):
    """Comprehensive column debugging function."""
    st.write(f"🔍 **DEBUG [{step_name}] - Column: {col_name}**")
    
    try:
        # Get column info
        col_data = df[col_name]
        st.write(f"  - Data type: {col_data.dtype}")
        st.write(f"  - Python type: {type(col_data)}")
        st.write(f"  - Column name type: {type(col_name)}")
        st.write(f"  - Column name value: {col_name}")
        
        # Check if it's a datetime
        if col_data.dtype == 'datetime64[ns]':
            st.write(f"  - ⚠️ DATETIME COLUMN DETECTED!")
        elif 'datetime' in str(col_data.dtype):
            st.write(f"  - ⚠️ DATETIME-LIKE COLUMN DETECTED!")
        elif isinstance(col_name, pd.Timestamp):
            st.write(f"  - ⚠️ COLUMN NAME IS DATETIME!")
        elif 'datetime' in str(col_name).lower():
            st.write(f"  - ⚠️ COLUMN NAME CONTAINS 'datetime'!")
        
        # Show sample data
        sample_data = col_data.iloc[1:6] if len(col_data) > 1 else col_data.head()
        st.write(f"  - Sample data (rows 1-5): {sample_data.tolist()}")
        
        # Try numeric conversion
        try:
            numeric_data = pd.to_numeric(col_data.iloc[1:], errors='coerce')
            st.write(f"  - Numeric conversion: {not numeric_data.isna().all()}")
            if not numeric_data.isna().all():
                st.write(f"  - Numeric sample: {numeric_data.dropna().head(3).tolist()}")
        except Exception as e:
            st.write(f"  - Numeric conversion failed: {str(e)}")
            
    except Exception as e:
        st.write(f"  - Error in debug_column_info: {str(e)}")

def analyze_data_structure_debug(df):
    """Debug version of data structure analysis."""
    st.write("🔍 **Starting DEBUG data structure analysis...**")
    
    # Step 1: Find the transaction type column
    transaction_col = None
    st.write("🔄 **Step 1: Finding Transaction Type column...**")
    
    for col in df.columns:
        try:
            # Get sample data and check for transaction types
            sample_data = df[col].dropna().iloc[1:].head(500)
            if len(sample_data) > 0:
                str_vals = [str(val).upper().strip() for val in sample_data]
                nb_count = str_vals.count('NB')
                c_count = str_vals.count('C')
                r_count = str_vals.count('R')
                
                if (nb_count > 0 or c_count > 0 or r_count > 0) and (nb_count + c_count + r_count) > len(sample_data) * 0.1:
                    transaction_col = col
                    st.write(f"✅ **Found Transaction Type column:** {col}")
                    st.write(f"  Sample counts: NB={nb_count}, C={c_count}, R={r_count}")
                    break
        except Exception as e:
            continue
    
    if not transaction_col:
        st.error("❌ **Could not find Transaction Type column with NB, C, R values**")
        return None
    
    # Step 2: Find Admin columns with DEBUG
    st.write("🔄 **Step 2: Finding Admin columns with DEBUG...**")
    admin_cols = []
    
    for col in df.columns:
        try:
            debug_column_info(df, col, "Column Analysis")
            
            # Skip datetime columns
            if df[col].dtype == 'datetime64[ns]' or 'datetime' in str(df[col].dtype):
                st.write(f"  ⚠️ Column {col} rejected: datetime column")
                continue
            
            # Check if column name is datetime
            if isinstance(col, pd.Timestamp) or 'datetime' in str(col).lower():
                st.write(f"  ⚠️ Column {col} rejected: datetime column name")
                continue
            
            # Try to convert to numeric
            col_data = df[col].iloc[1:]  # Skip header
            numeric_data = pd.to_numeric(col_data, errors='coerce')
            
            if not numeric_data.isna().all():
                non_zero_count = (numeric_data != 0).sum()
                total_count = numeric_data.notna().sum()
                
                # More strict criteria: need significant financial data
                if non_zero_count > 10 and total_count > 50:
                    # Check if this looks like financial data (not just row numbers or years)
                    sample_values = numeric_data.dropna().head(10)
                    if len(sample_values) > 0:
                        # Skip columns that are just sequential numbers or years
                        if (sample_values.diff().abs().mean() <= 1.0 or  # Sequential numbers
                            (sample_values >= 1900).all() and (sample_values <= 2030).all()):  # Years
                            st.write(f"  ⚠️ Column {col} rejected: appears to be row numbers or years")
                            continue
                        
                        st.write(f"✅ **Column {col} ACCEPTED as Admin column**")
                        admin_cols.append(col)
                    else:
                        st.write(f"  ⚠️ Column {col} rejected: insufficient data")
                else:
                    st.write(f"  ⚠️ Column {col} rejected: insufficient data")
            else:
                st.write(f"  ⚠️ Column {col} rejected: not numeric")
                
        except Exception as e:
            st.write(f"  ❌ Error analyzing column {col}: {str(e)}")
            continue
    
    if len(admin_cols) < 4:
        st.error(f"❌ **Not enough Admin columns found. Need 4, found {len(admin_cols)}**")
        return None
    
    # Select the first 4 Admin columns
    selected_admin_cols = admin_cols[:4]
    st.write(f"✅ **Selected Admin columns:** {selected_admin_cols}")
    
    return {
        'transaction_col': transaction_col,
        'admin_columns': selected_admin_cols
    }

def process_data_debug(df):
    """Debug version of data processing."""
    st.write("🔍 **Starting DEBUG data processing...**")
    
    try:
        # Analyze data structure using the debug version
        structure_info = analyze_data_structure_debug(df)
        if not structure_info:
            st.error("❌ **Could not analyze data structure**")
            return None
        
        transaction_col = structure_info['transaction_col']
        admin_cols = structure_info['admin_columns']
        
        st.write("✅ **Data structure analysis complete**")
        st.write(f"  Transaction column: {transaction_col}")
        st.write(f"  Admin columns: {admin_cols}")
        
        # Filter by transaction type
        st.write("🔄 **Filtering data by transaction type...**")
        
        # Skip the first row (header) and work with actual data
        data_df = df.iloc[1:].copy()
        
        # Filter by transaction type using the working pattern
        nb_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip().str.contains('NB', na=False)].copy()
        c_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip().str.contains('C', na=False)].copy()
        r_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip().str.contains('R', na=False)].copy()
        
        st.write(f"  - NB records found: {len(nb_df)}")
        st.write(f"  - C records found: {len(c_df)}")
        st.write(f"  - R records found: {len(r_df)}")
        
        # Apply Admin amount filtering with the working logic
        st.write("🔄 **Applying Admin amount filtering...**")
        
        # For NB: sum > 0 AND all individual amounts > 0
        if len(nb_df) > 0:
            admin_cols_list = list(admin_cols)
            
            st.write(f"🔍 **Admin columns for NB filtering:**")
            for i, col in enumerate(admin_cols_list):
                st.write(f"  Admin {i+1}: {col}")
                debug_column_info(df, col, f"Admin {i+1} Analysis")
            
            # Convert to numeric and check for errors
            st.write("🔄 **Converting Admin columns to numeric...**")
            numeric_admin_cols = []
            
            for col in admin_cols_list:
                try:
                    numeric_col = pd.to_numeric(nb_df[col], errors='coerce')
                    if not numeric_col.isna().all():
                        numeric_admin_cols.append(numeric_col)
                        st.write(f"✅ **Successfully converted {col} to numeric**")
                    else:
                        st.write(f"❌ **Failed to convert {col} to numeric**")
                except Exception as e:
                    st.write(f"❌ **Error converting {col}: {str(e)}**")
            
            if len(numeric_admin_cols) < 4:
                st.error(f"❌ **Only {len(numeric_admin_cols)} Admin columns could be converted to numeric**")
                return None
            
            # Calculate sum using the working logic
            st.write("🔄 **Calculating Admin sum...**")
            try:
                nb_df['Admin_Sum'] = pd.concat(numeric_admin_cols, axis=1).sum(axis=1)
                st.write(f"✅ **Admin sum calculated successfully**")
                
                # Show sum statistics
                sum_stats = nb_df['Admin_Sum'].describe()
                st.write(f"  - Admin sum statistics:")
                st.write(f"    - Min: {sum_stats['min']:.2f}")
                st.write(f"    - Max: {sum_stats['max']:.2f}")
                st.write(f"    - Mean: {sum_stats['mean']:.2f}")
                st.write(f"    - Std: {sum_stats['std']:.2f}")
                
            except Exception as e:
                st.write(f"❌ **Error calculating Admin sum: {str(e)}**")
                return None
            
            # First filter: sum > 0
            nb_filtered = nb_df[nb_df['Admin_Sum'] > 0]
            st.write(f"  - NB after Admin sum > 0 filter: {len(nb_filtered)} records")
            
            # Second filter: all individual amounts > 0
            nb_final = nb_filtered[
                (numeric_admin_cols[0] > 0) & 
                (numeric_admin_cols[1] > 0) & 
                (numeric_admin_cols[2] > 0) & 
                (numeric_admin_cols[3] > 0)
            ]
            st.write(f"  - NB after individual Admin > 0 filter: {len(nb_final)} records")
            nb_df = nb_final
        
        # For R: sum > 0
        if len(r_df) > 0:
            admin_cols_list = list(admin_cols)
            
            # Convert to numeric and check for errors
            st.write("🔄 **Converting Admin columns to numeric for R records...**")
            numeric_admin_cols_r = []
            
            for col in admin_cols_list:
                try:
                    numeric_col = pd.to_numeric(r_df[col], errors='coerce')
                    if not numeric_col.isna().all():
                        numeric_admin_cols_r.append(numeric_col)
                        st.write(f"✅ **Successfully converted {col} to numeric for R**")
                    else:
                        st.write(f"❌ **Failed to convert {col} to numeric for R**")
                except Exception as e:
                    st.write(f"❌ **Error converting {col} for R: {str(e)}**")
            
            if len(numeric_admin_cols_r) < 4:
                st.write(f"⚠️ **Only {len(numeric_admin_cols_r)} Admin columns could be converted to numeric for R**")
                # Use the ones we have
                if len(numeric_admin_cols_r) > 0:
                    r_df['Admin_Sum'] = pd.concat(numeric_admin_cols_r, axis=1).sum(axis=1)
                    r_filtered = r_df[r_df['Admin_Sum'] > 0]
                    st.write(f"  - R after Admin sum > 0 filter: {len(r_filtered)} records")
                    r_df = r_filtered
                else:
                    st.write(f"  - R records skipped due to no numeric Admin columns")
            else:
                r_df['Admin_Sum'] = pd.concat(numeric_admin_cols_r, axis=1).sum(axis=1)
                r_filtered = r_df[r_df['Admin_Sum'] > 0]
                st.write(f"  - R after Admin sum > 0 filter: {len(r_filtered)} records")
                r_df = r_filtered
        
        # For C: sum < 0
        if len(c_df) > 0:
            admin_cols_list = list(admin_cols)
            
            # Convert to numeric and check for errors
            st.write("🔄 **Converting Admin columns to numeric for C records...**")
            numeric_admin_cols_c = []
            
            for col in admin_cols_list:
                try:
                    numeric_col = pd.to_numeric(c_df[col], errors='coerce')
                    if not numeric_col.isna().all():
                        numeric_admin_cols_c.append(numeric_col)
                        st.write(f"✅ **Successfully converted {col} to numeric for C**")
                    else:
                        st.write(f"❌ **Failed to convert {col} to numeric for C**")
                except Exception as e:
                    st.write(f"❌ **Error converting {col} for C: {str(e)}**")
            
            if len(numeric_admin_cols_c) < 4:
                st.write(f"⚠️ **Only {len(numeric_admin_cols_c)} Admin columns could be converted to numeric for C**")
                # Use the ones we have
                if len(numeric_admin_cols_c) > 0:
                    c_df['Admin_Sum'] = pd.concat(numeric_admin_cols_c, axis=1).sum(axis=1)
                    c_filtered = c_df[c_df['Admin_Sum'] < 0]
                    st.write(f"  - C after Admin sum < 0 filter: {len(c_filtered)} records")
                    c_df = c_filtered
                else:
                    st.write(f"  - C records skipped due to no numeric Admin columns")
            else:
                c_df['Admin_Sum'] = pd.concat(numeric_admin_cols_c, axis=1).sum(axis=1)
                c_filtered = c_df[c_df['Admin_Sum'] < 0]
                st.write(f"  - C after Admin sum < 0 filter: {len(c_filtered)} records")
                c_df = c_filtered
        
        # Create output dataframes using the working logic
        st.write("🔄 **Creating output dataframes...**")
        
        nb_output = create_output_dataframe_debug(nb_df, 'NB', admin_cols, 'New Business', False)
        c_output = create_output_dataframe_debug(c_df, 'C', admin_cols, 'Cancellation', True)
        r_output = create_output_dataframe_debug(r_df, 'R', admin_cols, 'Reinstatement', False)
        
        st.write("✅ **Data processing complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        return nb_output, c_output, r_output
        
    except Exception as e:
        st.error(f"❌ **Error in data processing:** {str(e)}")
        st.write("**Full error details:**")
        st.code(traceback.format_exc())
        return None

def create_output_dataframe_debug(df, transaction_type, admin_cols, description, is_cancellation):
    """Create output dataframe with debugging."""
    if len(df) == 0:
        return pd.DataFrame()
    
    st.write(f"🔄 **Creating {description} output dataframe...**")
    
    # Create output dataframe with the working logic
    output = pd.DataFrame()
    
    # Add basic columns
    output['Transaction_Type'] = transaction_type
    
    # Add Admin columns
    for i, col in enumerate(admin_cols):
        output[f'Admin_{i+1}_Amount'] = df[col]
    
    # Add other columns if available
    if 'Insurer_Code' in df.columns:
        output['Insurer_Code'] = df['Insurer_Code']
    if 'Product_Type_Code' in df.columns:
        output['Product_Type_Code'] = df['Product_Type_Code']
    if 'Coverage_Code' in df.columns:
        output['Coverage_Code'] = df['Coverage_Code']
    if 'Dealer_Number' in df.columns:
        output['Dealer_Number'] = df['Dealer_Number']
    if 'Dealer_Name' in df.columns:
        output['Dealer_Name'] = df['Dealer_Name']
    if 'Contract_Number' in df.columns:
        output['Contract_Number'] = df['Contract_Number']
    if 'Contract_Sale_Date' in df.columns:
        output['Contract_Sale_Date'] = df['Contract_Sale_Date']
    if 'Transaction_Date' in df.columns:
        output['Transaction_Date'] = df['Transaction_Date']
    if 'Customer_Last_Name' in df.columns:
        output['Customer_Last_Name'] = df['Customer_Last_Name']
    
    # Add cancellation-specific columns for C transactions
    if is_cancellation:
        if 'Contract_Term' in df.columns:
            output['Contract_Term'] = df['Contract_Term']
        if 'Cancellation_Date' in df.columns:
            output['Cancellation_Date'] = df['Cancellation_Date']
        if 'Cancellation_Reason' in df.columns:
            output['Cancellation_Reason'] = df['Cancellation_Reason']
        if 'Cancellation_Factor' in df.columns:
            output['Cancellation_Factor'] = df['Cancellation_Factor']
    
    st.write(f"✅ **{description} output created with {len(output)} columns**")
    return output

def create_excel_download_clean(df, sheet_name):
    """Create Excel download with proper formatting."""
    if len(df) == 0:
        return None
    
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Create a header format
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Apply header format to the first row
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                max_len = max(
                    df[col].astype(str).map(len).max(),
                    len(col)
                )
                worksheet.set_column(i, i, min(max_len + 2, 50))
        
        output.seek(0)
        return output
        
    except Exception as e:
        st.error(f"❌ **Error creating Excel download:** {str(e)}")
        return None

def analyze_data_tab_structure(data_df):
    """Analyze the DATA tab structure to find NCB columns and handle dropdowns."""
    st.write("🔍 **Analyzing DATA tab structure for NCB columns...**")
    
    # Handle potential dropdown issues by examining the first few rows
    st.write("🔍 **Examining first few rows to handle dropdowns and find headers...**")
    
    # Show the first 5 rows to see the structure
    st.write("**First 5 rows of DATA tab:**")
    st.dataframe(data_df.head())
    
    # Look for the actual header row (might not be row 0 due to dropdowns)
    header_row = None
    for row_idx in range(min(5, len(data_df))):
        row_data = data_df.iloc[row_idx]
        # Check if this row contains NCB-related content
        row_str = ' '.join(str(val).upper() for val in row_data if pd.notna(val))
        if any(pattern in row_str for pattern in ['NCB', 'AGENT', 'DEALER', 'ADMIN', 'FEE', 'OFFSET']):
            header_row = row_idx
            st.write(f"✅ **Found header row at index {row_idx}:** Contains NCB-related content")
            break
    
    if header_row is None:
        st.write("⚠️ **No clear header row found with NCB content**")
        st.write("🔍 **Let's examine the data more carefully...**")
        header_row = 0  # Default to first row
    
    # Now search for NCB columns in the DATA tab
    st.write("🔍 **Searching for NCB columns in DATA tab...**")
    
    ncB_columns = []
    
    # Search through all columns for NCB patterns
    for col_idx, col_name in enumerate(data_df.columns):
        try:
            # Get data starting from the header row
            col_data = data_df[col_name].iloc[header_row:].astype(str).str.upper()
            
            # Look for NCB-related content
            ncB_patterns = ['NCB', 'AGENT NCB', 'DEALER NCB', 'ADMIN', 'FEE', 'OFFSET']
            for pattern in ncB_patterns:
                if col_data.str.contains(pattern, na=False).any():
                    matches = col_data.str.contains(pattern, na=False).sum()
                    if matches > 3:  # Need some matches
                        st.write(f"  ✅ **Column {col_idx} ({col_name}):** Contains '{pattern}' ({matches} matches)")
                        
                        # Show sample values from this column
                        sample_values = data_df[col_name].iloc[header_row:].dropna().head(5)
                        st.write(f"    Sample values: {sample_values.tolist()}")
                        
                        # Check if this column contains numeric data (amounts)
                        numeric_data = pd.to_numeric(data_df[col_name].iloc[header_row+1:], errors='coerce')
                        if not numeric_data.isna().all():
                            st.write(f"    ✅ **Contains numeric data:** Min={numeric_data.min():.2f}, Max={numeric_data.max():.2f}")
                        
                        ncB_columns.append({
                            'column_index': col_idx,
                            'column_name': col_name,
                            'pattern_found': pattern,
                            'match_count': matches,
                            'sample_values': sample_values.tolist(),
                            'is_numeric': not numeric_data.isna().all() if 'numeric_data' in locals() else False
                        })
                        break
                        
        except Exception as e:
            continue
    
    if not ncB_columns:
        st.write("⚠️ **No NCB-related columns found in DATA tab**")
        st.write("🔍 **Let's try a broader search...**")
        
        # Try searching for any columns that might be financial
        st.write("🔍 **Searching for potential financial columns...**")
        for col_idx, col_name in enumerate(data_df.columns):
            try:
                # Check if column contains numeric data
                col_data = data_df[col_name].iloc[header_row+1:]  # Skip header
                numeric_data = pd.to_numeric(col_data, errors='coerce')
                
                if not numeric_data.isna().all():
                    non_zero = (numeric_data != 0).sum()
                    total = numeric_data.notna().sum()
                    
                    if non_zero > 10 and total > 50:  # Significant financial data
                        st.write(f"  💰 **Column {col_idx} ({col_name}):** Potential financial column")
                        st.write(f"    - Non-zero: {non_zero}/{total}")
                        st.write(f"    - Range: {numeric_data.min():.2f} to {numeric_data.max():.2f}")
                        
            except Exception as e:
                continue
    
    return ncB_columns, header_row

def main():
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.write("**Expected Output:** 2k-2500 rows in specific order with proper column mapping")
    st.write("**Focus:** DATA tab analysis (per Karen 2.0 instructions)")
    
    # File upload
    uploaded_file = st.file_uploader(
        "📁 Upload Excel File",
        type=['xlsx', 'xls'],
        help="Upload Excel file with NCB data - we'll analyze the DATA tab directly"
    )
    
    if uploaded_file is not None:
        try:
            # Load the file
            df_dict = pd.read_excel(uploaded_file, sheet_name=None)
            
            # Show available sheets
            st.write("🔍 **Available sheets:**", list(df_dict.keys()))
            
            # Focus on the DATA tab as per Karen 2.0 instructions
            data_sheet = None
            for sheet_name in ['Data', 'data', 'DATA']:
                if sheet_name in df_dict:
                    data_sheet = sheet_name
                    break
            
            if not data_sheet:
                st.error("❌ **No 'Data' sheet found**")
                st.write("Karen 2.0 instructions specify using the DATA tab")
                return
            
            st.write(f"✅ **Found DATA sheet:** {data_sheet}")
            
            # Load the DATA sheet
            df = pd.read_excel(uploaded_file, sheet_name=data_sheet)
            st.write(f"📊 **DATA sheet loaded:** {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Analyze the DATA tab structure (handle dropdowns and find NCB columns)
            ncB_columns, header_row = analyze_data_tab_structure(df)
            
            # Show summary of what we found
            st.write("📊 **Summary of DATA Tab Analysis:**")
            st.write(f"  - Header row identified at index: {header_row}")
            st.write(f"  - NCB-related columns found: {len(ncB_columns)}")
            
            if ncB_columns:
                st.write("✅ **NCB columns identified in DATA tab!**")
                st.write("🔄 **Next step:** Use these columns for proper Admin filtering.")
                
                # Show what we found
                st.write("🔍 **NCB Columns found in DATA tab:**")
                for col_info in ncB_columns:
                    st.write(f"  - Column {col_info['column_index']}: {col_info['column_name']}")
                    st.write(f"    - Contains: {col_info['pattern_found']} ({col_info['match_count']} matches)")
                    st.write(f"    - Numeric data: {'Yes' if col_info['is_numeric'] else 'No'}")
                    st.write(f"    - Sample values: {col_info['sample_values']}")
                
                # Now we can build the proper processing logic
                st.write("🎯 **Ready to implement proper NCB processing logic!**")
                
            else:
                st.write("⚠️ **No NCB columns found in DATA tab.**")
                st.write("🔍 **We need to understand your data structure better.**")
                st.write("💡 **Tip:** Check if your Excel file has dropdown arrows that might be hiding data.")
        
        except Exception as e:
            st.error(f"❌ **Error loading file:** {str(e)}")
            st.write("Please ensure your file is a valid Excel file (.xlsx or .xls) with the required sheets.")

if __name__ == "__main__":
    main()
