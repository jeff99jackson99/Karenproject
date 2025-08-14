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
                
                if non_zero_count > 5 and total_count > 10:
                    st.write(f"✅ **Column {col} ACCEPTED as Admin column**")
                    admin_cols.append(col)
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

def main():
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.write("**Expected Output:** 2k-2500 rows in specific order with proper column mapping")
    
    # File upload
    uploaded_file = st.file_uploader(
        "📁 Upload Excel File",
        type=['xlsx', 'xls'],
        help="Upload Excel file with NCB data and 'Col Ref' sheet"
    )
    
    if uploaded_file is not None:
        try:
            # Load the file
            df = pd.read_excel(uploaded_file, sheet_name=None)
            
            # Show available sheets
            st.write("🔍 **Available sheets:**", list(df.keys()))
            
            # Find the data sheet
            data_sheet = None
            for sheet_name in ['Data', 'data', 'DATA']:
                if sheet_name in df:
                    data_sheet = sheet_name
                    break
            
            if not data_sheet:
                st.error("❌ **No 'Data' sheet found**")
                return
            
            # Load the data sheet
            df = pd.read_excel(uploaded_file, sheet_name=data_sheet)
            st.write(f"📊 **Data sheet loaded:** {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Process data button
            if st.button("🔍 **Process Data with Debug Analysis**", type="primary"):
                with st.spinner("🔍 Running debug analysis..."):
                    # Process the data using the debug version
                    result = process_data_debug(df)
                    
                    if result is not None:
                        nb_df, c_df, r_df = result
                        
                        # Display results
                        st.write("📊 **Processing Results**")
                        total_records = len(nb_df) + len(c_df) + len(r_df)
                        st.write(f"**Total Records:** {total_records}")
                        
                        st.write("**Breakdown by Transaction Type:**")
                        st.write(f"- NB (New Business): {len(nb_df)} records")
                        st.write(f"- C (Cancellation): {len(c_df)} records")
                        st.write(f"- R (Reinstatement): {len(r_df)} records")
                        
                        # Show sample data in collapsible sections
                        with st.expander("📋 Sample New Business Data (First 5 rows)"):
                            if len(nb_df) > 0:
                                st.dataframe(nb_df.head())
                            else:
                                st.write("No New Business data found.")
                        
                        with st.expander("📋 Sample Cancellation Data (First 5 rows)"):
                            if len(c_df) > 0:
                                st.dataframe(c_df.head())
                            else:
                                st.write("No Cancellation data found.")
                        
                        with st.expander("📋 Sample Reinstatement Data (First 5 rows)"):
                            if len(r_df) > 0:
                                st.dataframe(r_df.head())
                            else:
                                st.write("No Reinstatement data found.")
                        
                        # Download options
                        st.write("💾 **Download Options**")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if len(nb_df) > 0:
                                nb_download = create_excel_download_clean(nb_df, "New_Business_NB")
                                if nb_download:
                                    st.download_button(
                                        label=f"📥 Download NB Data ({len(nb_df)} rows)",
                                        data=nb_download.getvalue(),
                                        file_name=f"NCB_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                            
                            if len(r_df) > 0:
                                r_download = create_excel_download_clean(r_df, "Reinstatements_R")
                                if r_download:
                                    st.download_button(
                                        label=f"📥 Download R Data ({len(r_df)} rows)",
                                        data=r_download.getvalue(),
                                        file_name=f"NCB_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        with col2:
                            if len(c_df) > 0:
                                c_download = create_excel_download_clean(c_df, "Cancellations_C")
                                if c_download:
                                    st.download_button(
                                        label=f"📥 Download C Data ({len(c_df)} rows)",
                                        data=c_download.getvalue(),
                                        file_name=f"NCB_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        # Combined download
                        st.write("📊 **Combined Download (All Data in One File)**")
                        # This would need to be implemented if needed
        
        except Exception as e:
            st.error(f"❌ **Error loading file:** {str(e)}")
            st.write("Please ensure your file is a valid Excel file (.xlsx or .xls) with the required sheets.")

if __name__ == "__main__":
    main()
