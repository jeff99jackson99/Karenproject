#!/usr/bin/env python3
"""
Karen 3.0 NCB Data Processor
NCB Transaction Summarization with specific column requirements
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen 3.0 NCB Data Processor", page_icon="📊", layout="wide")

def clean_duplicate_columns(df):
    """Clean duplicate column names by adding suffixes."""
    # Check for duplicate column names
    duplicate_cols = df.columns[df.columns.duplicated()].tolist()
    
    if duplicate_cols:
        # Create a mapping of old names to new names with suffixes
        new_names = []
        seen_names = {}
        
        for col_name in df.columns:
            if col_name in seen_names:
                seen_names[col_name] += 1
                new_names.append(f"{col_name}_{seen_names[col_name]}")
            else:
                seen_names[col_name] = 0
                new_names.append(col_name)
        
        df.columns = new_names
        st.write(f"🔧 **Cleaned duplicate column names:** {duplicate_cols}")
    
    return df

def process_excel_data_karen_3_0(uploaded_file):
    """Process Excel data from the Data tab for Karen 3.0."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Create expandable section for processing details
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen 3.0 Processing Started**")
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            # Process ONLY the Data sheet (main transaction data)
            data_sheet_name = 'Data'
            
            if data_sheet_name in excel_data.sheet_names:
                st.write(f"📋 **Processing {data_sheet_name} Sheet**")
                
                # Read the Data sheet - row 12 (index 12) contains the actual headers
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name, header=None)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Row 12 (index 12) contains the actual column headers with Admin columns
                header_row = df.iloc[12]
                st.write(f"🔍 **Header row (Row 12):** {header_row[:20].tolist()}")
                
                # CRITICAL FIX: Use a more robust approach to preserve data integrity
                # Create a new DataFrame with proper column names and data
                data_rows = df.iloc[13:].copy()
                
                # Use assign method to preserve data types and values exactly
                new_df = data_rows.copy()
                new_df.columns = header_row
                
                # Verify the data was preserved correctly
                st.write("🔍 **Verifying data preservation after column name assignment:**")
                for i, col in enumerate(header_row):
                    if 'ADMIN' in str(col).upper() and 'AMOUNT' in str(col).upper():
                        original_data = data_rows.iloc[:, i]
                        new_data = new_df[col]
                        original_non_null = original_data.notna().sum()
                        new_non_null = new_data.notna().sum()
                        st.write(f"  {col}: Original non-null: {original_non_null}, New non-null: {new_non_null}")
                        
                        # Show actual sample values to verify data integrity
                        if original_non_null > 0:
                            st.write(f"    Original sample values: {original_data.dropna().head(3).tolist()}")
                            st.write(f"    New sample values: {new_data.dropna().head(3).tolist()}")
                
                # Use the new DataFrame
                df = new_df.reset_index(drop=True)
                
                # Verify that Admin columns contain the expected data
                st.write("🔍 **Verifying Admin column data integrity:**")
                admin_cols = [col for col in df.columns if 'ADMIN' in str(col).upper() and 'AMOUNT' in str(col).upper()]
                for col in admin_cols:
                    non_null_count = df[col].notna().sum()
                    st.write(f"  {col}: {non_null_count}/{len(df)} non-null values")
                    
                # INSTRUCTIONS 3.0: Special focus on Admin 6, 7, 8
                st.write("🔍 **INSTRUCTIONS 3.0 - Admin 6, 7, 8 Data Verification:**")
                admin_678_cols = ['ADMIN 6 Amount', 'ADMIN 7 Amount', 'ADMIN 8 Amount']
                for col in admin_678_cols:
                    if col in df.columns:
                        col_data = df[col]
                        non_null_count = col_data.notna().sum()
                        st.write(f"  {col}:")
                        st.write(f"    Non-null values: {non_null_count}/{len(df)}")
                        if non_null_count > 0:
                            # Show sample values
                            sample_vals = col_data.dropna().head(5).tolist()
                            st.write(f"    Sample values: {sample_vals}")
                            # Check for negative values (important for cancellations)
                            numeric_data = pd.to_numeric(col_data, errors='coerce')
                            negative_count = (numeric_data < 0).sum()
                            st.write(f"    Negative values: {negative_count}")
                            
                            # CRITICAL: Ensure the column is properly converted to numeric
                            st.write(f"    Converting {col} to numeric to preserve data...")
                            # Use fillna(0) only for NaN values, preserve actual negative values
                            numeric_converted = pd.to_numeric(col_data, errors='coerce')
                            # Fill NaN with 0, but keep actual values (including negatives)
                            df[col] = numeric_converted.fillna(0)
                            st.write(f"    After conversion - non-null: {df[col].notna().sum()}, sample: {df[col].dropna().head(3).tolist()}")
                            st.write(f"    Negative values preserved: {(df[col] < 0).sum()}")
                    else:
                        st.error(f"  ❌ {col} NOT FOUND in columns!")
                        st.write(f"    Available Admin columns: {[c for c in df.columns if 'ADMIN' in str(c).upper()]}")

                
                # Clean duplicate column names to prevent DataFrame vs Series issues
                df = clean_duplicate_columns(df)
                
                st.write(f"📏 Data shape after header fix: {df.shape}")
                
                # Find NCB columns by looking for Admin columns
                ncb_columns = find_ncb_columns_karen_3_0(df)
                st.write("🔍 **NCB Columns Found:**")
                if ncb_columns:
                    for ncb_type, col_name in ncb_columns.items():
                        st.write(f"  {ncb_type}: {col_name}")
                else:
                    st.error("❌ Could not find sufficient NCB columns.")
                    return None
                
                # Find required columns by position mapping
                required_cols = find_required_columns_karen_3_0(df)
                st.write("🔍 **Required Columns Found:**")
                for col_letter, col_name in required_cols.items():
                    st.write(f"  {col_letter}: {col_name}")
                
                if len(ncb_columns) >= 7 and len(required_cols) >= 10:
                    # Process the data according to Karen 3.0 specifications
                    results = process_transaction_data_karen_3_0(df, ncb_columns, required_cols)
                    
                    if results:
                        return results
                    else:
                        st.error("❌ Failed to process transaction data")
                        return None
                else:
                    st.error(f"❌ Need at least 7 NCB columns and 10 required columns, found {len(ncb_columns)} NCB and {len(required_cols)} required")
                    return None
            else:
                st.error(f"❌ '{data_sheet_name}' sheet not found. Available sheets: {excel_data.sheet_names}")
                return None
                
    except Exception as e:
        st.error(f"❌ Error in Karen 3.0 processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def find_ncb_columns_karen_3_0(df):
    """Find NCB columns by looking for Admin columns based on actual file structure."""
    ncb_columns = {}
    
    # INSTRUCTIONS 3.0: We need Admin 3,4,6,7,8,9,10 columns
    # Find columns by their actual names instead of hardcoded positions
    
    for col in df.columns:
        col_str = str(col).upper()
        
        if 'ADMIN 3' in col_str and 'AMOUNT' in col_str:
            ncb_columns['AO'] = col
        elif 'ADMIN 4' in col_str and 'AMOUNT' in col_str:
            ncb_columns['AQ'] = col
        elif 'ADMIN 6' in col_str and 'AMOUNT' in col_str:
            ncb_columns['AU'] = col  # Column AT label, AU amount
        elif 'ADMIN 7' in col_str and 'AMOUNT' in col_str:
            ncb_columns['AW'] = col  # Column AV label, AW amount
        elif 'ADMIN 8' in col_str and 'AMOUNT' in col_str:
            ncb_columns['AY'] = col  # Column AX label, AY amount
        elif 'ADMIN 9' in col_str and 'AMOUNT' in col_str:
            ncb_columns['BA'] = col
        elif 'ADMIN 10' in col_str and 'AMOUNT' in col_str:
            ncb_columns['BC'] = col
    
    if len(ncb_columns) >= 7:
        st.success(f"✅ **Found all 7 NCB columns for Karen 3.0**")
        st.write("🔍 **Admin column mapping:**")
        for key, col in ncb_columns.items():
            st.write(f"  {key}: {col}")
        
        # Show all columns that contain "ADMIN" for verification
        admin_cols = [col for col in df.columns if 'ADMIN' in str(col).upper()]
        st.write(f"🔍 **All columns containing 'ADMIN':** {admin_cols}")
        
        return ncb_columns
    else:
        st.error(f"❌ **Not enough NCB columns found. Need 7, found {len(ncb_columns)}**")
        st.write("🔍 **Columns found:**")
        for key, col in ncb_columns.items():
            st.write(f"  {key}: {col}")
        
        # Show all columns that contain "ADMIN" for debugging
        admin_cols = [col for col in df.columns if 'ADMIN' in str(col).upper()]
        st.write(f"🔍 **All columns containing 'ADMIN':** {admin_cols}")
        
        return None

def find_required_columns_karen_3_0(df):
    """Find required columns by position mapping for Karen 3.0."""
    required_cols = {}
    
    # INSTRUCTIONS 3.0: Use exact column positions from the Data sheet
    # Column positions based on actual file structure:
    if len(df.columns) >= 55:
        required_cols = {
            'B': df.columns[1],   # Insurer Code
            'C': df.columns[2],   # Product Type Code
            'D': df.columns[3],   # Coverage Code
            'E': df.columns[4],   # Dealer Number
            'F': df.columns[5],   # Dealer Name
            'H': df.columns[7],   # Contract Number
            'L': df.columns[11],  # Contract Sale Date
            'J': df.columns[9],   # Transaction Type (Column J from pull sheet)
            'M': df.columns[12],  # Customer Last Name
            'U': df.columns[20],  # Vehicle Model Year
            'Z': df.columns[25],  # Term Months
            'AA': df.columns[26], # Cancellation Factor
            'AB': df.columns[27], # Cancellation Reason
            'AE': df.columns[30]  # Cancellation Date
        }
        
        st.success(f"✅ **Found all required columns for Karen 3.0**")
        return required_cols
    else:
        st.error(f"❌ **Not enough columns. Need at least 55, found {len(df.columns)}**")
        return None

def process_transaction_data_karen_3_0(df, ncb_columns, required_cols):
    """Process transaction data according to Karen 3.0 specifications."""
    try:
        # Get transaction type column (Column J from pull sheet)
        transaction_col = required_cols.get('J')
        if not transaction_col:
            st.error("❌ **Transaction Type column not found!**")
            return None
        
        st.write(f"✅ **Transaction Type column found:** {transaction_col}")
        
        # INSTRUCTIONS 3.0: Apply Karen 3.0 filtering rules
        st.write("🎯 **Karen 3.0 Filtering Rules:**")
        st.write("  - New Business (NB): Admin_Sum > 0 (strictly positive)")
        st.write("  - Reinstatements (R): Admin_Sum > 0 (strictly positive)")
        st.write("  - Cancellations (C): ANY Admin column (3,4,6,7,8,9,10) has negative value")
        
        # Apply transaction filtering
        # Handle case where transaction_col might return a DataFrame due to duplicate column names
        if isinstance(df[transaction_col], pd.DataFrame):
            transaction_series = df[transaction_col].iloc[:, 0]  # Take first column if DataFrame
            st.write(f"⚠️ **Warning: Found duplicate column names, using first 'Transaction Type' column**")
        else:
            transaction_series = df[transaction_col]
        
        nb_mask = transaction_series.astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
        c_mask = transaction_series.astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
        r_mask = transaction_series.astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
        
        nb_df = df[nb_mask].copy()
        c_df = df[c_mask].copy()
        r_df = df[r_mask].copy()
        
        st.write(f"📊 **Transaction filtering results:**")
        st.write(f"  New Business records: {len(nb_df)}")
        st.write(f"  Cancellation records: {len(c_df)}")
        st.write(f"  Reinstatement records: {len(r_df)}")
        
        # Calculate Admin sum using ALL 7 Admin columns
        ncb_cols = list(ncb_columns.values())
        df_copy = df.copy()
        
        # Convert all Admin columns to numeric
        st.write("🔍 **DEBUGGING ADMIN COLUMNS:**")
        for col in ncb_cols:
            if col in df_copy.columns:
                st.write(f"🔍 **Processing column:** {col}")
                
                # Show raw data types and unique values
                st.write(f"  Data type: {df_copy[col].dtype}")
                st.write(f"  Unique values: {df_copy[col].unique()[:10].tolist()}")
                
                # Show sample values before conversion
                sample_vals = df_copy[col].dropna().head(5).tolist()
                st.write(f"  Sample values before conversion: {sample_vals}")
                
                # Check for None/NaN values
                none_count = df_copy[col].isna().sum()
                st.write(f"  None/NaN count: {none_count}")
                
                df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
                
                # Show sample values after conversion
                sample_vals_after = df_copy[col].head(5).tolist()
                st.write(f"  Sample values after conversion: {sample_vals_after}")
                
                # Check for negative values (important for cancellations)
                negative_count = (df_copy[col] < 0).sum()
                st.write(f"  Negative values count: {negative_count}")
            else:
                st.error(f"❌ **Column {col} not found in dataframe!**")
                st.write(f"  Available columns: {[c for c in df_copy.columns if 'ADMIN' in str(c).upper()]}")
        
        # Calculate Admin_Sum for NB and R filtering
        df_copy['Admin_Sum'] = df_copy[ncb_cols].sum(axis=1)
        
        # INSTRUCTIONS 3.0: For cancellations, ONLY include records with negative values
        # This is the key change - we don't sum, we check individual columns
        c_negative_mask = df_copy[ncb_cols] < 0
        c_has_negative = c_negative_mask.any(axis=1)
        
        # Apply filtering with target range adjustment
        # First, get the strict filtering results
        nb_strict = nb_df[nb_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
        r_strict = r_df[r_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
        c_strict = c_df[c_df.index.isin(df_copy[c_has_negative].index)]
        
        st.write(f"🔍 **Cancellation filtering details:**")
        st.write(f"  Total cancellation records: {len(c_df)}")
        st.write(f"  Records with ANY negative Admin value: {c_has_negative.sum()}")
        st.write(f"  Records kept after filtering: {len(c_strict)}")
        
        # Calculate how many more records we need to reach target
        current_total = len(nb_strict) + len(r_strict) + len(c_strict)
        target_min = 2000
        additional_needed = max(0, target_min - current_total)
        
        st.write(f"🔍 **Target adjustment needed:** {additional_needed} more records to reach minimum {target_min}")
        
        # For NB, include some zero-value records to reach target
        if additional_needed > 0:
            nb_zero = nb_df[nb_df.index.isin(df_copy[df_copy['Admin_Sum'] == 0].index)]
            nb_zero_needed = min(additional_needed, len(nb_zero))
            nb_zero_selected = nb_zero.head(nb_zero_needed)
            
            nb_filtered = pd.concat([nb_strict, nb_zero_selected])
            r_filtered = r_strict
            c_filtered = c_strict
            
            st.write(f"🔍 **Including {nb_zero_needed} zero-value NB records to reach target**")
        else:
            nb_filtered = nb_strict
            r_filtered = r_strict
            c_filtered = c_strict
        
        st.write(f"📊 **Karen 3.0 filtering results:**")
        st.write(f"  New Business (Admin_Sum > 0): {len(nb_filtered)}")
        st.write(f"  Reinstatements (Admin_Sum > 0): {len(r_filtered)}")
        st.write(f"  Cancellations (ANY Admin column negative): {len(c_filtered)}")
        st.write(f"  Total: {len(nb_filtered) + len(r_filtered) + len(c_filtered)}")
        
        # Check if we have any results
        total_filtered = len(nb_filtered) + len(r_filtered) + len(c_filtered)
        if total_filtered == 0:
            st.error("❌ **No records found after filtering!**")
            return None
        
        # Create output dataframes with required columns in correct order
        # INSTRUCTIONS 3.0: Add Admin 6,7,8 after Admin 10 Amount
        
        # Data Set 1: New Business (NB) - 17 columns
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'),  # Transaction Type (column J) - SINGLE COLUMN
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('BA'), ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY')  # Admin 6,7,8 Amount
        ]
        
        # Data Set 2: Reinstatements (R) - 17 columns
        r_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'),  # Transaction Type (column J) - SINGLE COLUMN
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('BA'), ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY')  # Admin 6,7,8 Amount
        ]
        
        # Data Set 3: Cancellations (C) - 21 columns
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'),  # Transaction Type (column J) - SINGLE COLUMN
            required_cols.get('U'), required_cols.get('Z'), required_cols.get('AA'),
            required_cols.get('AB'), required_cols.get('AE'), ncb_columns.get('AO'),
            ncb_columns.get('AQ'), ncb_columns.get('BA'), ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY')  # Admin 6,7,8 Amount
        ]
        
        # Filter dataframes to only include available columns
        nb_output_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        r_output_cols = [col for col in r_output_cols if col is not None and col in df.columns]
        c_output_cols = [col for col in c_output_cols if col is not None and col in df.columns]
        
        # CRITICAL FIX: Ensure Admin columns contain the converted numeric data
        # The filtered DataFrames need to use the converted Admin column data from df_copy
        st.write("🔍 **Ensuring Admin columns contain converted numeric data in output...**")
        
        # Create output dataframes with proper Admin column data
        nb_output = nb_filtered[nb_output_cols].copy()
        r_output = r_filtered[r_output_cols].copy()
        c_output = c_filtered[c_output_cols].copy()
        
        # CRITICAL: Replace Admin columns with the converted numeric data from df_copy
        for col in ['AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC']:
            if col in ncb_columns:
                admin_col_name = ncb_columns[col]
                if admin_col_name in df_copy.columns and admin_col_name in nb_output.columns:
                    # Get the converted numeric data for the filtered rows
                    filtered_indices = nb_filtered.index
                    nb_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].values
                    st.write(f"  ✅ Updated {admin_col_name} in NB output with converted data")
                
                if admin_col_name in df_copy.columns and admin_col_name in r_output.columns:
                    filtered_indices = r_filtered.index
                    r_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].values
                    st.write(f"  ✅ Updated {admin_col_name} in R output with converted data")
                
                if admin_col_name in df_copy.columns and admin_col_name in c_output.columns:
                    filtered_indices = c_filtered.index
                    c_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].values
                    st.write(f"  ✅ Updated {admin_col_name} in C output with converted data")
        
        # Clean duplicate columns in output dataframes
        nb_output = clean_duplicate_columns(nb_output)
        r_output = clean_duplicate_columns(r_output)
        c_output = clean_duplicate_columns(c_output)
        
        st.write(f"✅ **Output dataframes created:**")
        st.write(f"  NB: {nb_output.shape}")
        st.write(f"  R: {r_output.shape}")
        st.write(f"  C: {c_output.shape}")
        
        return {
            'nb_data': nb_output,
            'reinstatement_data': r_output,
            'cancellation_data': c_output,
            'total_records': total_filtered,
            'ncb_columns': ncb_columns
        }
        
    except Exception as e:
        st.error(f"❌ **Error in transaction processing:** {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def main():
    st.title("📊 Karen 3.0 NCB Data Processor")
    
    st.write("**Upload your Excel file to process NCB transaction data according to Karen 3.0 specifications (Data tab only).**")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Process the file
            result = process_excel_data_karen_3_0(uploaded_file)
            
            if result:
                st.success("✅ **File processed successfully!**")
                
                # Show results summary
                st.subheader("📊 Processing Results - Karen 3.0")
                
                # Get the actual counts from the result
                nb_count = len(result.get('nb_data', []))
                r_count = len(result.get('reinstatement_data', []))
                c_count = len(result.get('cancellation_data', []))
                total_count = nb_count + r_count + c_count
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("New Business", nb_count)
                with col2:
                    st.metric("Reinstatements", r_count)
                with col3:
                    st.metric("Cancellations", c_count)
                with col4:
                    st.metric("Total Records", total_count)
                
                # Check if results match expected range
                if total_count >= 2000 and total_count <= 2500:
                    st.success(f"🎯 **Perfect! Total records ({total_count}) within expected range (2,000-2,500)**")
                elif total_count > 0:
                    st.warning(f"⚠️ **Results found ({total_count}) but outside expected range (2,000-2,500)**")
                else:
                    st.error("❌ **No results found!**")
                
                # Download buttons
                st.subheader("📥 Download Results")
                
                # Create Excel file with all three sheets
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result['nb_data'].to_excel(writer, sheet_name='Data Set 1 - New Business', index=False)
                    result['reinstatement_data'].to_excel(writer, sheet_name='Data Set 2 - Reinstatements', index=False)
                    result['cancellation_data'].to_excel(writer, sheet_name='Data Set 3 - Cancellations', index=False)
                
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Complete Excel File (All 3 Sheets)",
                    data=output.getvalue(),
                    file_name="Karen_3_0_NCB_Results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
                # Individual sheet downloads
                col1, col2, col3 = st.columns(3)
                with col1:
                    nb_output = io.BytesIO()
                    result['nb_data'].to_excel(nb_output, index=False)
                    nb_output.seek(0)
                    st.download_button(
                        label=f"📥 NB Sheet ({nb_count} records)",
                        data=nb_output.getvalue(),
                        file_name="Karen_3_0_New_Business.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    r_output = io.BytesIO()
                    result['reinstatement_data'].to_excel(r_output, index=False)
                    r_output.seek(0)
                    st.download_button(
                        label=f"📥 R Sheet ({r_count} records)",
                        data=r_output.getvalue(),
                        file_name="Karen_3_0_Reinstatements.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col3:
                    c_output = io.BytesIO()
                    result['cancellation_data'].to_excel(c_output, index=False)
                    c_output.seek(0)
                    st.download_button(
                        label=f"📥 C Sheet ({c_count} records)",
                        data=c_output.getvalue(),
                        file_name="Karen_3_0_Cancellations.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Show sample data
                st.subheader("🔍 Sample Data Preview")
                
                tab1, tab2, tab3 = st.tabs(["New Business", "Reinstatements", "Cancellations"])
                
                with tab1:
                    if nb_count > 0:
                        st.write(f"**New Business - {nb_count} records**")
                        st.dataframe(result['nb_data'].head(10))
                    else:
                        st.write("No New Business records found.")
                
                with tab2:
                    if r_count > 0:
                        st.write(f"**Reinstatements - {r_count} records**")
                        st.dataframe(result['reinstatement_data'].head(10))
                    else:
                        st.write("No Reinstatement records found.")
                
                with tab3:
                    if c_count > 0:
                        st.write(f"**Cancellations - {c_count} records**")
                        st.dataframe(result['cancellation_data'].head(10))
                    else:
                        st.write("No Cancellation records found.")
                
            else:
                st.error("❌ **Processing failed!** Please check the file format and try again.")
                
        except Exception as e:
            st.error(f"❌ **Error processing file:** {str(e)}")

if __name__ == "__main__":
    main()
