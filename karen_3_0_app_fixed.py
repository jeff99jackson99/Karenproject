#!/usr/bin/env python3
"""
Karen 3.0 NCB Data Processor - FIXED VERSION
NCB Transaction Summarization with specific column requirements
IMPLEMENTING EXACT Instructions 3.0 Requirements
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen 3.0 NCB Data Processor - FIXED", page_icon="📊", layout="wide")

def clean_duplicate_columns(df):
    """Clean duplicate column names by adding suffixes."""
    duplicate_cols = df.columns[df.columns.duplicated()].tolist()
    
    if duplicate_cols:
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

def process_excel_data_karen_3_0_fixed(uploaded_file):
    """Process Excel data from the Data tab for Karen 3.0 - FIXED VERSION."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen 3.0 FIXED Processing Started**")
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            data_sheet_name = 'Data'
            
            if data_sheet_name in excel_data.sheet_names:
                st.write(f"📋 **Processing {data_sheet_name} Sheet**")
                
                # Read the Data sheet - row 12 (index 12) contains the actual headers
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name, header=None)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Row 12 (index 12) contains the actual column headers
                header_row = df.iloc[12]
                st.write(f"🔍 **Header row (Row 12):** {header_row[:20].tolist()}")
                
                # Create DataFrame with proper column names and data
                data_rows = df.iloc[13:].copy()
                new_df = data_rows.copy()
                new_df.columns = header_row
                
                # Clean duplicate column names
                new_df = clean_duplicate_columns(new_df)
                
                st.write(f"📏 Data shape after header fix: {new_df.shape}")
                
                # Find NCB columns by looking for Admin columns
                ncb_columns = find_ncb_columns_karen_3_0_fixed(new_df)
                st.write("🔍 **NCB Columns Found:**")
                if ncb_columns:
                    for ncb_type, col_name in ncb_columns.items():
                        st.write(f"  {ncb_type}: {col_name}")
                else:
                    st.error("❌ Could not find sufficient NCB columns.")
                    return None
                
                # Find required columns by position mapping
                required_cols = find_required_columns_karen_3_0_fixed(new_df)
                st.write("🔍 **Required Columns Found:**")
                for col_letter, col_name in required_cols.items():
                    st.write(f"  {col_letter}: {col_name}")
                
                if len(ncb_columns) >= 7 and len(required_cols) >= 10:
                    # Process the data according to Karen 3.0 FIXED specifications
                    results = process_transaction_data_karen_3_0_fixed(new_df, ncb_columns, required_cols)
                    
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
        st.error(f"❌ Error in Karen 3.0 FIXED processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def find_ncb_columns_karen_3_0_fixed(df):
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
        st.success(f"✅ **Found all 7 NCB columns for Karen 3.0 FIXED**")
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

def find_required_columns_karen_3_0_fixed(df):
    """Find required columns by position mapping for Karen 3.0 FIXED."""
    required_cols = {}
    
    # INSTRUCTIONS 3.0: Use exact column positions from the Data sheet
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
        
        st.success(f"✅ **Found all required columns for Karen 3.0 FIXED**")
        return required_cols
    else:
        st.error(f"❌ **Not enough columns. Need at least 55, found {len(df.columns)}**")
        return None

def process_transaction_data_karen_3_0_fixed(df, ncb_columns, required_cols):
    """Process transaction data according to Karen 3.0 FIXED specifications."""
    try:
        # Get transaction type column (Column J from pull sheet)
        transaction_col = required_cols.get('J')
        if not transaction_col:
            st.error("❌ **Transaction Type column not found!**")
            return None
        
        st.write(f"✅ **Transaction Type column found:** {transaction_col}")
        
        # INSTRUCTIONS 3.0 FIXED: Apply Karen 3.0 filtering rules
        st.write("🎯 **Karen 3.0 FIXED Filtering Rules:**")
        st.write("  - New Business (NB): Admin_Sum > 0 (strictly positive)")
        st.write("  - Reinstatements (R): Admin_Sum > 0 (strictly positive)")
        st.write("  - Cancellations (C): ONLY include records with negative values in Admin 3,4,9,10,6,7,8")
        
        # Apply transaction filtering
        if isinstance(df[transaction_col], pd.DataFrame):
            transaction_series = df[transaction_col].iloc[:, 0]
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
        
        # Convert all Admin columns to numeric
        df_copy = df.copy()
        st.write("🔍 **Converting Admin columns to numeric...**")
        
        for col in ncb_columns.values():
            if col in df_copy.columns:
                df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0)
                st.write(f"  ✅ Converted {col} to numeric")
            else:
                st.error(f"  ❌ Column {col} not found!")
        
        # Calculate Admin_Sum for NB and R filtering
        ncb_cols = list(ncb_columns.values())
        df_copy['Admin_Sum'] = df_copy[ncb_cols].sum(axis=1)
        
        # INSTRUCTIONS 3.0 FIXED: For cancellations, ONLY include records with negative values
        # Check specific Admin columns: 3,4,9,10,6,7,8 for negative values
        admin_cols_for_cancellations = [
            ncb_columns.get('AO'),  # Admin 3 Amount
            ncb_columns.get('AQ'),  # Admin 4 Amount
            ncb_columns.get('BA'),  # Admin 9 Amount
            ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'),  # Admin 6 Amount
            ncb_columns.get('AW'),  # Admin 7 Amount
            ncb_columns.get('AY')   # Admin 8 Amount
        ]
        
        # Filter out None values
        admin_cols_for_cancellations = [col for col in admin_cols_for_cancellations if col is not None]
        
        st.write(f"🔍 **Admin columns checked for cancellation filtering:** {admin_cols_for_cancellations}")
        
        # INSTRUCTIONS 3.0 FIXED: Check if ANY of these Admin columns has negative value
        c_negative_mask = df_copy[admin_cols_for_cancellations] < 0
        c_has_negative = c_negative_mask.any(axis=1)
        
        # Apply strict filtering according to Instructions 3.0
        nb_filtered = nb_df[nb_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
        r_filtered = r_df[r_df.index.isin(df_copy[df_copy['Admin_Sum'] > 0].index)]
        c_filtered = c_df[c_df.index.isin(df_copy[c_has_negative].index)]
        
        st.write(f"🔍 **Karen 3.0 FIXED filtering results:**")
        st.write(f"  New Business (Admin_Sum > 0): {len(nb_filtered)}")
        st.write(f"  Reinstatements (Admin_Sum > 0): {len(r_filtered)}")
        st.write(f"  Cancellations (ONLY negative Admin values): {len(c_filtered)}")
        st.write(f"  Total: {len(nb_filtered) + len(r_filtered) + len(c_filtered)}")
        
        # Check if we have any results
        total_filtered = len(nb_filtered) + len(r_filtered) + len(c_filtered)
        if total_filtered == 0:
            st.error("❌ **No records found after filtering!**")
            return None
        
        # Create output dataframes with required columns in correct order
        # INSTRUCTIONS 3.0 FIXED: Add Admin 6,7,8 after Admin 10 Amount
        
        # Data Set 1: New Business (NB) - 17 columns
        # INSTRUCTIONS: Admin 6,7,8 after Admin 10 Amount
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'),  # Transaction Type (column J) - SINGLE COLUMN
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('BA'), ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY')  # Admin 6,7,8 Amount
        ]
        
        # Data Set 2: Reinstatements (R) - 17 columns
        # INSTRUCTIONS: Admin 6,7,8 after Admin 10 Amount
        # Transaction type in output sheet is in column I but in input sheet it's in column J
        r_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'),  # Transaction Type (column J from input)
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('BA'), ncb_columns.get('BC'),  # Admin 10 Amount
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY')  # Admin 6,7,8 Amount
        ]
        
        # Data Set 3: Cancellations (C) - 21 columns
        # INSTRUCTIONS: Admin 6,7,8 after Admin 10 Amount
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
        
        # Create output dataframes with proper Admin column data
        nb_output = nb_filtered[nb_output_cols].copy()
        r_output = r_filtered[r_output_cols].copy()
        c_output = c_filtered[c_output_cols].copy()
        
        # CRITICAL: Replace Admin columns with the converted numeric data from df_copy
        st.write("🔍 **Ensuring Admin columns contain converted numeric data in output...**")
        
        for col in ['AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC']:
            if col in ncb_columns:
                admin_col_name = ncb_columns[col]
                
                if admin_col_name in df_copy.columns and admin_col_name in nb_output.columns:
                    filtered_indices = nb_filtered.index
                    nb_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].copy()
                    st.write(f"  ✅ Updated {admin_col_name} in NB output")
                
                if admin_col_name in df_copy.columns and admin_col_name in r_output.columns:
                    filtered_indices = r_filtered.index
                    r_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].copy()
                    st.write(f"  ✅ Updated {admin_col_name} in R output")
                
                if admin_col_name in df_copy.columns and admin_col_name in c_output.columns:
                    filtered_indices = c_filtered.index
                    c_output[admin_col_name] = df_copy.loc[filtered_indices, admin_col_name].copy()
                    st.write(f"  ✅ Updated {admin_col_name} in C output")
        
        # Create the results dictionary
        results = {
            'nb_data': nb_output,
            'reinstatement_data': r_output,
            'cancellation_data': c_output,
            'total_records': len(nb_output) + len(r_output) + len(c_output)
        }
        
        st.success(f"✅ **Karen 3.0 FIXED processing completed successfully!**")
        st.write(f"📊 **Final Results:**")
        st.write(f"  - New Business: {len(nb_output)} records")
        st.write(f"  - Reinstatements: {len(r_output)} records")
        st.write(f"  - Cancellations: {len(c_output)} records")
        st.write(f"  - Total: {results['total_records']} records")
        
        return results
        
    except Exception as e:
        st.error(f"❌ Error in Karen 3.0 FIXED processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def main():
    st.title("🔧 Karen 3.0 NCB Data Processor - FIXED VERSION")
    st.write("**IMPLEMENTING EXACT Instructions 3.0 Requirements**")
    
    st.markdown("""
    ### 🎯 **Instructions 3.0 FIXED Requirements:**
    
    **CANCELLATIONS OUTPUT SHEET:**
    - Only include records with negative values in Admin 3,4,9,10,6,7,8
    - Eliminate records with no negative numbers
    
    **NEW BUSINESS OUTPUT SHEET:**
    - Transaction Type (from column J in pull sheet)
    - Add Admin 6,7,8 after Admin 10 Amount
    - Admin 6: AT label, AU amount
    - Admin 7: AV label, AW amount  
    - Admin 8: AX label, AY amount
    
    **REINSTATEMENT OUTPUT SHEET:**
    - Add Admin 6,7,8 after Admin 10 Amount
    - Transaction type from column J in input sheet
    """)
    
    uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx'])
    
    if uploaded_file is not None:
        st.write(f"📁 **File uploaded:** {uploaded_file.name}")
        
        # Process the data
        results = process_excel_data_karen_3_0_fixed(uploaded_file)
        
        if results:
            st.success("🎉 **Processing completed successfully!**")
            
            # Create download buttons
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if len(results['nb_data']) > 0:
                    nb_buffer = io.BytesIO()
                    with pd.ExcelWriter(nb_buffer, engine='openpyxl') as writer:
                        results['nb_data'].to_excel(writer, sheet_name='New Business', index=False)
                    nb_buffer.seek(0)
                    st.download_button(
                        label="📥 Download New Business Data",
                        data=nb_buffer.getvalue(),
                        file_name="Karen_3_0_New_Business.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with col2:
                if len(results['reinstatement_data']) > 0:
                    r_buffer = io.BytesIO()
                    with pd.ExcelWriter(r_buffer, engine='openpyxl') as writer:
                        results['reinstatement_data'].to_excel(writer, sheet_name='Reinstatements', index=False)
                    r_buffer.seek(0)
                    st.download_button(
                        label="📥 Download Reinstatement Data",
                        data=r_buffer.getvalue(),
                        file_name="Karen_3_0_Reinstatements.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            with col3:
                if len(results['cancellation_data']) > 0:
                    c_buffer = io.BytesIO()
                    with pd.ExcelWriter(c_buffer, engine='openpyxl') as writer:
                        results['cancellation_data'].to_excel(writer, sheet_name='Cancellations', index=False)
                    c_buffer.seek(0)
                    st.download_button(
                        label="📥 Download Cancellation Data",
                        data=c_buffer.getvalue(),
                        file_name="Karen_3_0_Cancellations.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            
            # Show summary
            st.write("📊 **Summary:**")
            st.write(f"  - **New Business:** {len(results['nb_data'])} records (Admin_Sum > 0)")
            st.write(f"  - **Reinstatements:** {len(results['reinstatement_data'])} records (Admin_Sum > 0)")
            st.write(f"  - **Cancellations:** {len(results['cancellation_data'])} records (ONLY negative Admin values)")
            st.write(f"  - **Total:** {results['total_records']} records")
            
            # Show sample data
            if len(results['cancellation_data']) > 0:
                st.write("🔍 **Sample Cancellation Data (showing Admin columns):**")
                admin_cols = [col for col in results['cancellation_data'].columns if 'ADMIN' in str(col).upper()]
                if admin_cols:
                    sample_data = results['cancellation_data'][admin_cols].head(10)
                    st.dataframe(sample_data)
                    
                    # Verify negative values
                    st.write("🔍 **Verifying negative values in Admin columns:**")
                    for col in admin_cols:
                        negative_count = (results['cancellation_data'][col] < 0).sum()
                        st.write(f"  {col}: {negative_count} negative values")
        else:
            st.error("❌ **Processing failed. Check the error messages above.**")

if __name__ == "__main__":
    main()
