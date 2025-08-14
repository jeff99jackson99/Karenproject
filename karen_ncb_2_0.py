#!/usr/bin/env python3
"""
Karen NCB Data Processor - Iteration 2.0
Expected Output: 2k-2500 rows in specific order with proper column mapping
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen NCB 2.0", page_icon="🚀", layout="wide")

def detect_column_structure_v2(excel_file):
    """Detect column structure from Col Ref sheet for version 2.0 - using working logic from first version."""
    try:
        # First, let's see what sheets are available
        excel_data = pd.ExcelFile(excel_file)
        st.write(f"🔍 **Available sheets:** {excel_data.sheet_names}")
        
        # Look for the Col Ref sheet (try different possible names)
        col_ref_sheet = None
        possible_names = ['Col Ref', 'ColRef', 'Column Reference', 'ColumnRef', 'Reference', 'Col_Ref', 'xref', 'Xref']
        
        for sheet_name in excel_data.sheet_names:
            if any(name.lower() in sheet_name.lower() for name in possible_names):
                col_ref_sheet = sheet_name
                break
        
        if not col_ref_sheet:
            st.error(f"❌ Could not find Col Ref sheet. Available sheets: {excel_data.sheet_names}")
            return {}, {}, {}
        
        st.write(f"✅ **Using reference sheet:** {col_ref_sheet}")
        
        # Read the reference sheet
        col_ref_df = pd.read_excel(excel_file, sheet_name=col_ref_sheet)
        
        st.write(f"🔍 **Reference sheet loaded:** {col_ref_df.shape}")
        st.write(f"🔍 **Reference sheet first few rows:**")
        st.dataframe(col_ref_df.head(10))
        
        # Use the EXACT working logic from the first version
        column_mapping = {}
        admin_columns = {}
        
        # Look for the row that contains column descriptions
        for idx, row in col_ref_df.iterrows():
            if 'Admin' in str(row.iloc[2]) or 'NCB' in str(row.iloc[7]):
                # This is the row with column descriptions
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and str(value).strip():
                        desc = str(value).strip()
                        column_mapping[col_idx] = desc
                        
                        # Use the EXACT working logic from first version
                        if 'Admin' in desc or 'NCB' in desc:
                            if 'Amount' in desc or 'Fee' in desc:
                                # This is an Admin amount column
                                if 'Agent' in desc:
                                    admin_columns['Admin 3'] = col_idx
                                elif 'Dealer' in desc:
                                    admin_columns['Admin 4'] = col_idx
                                elif 'Offset' in desc:
                                    if 'Agent' in desc:
                                        admin_columns['Admin 9'] = col_idx
                                    elif 'Dealer' in desc:
                                        admin_columns['Admin 10'] = col_idx
                break
        
        # If we don't have enough admin columns, try to find them by content analysis (working logic from first version)
        if len(admin_columns) < 4:
            st.warning("⚠️ Could not identify all Admin columns from mapping, trying content-based detection...")
            
            # Look for columns that might contain Admin amounts by analyzing their content
            potential_admin_cols = []
            
            # Read the main data sheet to analyze column content
            data_sheet = None
            for sheet in excel_data.sheet_names:
                if sheet != col_ref_sheet and not any(keyword in sheet.lower() for keyword in ['summary', 'xref', 'ref', 'col']):
                    data_sheet = sheet
                    break
            
            if data_sheet:
                df = pd.read_excel(excel_file, sheet_name=data_sheet)
                
                for col in df.columns:
                    try:
                        # Skip datetime columns
                        if isinstance(col, pd.Timestamp) or 'datetime' in str(type(col)).lower():
                            continue
                            
                        # Try to convert to numeric
                        numeric_values = pd.to_numeric(df[col], errors='coerce')
                        if not numeric_values.isna().all() and numeric_values.dtype in ['int64', 'float64']:
                            # Check if this column has meaningful non-zero values
                            non_zero_count = (numeric_values > 0).sum()
                            total_count = len(numeric_values)
                            
                            # Only consider columns with a reasonable number of non-zero values
                            if non_zero_count > 0 and non_zero_count < total_count * 0.9:  # Not all zeros, not all non-zero
                                potential_admin_cols.append({
                                    'column': col,
                                    'non_zero_count': non_zero_count,
                                    'total_count': total_count,
                                    'non_zero_ratio': non_zero_count / total_count
                                })
                    except:
                        pass
                
                # Sort by non-zero ratio to find the most likely Admin columns
                potential_admin_cols.sort(key=lambda x: x['non_zero_ratio'], reverse=True)
                
                st.write(f"🔍 **Potential Admin columns found:** {len(potential_admin_cols)}")
                for i, col_info in enumerate(potential_admin_cols[:10]):  # Show top 10
                    st.write(f"  {i+1}. {col_info['column']}: {col_info['non_zero_count']}/{col_info['total_count']} non-zero ({col_info['non_zero_ratio']:.1%})")
                
                # Select the top 4 columns as Admin columns
                if len(potential_admin_cols) >= 4:
                    admin_columns = {
                        'Admin 3': potential_admin_cols[0]['column'],
                        'Admin 4': potential_admin_cols[1]['column'], 
                        'Admin 9': potential_admin_cols[2]['column'],
                        'Admin 10': potential_admin_cols[3]['column']
                    }
                    
                    st.write(f"✅ **Admin columns assigned by content analysis:**")
                    for admin_type, col_name in admin_columns.items():
                        st.write(f"  {admin_type}: {col_name}")
                else:
                    st.error(f"❌ Not enough potential Admin columns found. Need 4, found {len(potential_admin_cols)}")
        
        if admin_columns:
            st.write("🔍 **Column Structure Detected:**")
            st.write(f"  Total columns mapped: {len(column_mapping)}")
            st.write(f"  Admin columns found: {len(admin_columns)}")
            
            st.write("🔍 **Admin column mappings:**")
            for admin_type, col_idx in admin_columns.items():
                st.write(f"  {admin_type}: Column {col_idx}")
                if col_idx in column_mapping:
                    st.write(f"    Description: {column_mapping[col_idx]}")
            
            return column_mapping, admin_columns, {}
        else:
            st.warning("⚠️ No Admin column mapping found in reference sheet")
            return {}, {}, {}
        
    except Exception as e:
        st.error(f"❌ Error reading reference sheet: {e}")
        import traceback
        st.code(traceback.format_exc())
        return {}, {}, {}

def find_transaction_column(df):
    """Find the transaction type column."""
    transaction_col = None
    
    # Look for transaction type values in columns
    for col in df.columns:
        if df[col].dtype == 'object':
            sample_vals = df[col].dropna().head(10).astype(str).str.upper()
            if any(val in ['NB', 'C', 'R', 'NEW BUSINESS', 'CANCELLATION', 'REINSTATEMENT'] for val in sample_vals):
                transaction_col = col
                break
    
    # Fallback to name-based search
    if not transaction_col:
        for col in df.columns:
            if any(keyword in str(col).upper() for keyword in ['TRANSACTION', 'TYPE', 'TRANS']):
                transaction_col = col
                break
    
    return transaction_col

def process_data_v2(df, column_mapping, admin_columns, amount_columns):
    """Process data according to version 2.0 requirements - using label matching and exact filtering logic."""
    
    # Find transaction column
    transaction_col = find_transaction_column(df)
    if not transaction_col:
        st.error("❌ Could not find transaction type column")
        return None
    
    st.write(f"✅ **Transaction column found:** {transaction_col}")
    
    # Separate data by transaction type
    nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
    c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
    r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
    
    st.write(f"📊 **Data breakdown:**")
    st.write(f"  New Business: {nb_mask.sum()}")
    st.write(f"  Cancellations: {c_mask.sum()}")
    st.write(f"  Reinstatements: {r_mask.sum()}")
    
    # Find NCB-related columns using label matching (not fixed positions)
    ncbi_labels = [
        'Agent NCB', 'Agent NCB Fee', 'Dealer NCB', 'Dealer NCB Fee',
        'Agent NCB Offset', 'Agent NCB Offset Bucket', 'Dealer NCB Offset', 'Dealer NCB Offset Bucket'
    ]
    
    # Find label columns and their corresponding amount columns
    label_amount_pairs = []
    for i, col in enumerate(df.columns):
        col_str = str(col).strip()
        for label in ncbi_labels:
            if label.lower() in col_str.lower():
                # Found a label column, get the next column for the amount
                if i + 1 < len(df.columns):
                    amount_col = df.columns[i + 1]
                    label_amount_pairs.append({
                        'label_col': col,
                        'amount_col': amount_col,
                        'label_name': label
                    })
                    st.write(f"✅ **Found NCB pair:** {col} → {amount_col} ({label})")
                break
    
    if len(label_amount_pairs) < 4:
        st.error(f"❌ Need at least 4 NCB label-amount pairs, found {len(label_amount_pairs)}")
        return None
    
    # Get the amount columns for filtering
    amount_cols = [pair['amount_col'] for pair in label_amount_pairs]
    st.write(f"✅ **Using amount columns for filtering:** {amount_cols}")
    
    # Process New Business data (NB) - sum > 0
    nb_df = df[nb_mask].copy()
    if len(nb_df) > 0:
        # Convert amount columns to numeric
        for col in amount_cols:
            nb_df[col] = pd.to_numeric(nb_df[col], errors='coerce').fillna(0)
        
        # Calculate sum of all NCB amounts
        nb_df['NCB_Sum'] = nb_df[amount_cols].sum(axis=1)
        
        # Filter: sum > 0 (as per instructions)
        nb_filtered = nb_df[nb_df['NCB_Sum'] > 0]
        st.write(f"✅ **New Business filtered (sum > 0):** {len(nb_filtered)} records")
        st.write(f"  Expected: ~1200 records")
        st.write(f"  Actual: {len(nb_filtered)} records")
    else:
        nb_filtered = pd.DataFrame()
    
    # Process Reinstatement data (R) - sum > 0
    r_df = df[r_mask].copy()
    if len(r_df) > 0:
        # Convert amount columns to numeric
        for col in amount_cols:
            r_df[col] = pd.to_numeric(r_df[col], errors='coerce').fillna(0)
        
        # Calculate sum of all NCB amounts
        r_df['NCB_Sum'] = r_df[amount_cols].sum(axis=1)
        
        # Filter: sum > 0 (as per instructions)
        r_filtered = r_df[r_df['NCB_Sum'] > 0]
        st.write(f"✅ **Reinstatements filtered (sum > 0):** {len(r_filtered)} records")
    else:
        r_filtered = pd.DataFrame()
    
    # Process Cancellation data (C) - sum < 0
    c_df = df[c_mask].copy()
    if len(c_df) > 0:
        # Convert amount columns to numeric
        for col in amount_cols:
            c_df[col] = pd.to_numeric(c_df[col], errors='coerce').fillna(0)
        
        # Calculate sum of all NCB amounts
        c_df['NCB_Sum'] = c_df[amount_cols].sum(axis=1)
        
        # Filter: sum < 0 (as per instructions)
        c_filtered = c_df[c_df['NCB_Sum'] < 0]
        st.write(f"✅ **Cancellations filtered (sum < 0):** {len(c_filtered)} records")
    else:
        c_filtered = pd.DataFrame()
    
    # Create output dataframes with EXACT column order as specified in instructions
    
    # Data Set 1 - New Contracts (NB)
    nb_output = pd.DataFrame()
    if len(nb_filtered) > 0:
        # Find columns by searching for content (not fixed positions)
        nb_output['Insurer_Code'] = find_column_by_content(nb_filtered, ['INSURER', 'INSURER CODE'])
        nb_output['Product_Type_Code'] = find_column_by_content(nb_filtered, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'])
        nb_output['Coverage_Code'] = find_column_by_content(nb_filtered, ['COVERAGE CODE', 'COVERAGE'])
        nb_output['Dealer_Number'] = find_column_by_content(nb_filtered, ['DEALER NUMBER', 'DEALER #'])
        nb_output['Dealer_Name'] = find_column_by_content(nb_filtered, ['DEALER NAME', 'DEALER'])
        nb_output['Contract_Number'] = find_column_by_content(nb_filtered, ['CONTRACT NUMBER', 'CONTRACT #'])
        nb_output['Contract_Sale_Date'] = find_column_by_content(nb_filtered, ['CONTRACT SALE DATE', 'SALE DATE'])
        nb_output['Transaction_Date'] = find_column_by_content(nb_filtered, ['TRANSACTION DATE', 'ACTIVATION DATE'])
        nb_output['Transaction_Type'] = find_column_by_content(nb_filtered, ['TRANSACTION TYPE'])
        nb_output['Customer_Last_Name'] = find_column_by_content(nb_filtered, ['LAST NAME', 'CUSTOMER LAST NAME'])
        
        # Add NCB amount columns in order
        for i, pair in enumerate(label_amount_pairs):
            col_name = f"NCB_Amount_{i+1}_{pair['label_name'].replace(' ', '_')}"
            nb_output[col_name] = nb_filtered[pair['amount_col']]
        
        nb_output['Transaction_Type'] = 'NB'
        nb_output['Row_Type'] = 'New Business'
    
    # Data Set 2 - Reinstatements (R)
    r_output = pd.DataFrame()
    if len(r_filtered) > 0:
        # Same columns as NB
        r_output['Insurer_Code'] = find_column_by_content(r_filtered, ['INSURER', 'INSURER CODE'])
        r_output['Product_Type_Code'] = find_column_by_content(r_filtered, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'])
        r_output['Coverage_Code'] = find_column_by_content(r_filtered, ['COVERAGE CODE', 'COVERAGE'])
        r_output['Dealer_Number'] = find_column_by_content(r_filtered, ['DEALER NUMBER', 'DEALER #'])
        r_output['Dealer_Name'] = find_column_by_content(r_filtered, ['DEALER NAME', 'DEALER'])
        r_output['Contract_Number'] = find_column_by_content(r_filtered, ['CONTRACT NUMBER', 'CONTRACT #'])
        r_output['Contract_Sale_Date'] = find_column_by_content(r_filtered, ['CONTRACT SALE DATE', 'SALE DATE'])
        r_output['Transaction_Date'] = find_column_by_content(r_filtered, ['TRANSACTION DATE', 'ACTIVATION DATE'])
        r_output['Transaction_Type'] = find_column_by_content(r_filtered, ['TRANSACTION TYPE'])
        r_output['Customer_Last_Name'] = find_column_by_content(r_filtered, ['LAST NAME', 'CUSTOMER LAST NAME'])
        
        # Add NCB amount columns in order
        for i, pair in enumerate(label_amount_pairs):
            col_name = f"NCB_Amount_{i+1}_{pair['label_name'].replace(' ', '_')}"
            r_output[col_name] = r_filtered[pair['amount_col']]
        
        r_output['Transaction_Type'] = 'R'
        r_output['Row_Type'] = 'Reinstatement'
    
    # Data Set 3 - Cancellations (C)
    c_output = pd.DataFrame()
    if len(c_filtered) > 0:
        # Same base columns as NB
        c_output['Insurer_Code'] = find_column_by_content(c_filtered, ['INSURER', 'INSURER CODE'])
        c_output['Product_Type_Code'] = find_column_by_content(c_filtered, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'])
        c_output['Coverage_Code'] = find_column_by_content(c_filtered, ['COVERAGE CODE', 'COVERAGE'])
        c_output['Dealer_Number'] = find_column_by_content(c_filtered, ['DEALER NUMBER', 'DEALER #'])
        c_output['Dealer_Name'] = find_column_by_content(c_filtered, ['DEALER NAME', 'DEALER'])
        c_output['Contract_Number'] = find_column_by_content(c_filtered, ['CONTRACT NUMBER', 'CONTRACT #'])
        c_output['Contract_Sale_Date'] = find_column_by_content(c_filtered, ['CONTRACT SALE DATE', 'SALE DATE'])
        c_output['Transaction_Date'] = find_column_by_content(c_filtered, ['TRANSACTION DATE', 'ACTIVATION DATE'])
        c_output['Transaction_Type'] = find_column_by_content(c_filtered, ['TRANSACTION TYPE'])
        c_output['Customer_Last_Name'] = find_column_by_content(c_filtered, ['LAST NAME', 'CUSTOMER LAST NAME'])
        
        # Additional columns for cancellations
        c_output['Contract_Term'] = find_column_by_content(c_filtered, ['CONTRACT TERM', 'TERM'])
        c_output['Cancellation_Date'] = find_column_by_content(c_filtered, ['CANCELLATION DATE', 'CANCEL DATE'])
        c_output['Cancellation_Reason'] = find_column_by_content(c_filtered, ['CANCELLATION REASON', 'REASON'])
        c_output['Cancellation_Factor'] = find_column_by_content(c_filtered, ['CANCELLATION FACTOR', 'FACTOR'])
        
        # Add NCB amount columns in order
        for i, pair in enumerate(label_amount_pairs):
            col_name = f"NCB_Amount_{i+1}_{pair['label_name'].replace(' ', '_')}"
            c_output[col_name] = c_filtered[pair['amount_col']]
        
        c_output['Transaction_Type'] = 'C'
        c_output['Row_Type'] = 'Cancellation'
    
    # Create combined output for the main download
    combined_output = pd.concat([nb_output, c_output, r_output], ignore_index=True)
    
    st.write(f"✅ **Output created:** {len(combined_output)} total records")
    st.write(f"  New Business: {len(nb_output)} (target: ~1200)")
    st.write(f"  Cancellations: {len(c_output)}")
    st.write(f"  Reinstatements: {len(r_output)}")
    
    return {
        'nb_data': nb_output,
        'cancellation_data': c_output,
        'reinstatement_data': r_output,
        'combined_data': combined_output,
        'total_records': len(combined_output)
    }

def find_column_by_content(df, search_terms):
    """Find a column by searching for specific content in column names or first few rows."""
    for col in df.columns:
        col_str = str(col).upper()
        # Check column name
        for term in search_terms:
            if term.upper() in col_str:
                return df[col]
        
        # Check first few rows for content
        try:
            first_values = df[col].astype(str).str.upper().head(10)
            for term in search_terms:
                if any(term.upper() in val for val in first_values):
                    return df[col]
        except:
            continue
    
    # If not found, return empty series
    return pd.Series([None] * len(df))

def main():
    st.title("🚀 Karen NCB Data Processor - Iteration 2.0")
    st.markdown("---")
    
    st.write("""
    **Expected Output:** 2k-2500 rows in specific order with proper column mapping
    
    **Data Structure:**
    - **New Business (NB)**: Positive/empty/0 values expected
    - **Reinstatement (R)**: Positive/empty/0 values expected  
    - **Cancellation (C)**: Negative/empty/0 values expected
    
    **Column Mapping:** Labels matched to money values dynamically
    
    **Reference File:** Uses 'Col Ref' sheet or similar reference sheet for column mapping
    """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "📁 Upload Excel file with NCB data and reference sheet (Col Ref, xref, etc.)",
        type=['xlsx', 'xls'],
        help="File should contain data sheet and a reference sheet (like 'Col Ref' or 'xref') for column mapping. The reference sheet should define Admin columns and their relationships."
    )
    
    if uploaded_file is not None:
        st.write(f"📁 **File uploaded:** {uploaded_file.name}")
        
        if st.button("🚀 Process Data (Version 2.0)", type="primary", use_container_width=True):
            
            with st.spinner("🔍 Analyzing file structure..."):
                # Detect column structure
                column_mapping, admin_columns, amount_columns = detect_column_structure_v2(uploaded_file)
                
                if not column_mapping:
                    st.error("❌ Could not detect column structure. Please ensure file has a reference sheet (Col Ref, xref, etc.).")
                    st.write("💡 **Tip:** Your file should have a sheet that defines the Admin column structure (like 'Col Ref' or 'xref').")
                    return
            
            with st.spinner("📊 Processing data..."):
                # Read the main data sheet
                try:
                    excel_data = pd.ExcelFile(uploaded_file)
                    
                    # Find the main data sheet (not the reference sheet)
                    data_sheet = None
                    reference_sheet_names = ['Col Ref', 'ColRef', 'Column Reference', 'ColumnRef', 'Reference', 'Col_Ref', 'xref', 'Xref']
                    
                    for sheet in excel_data.sheet_names:
                        if not any(ref_name.lower() in sheet.lower() for ref_name in reference_sheet_names) and not any(keyword in sheet.lower() for keyword in ['summary', 'ref']):
                            data_sheet = sheet
                            break
                    
                    if not data_sheet:
                        st.error("❌ Could not find main data sheet")
                        return
                    
                    st.write(f"📋 **Processing sheet:** {data_sheet}")
                    
                    # Read data
                    df = pd.read_excel(uploaded_file, sheet_name=data_sheet)
                    st.write(f"📏 **Data shape:** {df.shape}")
                    
                    # Process data
                    result = process_data_v2(df, column_mapping, admin_columns, amount_columns)
                    
                    if result:
                        st.success(f"✅ **Processing Complete!** Generated {result['total_records']} records")
                        
                        # Display results
                        st.subheader("📊 Processing Results")
                        st.write(f"**Total Records:** {result['total_records']}")
                        
                        # Show breakdown by transaction type
                        type_counts = result['combined_data']['Transaction_Type'].value_counts()
                        st.write("**Breakdown by Transaction Type:**")
                        for trans_type, count in type_counts.items():
                            st.write(f"  {trans_type}: {count} records")
                        
                        # Show sample data in collapsible section
                        with st.expander("🔍 **Sample Output Data** (Click to expand)", expanded=False):
                            st.dataframe(result['combined_data'].head(10))
                        
                        # Generate Excel file with separate worksheets
                        if st.button("📥 Download Excel File (All Data Sets)"):
                            try:
                                # Create Excel file with separate worksheets
                                output = io.BytesIO()
                                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                    # Data Set 1 - New Contracts (NB)
                                    if len(result['nb_data']) > 0:
                                        result['nb_data'].to_excel(writer, sheet_name='New_Contracts_NB', index=False)
                                        st.write(f"✅ **New Contracts worksheet created:** {len(result['nb_data'])} records")
                                    
                                    # Data Set 2 - Reinstatements (R)
                                    if len(result['reinstatement_data']) > 0:
                                        result['reinstatement_data'].to_excel(writer, sheet_name='Reinstatements_R', index=False)
                                        st.write(f"✅ **Reinstatements worksheet created:** {len(result['reinstatement_data'])} records")
                                    
                                    # Data Set 3 - Cancellations (C)
                                    if len(result['cancellation_data']) > 0:
                                        result['cancellation_data'].to_excel(writer, sheet_name='Cancellations_C', index=False)
                                        st.write(f"✅ **Cancellations worksheet created:** {len(result['cancellation_data'])} records")
                                
                                output.seek(0)
                                
                                # Create download button
                                st.download_button(
                                    label="📥 Download Excel File (3 Worksheets)",
                                    data=output.getvalue(),
                                    file_name=f"NCB_Transaction_Summary_v2_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                                
                                st.success(f"✅ **Excel file generated successfully!**")
                                st.write(f"📊 **File contains:**")
                                st.write(f"  • New Contracts (NB): {len(result['nb_data'])} records")
                                st.write(f"  • Reinstatements (R): {len(result['reinstatement_data'])} records")
                                st.write(f"  • Cancellations (C): {len(result['cancellation_data'])} records")
                                st.write(f"  • **Total:** {result['total_records']} records across 3 worksheets")
                                
                            except Exception as e:
                                st.error(f"❌ Error generating Excel file: {str(e)}")
                                st.write(f"Debug info: {type(e).__name__}")
                        
                        # Individual download buttons for each data set
                        st.write("---")
                        st.write("### 📥 Individual Data Set Downloads")
                        
                        # New Business download
                        if len(result['nb_data']) > 0:
                            if st.button(f"📥 Download New Business Data ({len(result['nb_data'])} records)"):
                                nb_output = io.BytesIO()
                                result['nb_data'].to_excel(nb_output, index=False)
                                nb_output.seek(0)
                                
                                st.download_button(
                                    label=f"📥 Download NB Data ({len(result['nb_data'])} records)",
                                    data=nb_output.getvalue(),
                                    file_name=f"NCB_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        # Cancellation download
                        if len(result['cancellation_data']) > 0:
                            if st.button(f"📥 Download Cancellation Data ({len(result['cancellation_data'])} records)"):
                                c_output = io.BytesIO()
                                result['cancellation_data'].to_excel(c_output, index=False)
                                c_output.seek(0)
                                
                                st.download_button(
                                    label=f"📥 Download Cancellation Data ({len(result['cancellation_data'])} records)",
                                    data=c_output.getvalue(),
                                    file_name=f"NCB_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        # Reinstatement download
                        if len(result['reinstatement_data']) > 0:
                            if st.button(f"📥 Download Reinstatement Data ({len(result['reinstatement_data'])} records)"):
                                r_output = io.BytesIO()
                                result['reinstatement_data'].to_excel(r_output, index=False)
                                r_output.seek(0)
                                
                                st.download_button(
                                    label=f"📥 Download Reinstatement Data ({len(result['reinstatement_data'])} records)",
                                    data=r_output.getvalue(),
                                    file_name=f"NCB_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                    else:
                        st.error("❌ No data was processed successfully")
                        
                except Exception as e:
                    st.error(f"❌ Error processing file: {e}")
                    import traceback
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
