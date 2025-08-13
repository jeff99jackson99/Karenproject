#!/usr/bin/env python3
"""
Smart NCB Data Processor - Works with unnamed columns by detecting structure from Col Ref sheet
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Smart NCB Data Processor", page_icon="🧠", layout="wide")

def detect_column_structure(excel_file):
    """Detect the actual column structure from the Col Ref sheet."""
    try:
        # Read the Col Ref sheet to understand column mapping
        col_ref_df = pd.read_excel(excel_file, sheet_name='Col Ref')
        
        # Look for the row that contains column descriptions
        column_mapping = {}
        
        # Find the row with column descriptions (usually row 1)
        for idx, row in col_ref_df.iterrows():
            if 'Admin' in str(row.iloc[2]) or 'NCB' in str(row.iloc[7]):
                # This is the row with column descriptions
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and str(value).strip():
                        column_mapping[col_idx] = str(value).strip()
                break
        
        st.write("🔍 **Column Structure Detected:**")
        for col_idx, desc in column_mapping.items():
            st.write(f"  Column {col_idx}: {desc}")
        
        return column_mapping
        
    except Exception as e:
        st.warning(f"⚠️ Could not read 'Col Ref' sheet: {e}")
        st.write("🔄 **Trying alternative column detection methods...**")
        return {}

def find_admin_columns(df, column_mapping):
    """Find Admin columns based on the detected structure."""
    admin_columns = {}
    
    # Look for Admin columns based on the mapping
    for col_idx, desc in column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            
            if 'Admin' in desc or 'NCB' in desc:
                if 'Amount' in desc or 'Fee' in desc:
                    # This is an Admin amount column
                    if 'Agent' in desc:
                        admin_columns['Admin 3'] = col_name
                    elif 'Dealer' in desc:
                        admin_columns['Admin 4'] = col_name
                    elif 'Offset' in desc:
                        if 'Agent' in desc:
                            admin_columns['Admin 9'] = col_name
                        elif 'Dealer' in desc:
                            admin_columns['Admin 10'] = col_name
    
    # If we don't have enough admin columns, try to find them by content analysis
    if len(admin_columns) < 4:
        st.warning("⚠️ Could not identify all Admin columns from mapping, trying content-based detection...")
        
        # Look for columns that might contain Admin amounts by analyzing their content
        potential_admin_cols = []
        
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
    
    return admin_columns

def process_excel_data_smart(uploaded_file):
    """Process Excel data using smart column detection."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        st.write("🔍 **Smart Processing Started**")
        st.write(f"📊 Sheets found: {excel_data.sheet_names}")
        
        # Detect column structure
        column_mapping = detect_column_structure(uploaded_file)
        
        # Process the Data sheet (main transaction data)
        data_sheet_name = None
        
        # Look for common data sheet names
        for sheet_name in excel_data.sheet_names:
            if any(keyword in sheet_name.lower() for keyword in ['data', 'transaction', 'detail', 'main']):
                data_sheet_name = sheet_name
                break
        
        # If no obvious data sheet found, try the first sheet that's not a summary
        if not data_sheet_name:
            for sheet_name in excel_data.sheet_names:
                if not any(keyword in sheet_name.lower() for keyword in ['summary', 'xref', 'ref', 'instruction']):
                    data_sheet_name = sheet_name
                    break
        
        if data_sheet_name:
            st.write(f"📋 **Processing {data_sheet_name} Sheet**")
            
            # Read the data sheet
            df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name)
            st.write(f"📏 Data shape: {df.shape}")
            
            # Find admin columns
            admin_columns = find_admin_columns(df, column_mapping)
            st.write("🔍 **Admin Columns Found:**")
            for admin_type, col_name in admin_columns.items():
                st.write(f"  {admin_type}: {col_name}")
            
            if len(admin_columns) >= 4:
                # Process the data
                results = process_transaction_data(df, admin_columns)
                
                if results:
                    # Rename columns with meaningful names
                    renamed_nb, nb_rename_map = rename_columns_with_meaning(results['nb_data'], column_mapping)
                    renamed_cancellation, c_rename_map = rename_columns_with_meaning(results['cancellation_data'], column_mapping)
                    renamed_reinstatement, r_rename_map = rename_columns_with_meaning(results['reinstatement_data'], column_mapping)
                    
                    # Show column renaming information
                    st.write("🔍 **Column Renaming Applied:**")
                    if nb_rename_map:
                        st.write("  **New Business columns renamed:**")
                        for old_name, new_name in nb_rename_map.items():
                            st.write(f"    {old_name} → {new_name}")
                    
                    # Update results with renamed dataframes
                    results['nb_data'] = renamed_nb
                    results['cancellation_data'] = renamed_cancellation
                    results['reinstatement_data'] = renamed_reinstatement
                    results['column_mapping'] = column_mapping
                    
                    return results
                else:
                    return None
            else:
                st.error(f"❌ Need at least 4 Admin columns, found {len(admin_columns)}")
                return None
        else:
            st.error("❌ No suitable data sheet found")
            return None
            
    except Exception as e:
        st.error(f"❌ Error in smart processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def process_transaction_data(df, admin_columns):
    """Process transaction data with the detected admin columns."""
    try:
        # Get the admin column names
        admin_cols = [
            admin_columns.get('Admin 3'),
            admin_columns.get('Admin 4'), 
            admin_columns.get('Admin 9'),
            admin_columns.get('Admin 10')
        ]
        
        # Remove None values
        admin_cols = [col for col in admin_cols if col is not None]
        
        if len(admin_cols) < 4:
            st.error(f"❌ Need exactly 4 Admin columns, found {len(admin_cols)}")
            return None
        
        st.write(f"✅ **Processing with Admin columns:** {admin_cols}")
        
        # Convert admin columns to numeric
        for col in admin_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ Column {col} not found in data")
                return None
        
        # Calculate sum of admin amounts
        df['Admin_Sum'] = df[admin_cols].sum(axis=1)
        
        # Find transaction type column
        transaction_col = None
        
        # First, look for columns that actually contain transaction type data
        for col in df.columns:
            if df[col].dtype == 'object':
                sample_vals = df[col].dropna().head(10).tolist()
                # Check if this column contains transaction types (NB, C, R)
                if any(str(v).upper() in ['NB', 'C', 'R', 'NEW BUSINESS', 'CANCELLATION', 'REINSTATEMENT'] for v in sample_vals):
                    transaction_col = col
                    st.write(f"✅ **Found transaction column by content:** {col} with values: {sample_vals[:5]}")
                    break
        
        # If not found by content, fall back to name-based search
        if not transaction_col:
            for col in df.columns:
                if 'transaction' in str(col).lower() or 'type' in str(col).lower():
                    transaction_col = col
                    break
        
        if not transaction_col:
            st.error("❌ No transaction type column found")
            return None
        
        st.write(f"✅ **Transaction column found:** {transaction_col}")
        
        # Show unique values in transaction column to understand the data
        unique_transactions = df[transaction_col].astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")  # Show first 20
        
        # Filter by transaction type and admin amounts
        nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
        c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
        r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        
        # If no transactions found, try to understand the data better
        if nb_mask.sum() == 0 and c_mask.sum() == 0 and r_mask.sum() == 0:
            st.warning("⚠️ No standard transaction types found. Let me examine the data...")
            
            # Show sample values from the transaction column
            sample_values = df[transaction_col].dropna().head(20).tolist()
            st.write(f"🔍 **Sample transaction values:** {sample_values}")
            
            # Try to find any non-null values that might be transaction types
            non_null_values = df[transaction_col].dropna().unique()
            st.write(f"🔍 **All non-null transaction values:** {non_null_values[:20]}")
            
            # Look for patterns in the data
            for col in df.columns[:10]:  # Check first 10 columns
                if df[col].dtype == 'object':
                    sample_vals = df[col].dropna().head(5).tolist()
                    if any('NB' in str(v) or 'C' in str(v) or 'R' in str(v) for v in sample_vals):
                        st.write(f"🔍 **Potential transaction column {col}:** {sample_vals}")
        
        # Filter NB data (sum > 0 and all 4 admin values present)
        nb_df = df[nb_mask].copy()
        
        # Show Admin amount statistics for NB data
        st.write(f"🔍 **New Business Admin Amount Analysis:**")
        for i, col in enumerate(admin_cols):
            col_values = nb_df[col]
            non_zero_count = (col_values > 0).sum()
            zero_count = (col_values == 0).sum()
            st.write(f"  {col}: {non_zero_count} non-zero, {zero_count} zero values")
        
        # Show sum statistics
        nb_df['Admin_Sum'] = nb_df[admin_cols].sum(axis=1)
        st.write(f"  Admin Sum > 0: {(nb_df['Admin_Sum'] > 0).sum()} records")
        st.write(f"  Admin Sum = 0: {(nb_df['Admin_Sum'] == 0).sum()} records")
        
        # More flexible filtering - try different approaches
        st.write(f"🔍 **Trying different filtering approaches:**")
        
        # Approach 1: Sum > 0 (original requirement)
        nb_filtered_1 = nb_df[nb_df['Admin_Sum'] > 0]
        st.write(f"  Approach 1 (Sum > 0): {len(nb_filtered_1)} records")
        
        # Approach 2: At least 2 Admin amounts > 0
        admin_gt_zero = (nb_df[admin_cols] > 0).sum(axis=1)
        nb_filtered_2 = nb_df[admin_gt_zero >= 2]
        st.write(f"  Approach 2 (≥2 Admin > 0): {len(nb_filtered_2)} records")
        
        # Approach 3: Any Admin amount > 0
        nb_filtered_3 = nb_df[admin_gt_zero >= 1]
        st.write(f"  Approach 3 (≥1 Admin > 0): {len(nb_filtered_3)} records")
        
        # Approach 4: EXACT USER REQUIREMENT - ALL 4 Admin amounts > 0 AND sum > 0
        nb_filtered_4 = nb_df[
            (nb_df['Admin_Sum'] > 0) &
            (nb_df[admin_cols[0]] > 0) &
            (nb_df[admin_cols[1]] > 0) &
            (nb_df[admin_cols[2]] > 0) &
            (nb_df[admin_cols[3]] > 0)
        ]
        st.write(f"  Approach 4 (ALL 4 Admin > 0 AND Sum > 0): {len(nb_filtered_4)} records")
        
        # Use the EXACT user requirement (Approach 4)
        nb_filtered = nb_filtered_4
        st.write(f"✅ **Using EXACT user requirement (ALL 4 Admin > 0 AND Sum > 0): {len(nb_filtered)} records**")
        
        # Show detailed breakdown of why records are filtered out
        st.write(f"🔍 **Detailed filtering breakdown:**")
        st.write(f"  Records with Admin Sum > 0: {len(nb_filtered_1)}")
        st.write(f"  Records with ALL 4 Admin > 0: {len(nb_filtered_4)}")
        st.write(f"  Records filtered out by Admin requirement: {len(nb_filtered_1) - len(nb_filtered_4)}")
        
        # Show sample of records that meet sum > 0 but not all 4 Admin > 0
        if len(nb_filtered_1) > len(nb_filtered_4):
            sample_filtered_out = nb_df[
                (nb_df['Admin_Sum'] > 0) &
                ~((nb_df[admin_cols[0]] > 0) &
                  (nb_df[admin_cols[1]] > 0) &
                  (nb_df[admin_cols[2]] > 0) &
                  (nb_df[admin_cols[3]] > 0))
            ].head(5)
            
            st.write(f"🔍 **Sample records filtered out (Sum > 0 but not all 4 Admin > 0):**")
            for i, col in enumerate(admin_cols):
                st.write(f"  {col}: {sample_filtered_out[col].tolist()}")
            st.write(f"  Admin Sum: {sample_filtered_out['Admin_Sum'].tolist()}")
        
        # Filter C/R data (sum != 0)
        cr_df = df[c_mask | r_mask].copy()
        cr_filtered = cr_df[cr_df['Admin_Sum'] != 0]
        
        st.write(f"📊 **Filtered results:**")
        st.write(f"  New Business (filtered): {len(nb_filtered)}")
        st.write(f"  Cancellations/Reinstatements (filtered): {len(cr_filtered)}")
        
        # Select only the columns we need for output
        needed_columns = [
            admin_cols[0], admin_cols[1], admin_cols[2], admin_cols[3],  # Admin columns
            transaction_col,  # Transaction type
            'Admin_Sum'  # Calculated sum
        ]
        
        # Add other important columns if they exist
        for col in ['Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3', 'Unnamed: 5', 'Unnamed: 6', 'Unnamed: 8']:
            if col in df.columns:
                needed_columns.append(col)
        
        # Filter dataframes to only include needed columns
        nb_filtered = nb_filtered[needed_columns]
        
        # For C/R data, first filter by transaction type, then select columns
        cancellation_data = cr_df[cr_df[transaction_col] == 'C'][needed_columns]
        reinstatement_data = cr_df[cr_df[transaction_col] == 'R'][needed_columns]
        
        st.write(f"🔍 **Selected columns for output:** {needed_columns}")
        st.write(f"📊 **Output dataframe shapes:** NB: {nb_filtered.shape}, C: {cancellation_data.shape}, R: {reinstatement_data.shape}")
        
        return {
            'nb_data': nb_filtered,
            'cancellation_data': cancellation_data,
            'reinstatement_data': reinstatement_data,
            'admin_columns': admin_cols,
            'total_records': len(nb_filtered) + len(cancellation_data) + len(reinstatement_data)
        }
        
    except Exception as e:
        st.error(f"❌ Error processing transaction data: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_excel_download(df, sheet_name):
    """Create Excel download for a dataframe."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

def rename_columns_with_meaning(df, column_mapping):
    """Rename unnamed columns with their actual meanings."""
    renamed_df = df.copy()
    
    # Create a mapping from unnamed columns to meaningful names
    column_rename_map = {}
    
    # Map Admin columns based on their position and the Col Ref sheet
    admin_column_mapping = {
        0: 'Admin_Base',      # Admin(Base)
        4: 'Admin_CLP',       # Admin(CLP) 
        7: 'Agent_NCB_Fees',  # Agent NCB Fees
        16: 'Dealer_NCB_Fees' # Dealer NCB Fees
    }
    
    # Apply Admin column renaming
    for col_idx, new_name in admin_column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            column_rename_map[col_name] = new_name
    
    # Map other important columns based on the Col Ref sheet
    other_column_mapping = {
        1: 'Insurer_Code',
        2: 'Product_Type_Code', 
        3: 'Coverage_Code',
        5: 'Dealer_Name',
        6: 'Dealer_Insured_State',
        8: 'Reinsurance_Support',
        9: 'Transaction_Type'
    }
    
    # Apply other column renaming
    for col_idx, new_name in other_column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            if col_name not in column_rename_map:  # Don't overwrite Admin columns
                column_rename_map[col_name] = new_name
    
    # Also rename some common columns we know about
    for col in df.columns:
        if 'RPT908' in str(col):
            column_rename_map[col] = 'Coverage_Code'
        elif 'Transaction Type' in str(df[col].iloc[0] if len(df) > 0 else ''):
            column_rename_map[col] = 'Transaction_Type'
    
    # Apply the renaming
    renamed_df = renamed_df.rename(columns=column_rename_map)
    
    # Debug: Show what was renamed
    st.write(f"🔍 **Column renaming debug for {len(column_rename_map)} columns:**")
    for old_name, new_name in column_rename_map.items():
        st.write(f"  {old_name} → {new_name}")
    
    # Debug: Show final column names
    st.write(f"🔍 **Final column names (first 10):** {list(renamed_df.columns)[:10]}")
    
    return renamed_df, column_rename_map

def main():
    st.title("🧠 Smart NCB Data Processor")
    st.markdown("**Automatically detects column structure and processes NCB data**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Smart Processing")
        st.markdown("**This app will:**")
        st.markdown("• Detect column structure automatically")
        st.markdown("• Map unnamed columns to Admin fields")
        st.markdown("• Process data based on actual structure")
        st.markdown("• Show detailed debugging information")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file for processing", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🧠 Start Smart Processing", type="primary", use_container_width=True):
            with st.spinner("Smart processing in progress..."):
                results = process_excel_data_smart(uploaded_file)
                
                if results:
                    st.success("✅ Smart processing complete!")
                    
                    # Show results
                    st.header("📊 Processing Results")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col3:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    with col4:
                        st.metric("Total Records", results['total_records'])
                    
                    # Show admin columns used
                    st.subheader("🔍 Admin Columns Used")
                    st.write(f"**Columns processed:** {results['admin_columns']}")
                    
                    # Show column mapping if available
                    if 'column_mapping' in results:
                        st.write("**Column meanings from Col Ref sheet:**")
                        for col_idx, desc in results['column_mapping'].items():
                            st.write(f"  Column {col_idx}: {desc}")
                    
                    # Show sample data
                    if len(results['nb_data']) > 0:
                        st.subheader("🆕 Sample New Business Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['nb_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['nb_data'].shape}, Columns: {len(results['nb_data'].columns)}")
                        st.dataframe(results['nb_data'].head(10))
                        
                        # Download button
                        st.download_button(
                            label="Download New Business Data",
                            data=create_excel_download(results['nb_data'], 'New Business'),
                            file_name=f"NB_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        st.subheader("❌ Sample Cancellation Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['cancellation_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['cancellation_data'].shape}, Columns: {len(results['cancellation_data'].columns)}")
                        st.dataframe(results['cancellation_data'].head(10))
                        
                        st.download_button(
                            label="Download Cancellation Data",
                            data=create_excel_download(results['cancellation_data'], 'Cancellations'),
                            file_name=f"Cancellation_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        st.subheader("🔄 Sample Reinstatement Data")
                        st.write("**Columns in this data:**")
                        # Force display of actual column names
                        st.write([str(col) for col in results['reinstatement_data'].columns])
                        st.write(f"**Dataframe info:** Shape: {results['reinstatement_data'].shape}, Columns: {len(results['reinstatement_data'].columns)}")
                        st.dataframe(results['reinstatement_data'].head(10))
                        
                        st.download_button(
                            label="Download Reinstatement Data",
                            data=create_excel_download(results['reinstatement_data'], 'Reinstatements'),
                            file_name=f"Reinstatement_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("❌ Smart processing failed. Check the error messages above.")

if __name__ == "__main__":
    main()
