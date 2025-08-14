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
    """Detect column structure from Col Ref sheet for version 2.0."""
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
            st.write("🔍 **Looking for alternative reference sheets...**")
            
            # Try to find any sheet that might contain column information
            for sheet_name in excel_data.sheet_names:
                if not any(keyword in sheet_name.lower() for keyword in ['data', 'transaction', 'summary']):
                    st.write(f"🔍 **Checking sheet:** {sheet_name}")
                    try:
                        test_df = pd.read_excel(excel_file, sheet_name=sheet_name)
                        # Look for Admin or NCB content in this sheet
                        content_str = ' '.join([str(val) for val in test_df.values.flatten() if pd.notna(val)]).upper()
                        if any(keyword in content_str for keyword in ['ADMIN', 'NCB', 'AGENT', 'DEALER']):
                            st.write(f"✅ **Found potential reference sheet:** {sheet_name}")
                            col_ref_sheet = sheet_name
                            break
                    except:
                        continue
            
            if not col_ref_sheet:
                st.error("❌ No suitable reference sheet found. Please ensure your Excel file has a 'Col Ref' sheet or similar.")
                return {}, {}, {}
        
        st.write(f"✅ **Using reference sheet:** {col_ref_sheet}")
        
        # Read the reference sheet
        col_ref_df = pd.read_excel(excel_file, sheet_name=col_ref_sheet)
        
        st.write(f"🔍 **Reference sheet loaded:** {col_ref_df.shape}")
        st.write(f"🔍 **Reference sheet first few rows:**")
        st.dataframe(col_ref_df.head(10))
        
        # Look for the specific Admin column structure
        column_mapping = {}
        label_columns = {}
        amount_columns = {}
        
        # Search through all rows for Admin column information
        for idx, row in col_ref_df.iterrows():
            row_str = ' '.join([str(val) for val in row if pd.notna(val)]).upper()
            
            # Check if this row contains Admin column information
            if any(keyword in row_str for keyword in ['ADMIN', 'NCB', 'AGENT', 'DEALER']):
                st.write(f"🔍 **Found potential Admin row {idx}:** {row_str[:100]}...")
                
                # Look for Admin columns in this row
                for col_idx, value in enumerate(row):
                    if pd.notna(value) and str(value).strip():
                        desc = str(value).strip()
                        column_mapping[col_idx] = desc
                        
                        # Check if this is an Admin Label column
                        if any(keyword in desc.upper() for keyword in ['ADMIN', 'NCB']) and 'AMOUNT' not in desc.upper():
                            label_columns[col_idx] = desc
                            st.write(f"  ✅ **Found Label column:** {desc} at position {col_idx}")
                            
                            # Look for the corresponding Amount column (next column over)
                            if col_idx + 1 < len(row):
                                next_col_val = row.iloc[col_idx + 1]
                                if pd.notna(next_col_val) and str(next_col_val).strip():
                                    next_desc = str(next_col_val).strip()
                                    if 'AMOUNT' in next_desc.upper():
                                        amount_columns[col_idx + 1] = next_desc
                                        st.write(f"    ✅ **Found Amount column:** {next_desc} at position {col_idx + 1}")
                
                # If we found Admin columns, we can stop searching
                if label_columns:
                    break
        
        # If we still don't have Admin columns, try a different approach
        if not label_columns:
            st.write("🔄 **Trying alternative Admin column detection...**")
            
            # Look for columns that contain "Admin" in their names
            for col_idx, col_name in enumerate(col_ref_df.columns):
                if pd.notna(col_name) and 'ADMIN' in str(col_name).upper():
                    st.write(f"🔍 **Found Admin column:** {col_name} at position {col_idx}")
                    
                    # Check if this column contains Admin information
                    col_values = col_ref_df[col_name].dropna()
                    for val in col_values:
                        if pd.notna(val) and str(val).strip():
                            desc = str(val).strip()
                            if any(keyword in desc.upper() for keyword in ['ADMIN', 'NCB', 'AGENT', 'DEALER']) and 'AMOUNT' not in desc.upper():
                                label_columns[col_idx] = desc
                                st.write(f"  ✅ **Found Label value:** {desc}")
                                
                                # Look for corresponding Amount column
                                if col_idx + 1 < len(col_ref_df.columns):
                                    next_col = col_ref_df.columns[col_idx + 1]
                                    if pd.notna(next_col) and 'AMOUNT' in str(next_col).upper():
                                        amount_columns[col_idx + 1] = str(next_col)
                                        st.write(f"    ✅ **Found Amount column:** {next_col}")
                                break
        
        if label_columns:
            st.write("🔍 **Column Structure Detected:**")
            st.write(f"  Total columns mapped: {len(column_mapping)}")
            st.write(f"  Label columns found: {len(label_columns)}")
            st.write(f"  Amount columns found: {len(amount_columns)}")
            
            st.write("🔍 **Admin column mappings:**")
            for col_idx, desc in label_columns.items():
                st.write(f"  {desc}: Column {col_idx}")
                if col_idx + 1 in amount_columns:
                    st.write(f"    Amount: Column {col_idx + 1}")
            
            return column_mapping, label_columns, amount_columns
        else:
            st.warning("⚠️ No Admin column mapping found in reference sheet")
            st.write("🔍 **Debugging: Let's examine the reference sheet more carefully...**")
            
            # Show more detailed information about the reference sheet
            st.write("🔍 **All column names in reference sheet:**")
            for i, col_name in enumerate(col_ref_df.columns):
                st.write(f"  Column {i}: {col_name}")
            
            st.write("🔍 **Sample values from first few rows:**")
            for i in range(min(5, len(col_ref_df))):
                row_data = []
                for j, val in enumerate(col_ref_df.iloc[i]):
                    if pd.notna(val) and str(val).strip():
                        row_data.append(f"Col{j}:{str(val).strip()}")
                if row_data:
                    st.write(f"  Row {i}: {' | '.join(row_data[:10])}")
            
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

def process_data_v2(df, column_mapping, label_columns, amount_columns):
    """Process data according to version 2.0 requirements."""
    
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
    
    # Process New Business data with the WORKING filtering logic from old version
    nb_df = df[nb_mask].copy()
    
    # Apply the EXACT filtering logic that was working before (giving ~1200 records)
    if len(label_columns) >= 4:
        # Get the Admin amount columns (assuming they're in order)
        admin_cols = list(label_columns.keys())[:4]
        st.write(f"🔍 **Using Admin columns:** {admin_cols}")
        
        # Convert Admin columns to numeric
        for col in admin_cols:
            nb_df[col] = pd.to_numeric(nb_df[col], errors='coerce').fillna(0)
        
        # Calculate Admin sum
        nb_df['Admin_Sum'] = nb_df[admin_cols].sum(axis=1)
        
        # Apply the EXACT user requirement: ALL 4 Admin amounts > 0 AND sum > 0
        nb_filtered = nb_df[
            (nb_df['Admin_Sum'] > 0) &
            (nb_df[admin_cols[0]] > 0) &
            (nb_df[admin_cols[1]] > 0) &
            (nb_df[admin_cols[2]] > 0) &
            (nb_df[admin_cols[3]] > 0)
        ]
        
        st.write(f"✅ **New Business filtered:** {len(nb_filtered)} records (using strict Admin > 0 logic)")
        st.write(f"  Expected: ~1200 records")
        st.write(f"  Actual: {len(nb_filtered)} records")
    else:
        nb_filtered = nb_df
        st.write(f"⚠️ **Using unfiltered New Business data:** {len(nb_filtered)} records")
    
    # Process Cancellation data (negative/empty/0 values expected)
    c_df = df[c_mask].copy()
    
    # Apply filtering for cancellations (sum != 0)
    if len(label_columns) >= 4:
        admin_cols = list(label_columns.keys())[:4]
        for col in admin_cols:
            c_df[col] = pd.to_numeric(c_df[col], errors='coerce').fillna(0)
        c_df['Admin_Sum'] = c_df[admin_cols].sum(axis=1)
        c_filtered = c_df[c_df['Admin_Sum'] != 0]
        st.write(f"✅ **Cancellations filtered:** {len(c_filtered)} records")
    else:
        c_filtered = c_df
        st.write(f"⚠️ **Using unfiltered Cancellation data:** {len(c_filtered)} records")
    
    # Process Reinstatement data (positive/empty/0 values expected)
    r_df = df[r_mask].copy()
    
    # Apply filtering for reinstatements (sum != 0)
    if len(label_columns) >= 4:
        admin_cols = list(label_columns.keys())[:4]
        for col in admin_cols:
            r_df[col] = pd.to_numeric(r_df[col], errors='coerce').fillna(0)
        r_df['Admin_Sum'] = r_df[admin_cols].sum(axis=1)
        r_filtered = r_df[r_df['Admin_Sum'] != 0]
        st.write(f"✅ **Reinstatements filtered:** {len(r_filtered)} records")
    else:
        r_filtered = r_df
        st.write(f"⚠️ **Using unfiltered Reinstatement data:** {len(r_filtered)} records")
    
    # Create separate dataframes for each transaction type (as in old format)
    nb_output = nb_filtered.copy()
    c_output = c_filtered.copy()
    r_output = r_filtered.copy()
    
    # Add transaction type identifiers
    nb_output['Transaction_Type'] = 'NB'
    nb_output['Row_Type'] = 'New Business'
    c_output['Transaction_Type'] = 'C'
    c_output['Row_Type'] = 'Cancellation'
    r_output['Transaction_Type'] = 'R'
    r_output['Row_Type'] = 'Reinstatement'
    
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
                column_mapping, label_columns, amount_columns = detect_column_structure_v2(uploaded_file)
                
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
                    output_data = process_data_v2(df, column_mapping, label_columns, amount_columns)
                    
                    if output_data:
                        st.success(f"✅ **Processing Complete!** Generated {output_data['total_records']} records")
                        
                        # Display results
                        st.subheader("📊 Processing Results")
                        st.write(f"**Total Records:** {output_data['total_records']}")
                        
                        # Show breakdown by transaction type
                        type_counts = output_data['combined_data']['Transaction_Type'].value_counts()
                        st.write("**Breakdown by Transaction Type:**")
                        for trans_type, count in type_counts.items():
                            st.write(f"  {trans_type}: {count} records")
                        
                        # Show sample data in collapsible section
                        with st.expander("🔍 **Sample Output Data** (Click to expand)", expanded=False):
                            st.dataframe(output_data['combined_data'].head(10))
                        
                        # Download buttons for each section
                        st.subheader("💾 Download Results")
                        
                        # Create Excel files in memory for each section
                        nb_buffer = io.BytesIO()
                        c_buffer = io.BytesIO()
                        r_buffer = io.BytesIO()
                        
                        with pd.ExcelWriter(nb_buffer, engine='xlsxwriter') as writer:
                            output_data['nb_data'].to_excel(writer, sheet_name='New_Business', index=False)
                        with pd.ExcelWriter(c_buffer, engine='xlsxwriter') as writer:
                            output_data['cancellation_data'].to_excel(writer, sheet_name='Cancellation', index=False)
                        with pd.ExcelWriter(r_buffer, engine='xlsxwriter') as writer:
                            output_data['reinstatement_data'].to_excel(writer, sheet_name='Reinstatement', index=False)
                        
                        # Generate filenames with timestamp
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        nb_filename = f"Karen_NCB_2_0_New_Business_{timestamp}.xlsx"
                        c_filename = f"Karen_NCB_2_0_Cancellation_{timestamp}.xlsx"
                        r_filename = f"Karen_NCB_2_0_Reinstatement_{timestamp}.xlsx"
                        combined_filename = f"Karen_NCB_2_0_Combined_{timestamp}.xlsx"
                        
                        # Download buttons for each section
                        st.download_button(
                            label="📥 Download New Business Excel File",
                            data=nb_buffer.getvalue(),
                            file_name=nb_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="📥 Download Cancellation Excel File",
                            data=c_buffer.getvalue(),
                            file_name=c_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="📥 Download Reinstatement Excel File",
                            data=r_buffer.getvalue(),
                            file_name=r_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        # Create combined download
                        combined_buffer = io.BytesIO()
                        with pd.ExcelWriter(combined_buffer, engine='xlsxwriter') as writer:
                            output_data['combined_data'].to_excel(writer, sheet_name='Combined_Data', index=False)
                        
                        combined_buffer.seek(0)
                        
                        st.download_button(
                            label="📥 Download Combined Excel File",
                            data=combined_buffer.getvalue(),
                            file_name=combined_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.write(f"📁 **Files saved as:**")
                        st.write(f"  New Business: {nb_filename}")
                        st.write(f"  Cancellation: {c_filename}")
                        st.write(f"  Reinstatement: {r_filename}")
                        st.write(f"  Combined: {combined_filename}")
                        
                    else:
                        st.error("❌ No data was processed successfully")
                        
                except Exception as e:
                    st.error(f"❌ Error processing file: {e}")
                    import traceback
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
