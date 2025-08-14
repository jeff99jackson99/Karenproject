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
        possible_names = ['Col Ref', 'ColRef', 'Column Reference', 'ColumnRef', 'Reference', 'Col_Ref']
        
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
    
    # Process New Business data (positive/empty/0 values expected)
    nb_df = df[nb_mask].copy()
    
    # Process Cancellation data (negative/empty/0 values expected)
    c_df = df[c_mask].copy()
    
    # Process Reinstatement data (positive/empty/0 values expected)
    r_df = df[r_mask].copy()
    
    # Create the output structure with proper column mapping
    output_data = []
    
    # Add New Business records first (as per user requirement for specific order)
    st.write("🔄 **Processing New Business records...**")
    for _, row in nb_df.iterrows():
        record = {
            'Transaction_Type': 'NB',
            'Row_Type': 'New Business'
        }
        
        # Add all the mapped columns with proper labels
        for col_idx, desc in column_mapping.items():
            if col_idx < len(row):
                col_name = row.index[col_idx]
                record[f"Col_{col_idx}_{desc}"] = row[col_name]
        
        # Validate NB data: should have positive/empty/0 values for Admin amounts
        nb_valid = True
        for label_col_idx in label_columns:
            if label_col_idx + 1 in amount_columns:  # Check if we have the corresponding amount column
                amount_col_idx = label_col_idx + 1
                if amount_col_idx < len(row):
                    amount_val = row.iloc[amount_col_idx]
                    try:
                        amount_val = pd.to_numeric(amount_val, errors='coerce')
                        if pd.notna(amount_val) and amount_val < 0:
                            nb_valid = False
                            break
                    except:
                        pass
        
        if nb_valid:
            output_data.append(record)
    
    st.write(f"✅ **New Business records processed:** {len(output_data)}")
    
    # Add Cancellation records (negative/empty/0 values expected)
    st.write("🔄 **Processing Cancellation records...**")
    for _, row in c_df.iterrows():
        record = {
            'Transaction_Type': 'C',
            'Row_Type': 'Cancellation'
        }
        
        # Add all the mapped columns
        for col_idx, desc in column_mapping.items():
            if col_idx < len(row):
                col_name = row.index[col_idx]
                record[f"Col_{col_idx}_{desc}"] = row[col_name]
        
        # Validate C data: should have negative/empty/0 values for Admin amounts
        c_valid = True
        for label_col_idx in label_columns:
            if label_col_idx + 1 in amount_columns:
                amount_col_idx = label_col_idx + 1
                if amount_col_idx < len(row):
                    amount_val = row.iloc[amount_col_idx]
                    try:
                        amount_val = pd.to_numeric(amount_val, errors='coerce')
                        if pd.notna(amount_val) and amount_val > 0:
                            c_valid = False
                            break
                    except:
                        pass
        
        if c_valid:
            output_data.append(record)
    
    st.write(f"✅ **Cancellation records processed:** {len([r for r in output_data if r['Transaction_Type'] == 'C'])}")
    
    # Add Reinstatement records (positive/empty/0 values expected)
    st.write("🔄 **Processing Reinstatement records...**")
    for _, row in r_df.iterrows():
        record = {
            'Transaction_Type': 'R',
            'Row_Type': 'Reinstatement'
        }
        
        # Add all the mapped columns
        for col_idx, desc in column_mapping.items():
            if col_idx < len(row):
                col_name = row.index[col_idx]
                record[f"Col_{col_idx}_{desc}"] = row[col_name]
        
        # Validate R data: should have positive/empty/0 values for Admin amounts
        r_valid = True
        for label_col_idx in label_columns:
            if label_col_idx + 1 in amount_columns:
                amount_col_idx = label_col_idx + 1
                if amount_col_idx < len(row):
                    amount_val = row.iloc[amount_col_idx]
                    try:
                        amount_val = pd.to_numeric(amount_val, errors='coerce')
                        if pd.notna(amount_val) and amount_val < 0:
                            r_valid = False
                            break
                    except:
                        pass
        
        if r_valid:
            output_data.append(record)
    
    st.write(f"✅ **Reinstatement records processed:** {len([r for r in output_data if r['Transaction_Type'] == 'R'])}")
    
    # Create output dataframe
    output_df = pd.DataFrame(output_data)
    
    st.write(f"✅ **Output created:** {len(output_df)} total records")
    st.write(f"  Expected range: 2,000 - 2,500 rows")
    st.write(f"  Actual output: {len(output_df)} rows")
    
    # Show breakdown by transaction type
    if len(output_df) > 0:
        type_counts = output_df['Transaction_Type'].value_counts()
        st.write("📊 **Final breakdown by Transaction Type:**")
        for trans_type, count in type_counts.items():
            st.write(f"  {trans_type}: {count} records")
    
    return output_df

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
    """)
    
    # File uploader
    uploaded_file = st.file_uploader(
        "📁 Upload Excel file with NCB data and 'Col Ref' sheet",
        type=['xlsx', 'xls'],
        help="File should contain data sheet and 'Col Ref' sheet for column mapping"
    )
    
    if uploaded_file is not None:
        if st.button("🚀 Process Data (Version 2.0)", type="primary", use_container_width=True):
            
            with st.spinner("🔍 Analyzing file structure..."):
                # Detect column structure
                column_mapping, label_columns, amount_columns = detect_column_structure_v2(uploaded_file)
                
                if not column_mapping:
                    st.error("❌ Could not detect column structure. Please ensure file has 'Col Ref' sheet.")
                    return
            
            with st.spinner("📊 Processing data..."):
                # Read the main data sheet
                try:
                    excel_data = pd.ExcelFile(uploaded_file)
                    
                    # Find the main data sheet (not Col Ref)
                    data_sheet = None
                    for sheet in excel_data.sheet_names:
                        if sheet != 'Col Ref' and not any(keyword in sheet.lower() for keyword in ['summary', 'xref', 'ref']):
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
                    output_df = process_data_v2(df, column_mapping, label_columns, amount_columns)
                    
                    if output_df is not None and len(output_df) > 0:
                        st.success(f"✅ **Processing Complete!** Generated {len(output_df)} records")
                        
                        # Display results
                        st.subheader("📊 Processing Results")
                        st.write(f"**Total Records:** {len(output_df)}")
                        
                        # Show breakdown by transaction type
                        type_counts = output_df['Transaction_Type'].value_counts()
                        st.write("**Breakdown by Transaction Type:**")
                        for trans_type, count in type_counts.items():
                            st.write(f"  {trans_type}: {count} records")
                        
                        # Show sample data
                        st.subheader("🔍 Sample Output Data")
                        st.dataframe(output_df.head(10))
                        
                        # Download button
                        st.subheader("💾 Download Results")
                        
                        # Create Excel file in memory
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                            output_df.to_excel(writer, sheet_name='Data', index=False)
                        
                        output_buffer.seek(0)
                        
                        # Generate filename with timestamp
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"Karen_NCB_2_0_Output_{timestamp}.xlsx"
                        
                        st.download_button(
                            label="📥 Download Excel File",
                            data=output_buffer.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                        st.write(f"📁 **File saved as:** {filename}")
                        
                    else:
                        st.error("❌ No data was processed successfully")
                        
                except Exception as e:
                    st.error(f"❌ Error processing file: {e}")
                    import traceback
                    st.code(traceback.format_exc())

if __name__ == "__main__":
    main()
