#!/usr/bin/env python3
"""
Karen 2.0 NCB Data Processor
NCB Transaction Summarization with specific column requirements
"""

import streamlit as st
import pandas as pd
from datetime import datetime
import io

st.set_page_config(page_title="Karen 2.0 NCB Data Processor", page_icon="📊", layout="wide")

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

def find_ncb_columns(df, column_mapping):
    """Find NCB-related columns based on the Karen 2.0 specifications."""
    ncb_columns = {}
    
    # NCB-related labels to match against
    ncb_labels = [
        'Agent NCB', 'Agent NCB Fee', 'Dealer NCB', 'Dealer NCB Fee',
        'Agent NCB Offset', 'Agent NCB Offset Bucket', 'Dealer NCB Offset', 'Dealer NCB Offset Bucket'
    ]
    
    # Look for NCB columns based on the mapping
    for col_idx, desc in column_mapping.items():
        if col_idx < len(df.columns):
            col_name = df.columns[col_idx]
            
            # Check if this column description matches any NCB label
            for ncb_label in ncb_labels:
                if ncb_label.lower() in desc.lower():
                    # This is an NCB column, find the amount column (next cell over)
                    if col_idx + 1 < len(df.columns):
                        amount_col = df.columns[col_idx + 1]
                        ncb_columns[ncb_label] = amount_col
                        st.write(f"✅ **Found NCB column:** {ncb_label} → {amount_col}")
                    break
    
    # If we don't have enough NCB columns, try to find them by content analysis
    if len(ncb_columns) < 7:  # We need 7 NCB columns
        with st.expander("⚠️ **NCB Column Detection Details** (Click to expand)", expanded=False):
            st.warning("⚠️ Could not identify all NCB columns from mapping, trying content-based detection...")
            
            # Look for columns that might contain NCB amounts by analyzing their content
            potential_ncb_cols = []
            
            for col in df.columns:
                try:
                    # Skip datetime columns
                    if isinstance(col, pd.Timestamp) or 'datetime' in str(type(col)).lower():
                        continue
                        
                    # Try to convert to numeric
                    numeric_values = pd.to_numeric(df[col], errors='coerce')
                    if not numeric_values.isna().all() and numeric_values.dtype in ['int64', 'float64']:
                        # Check if this column has meaningful non-zero values
                        non_zero_count = (numeric_values != 0).sum()
                        total_count = len(numeric_values)
                        
                        # Only consider columns with a reasonable number of non-zero values
                        if non_zero_count > 0 and non_zero_count < total_count * 0.9:  # Not all zeros, not all non-zero
                            potential_ncb_cols.append({
                                'column': col,
                                'non_zero_count': non_zero_count,
                                'total_count': total_count,
                                'non_zero_ratio': non_zero_count / total_count
                            })
                except:
                    pass
            
            # Sort by non-zero ratio to find the most likely NCB columns
            potential_ncb_cols.sort(key=lambda x: x['non_zero_ratio'], reverse=True)
            
            st.write(f"🔍 **Potential NCB columns found:** {len(potential_ncb_cols)}")
            for i, col_info in enumerate(potential_ncb_cols[:10]):  # Show top 10
                st.write(f"  {i+1}. {col_info['column']}: {col_info['non_zero_count']}/{col_info['total_count']} non-zero ({col_info['non_zero_ratio']:.1%})")
            
            # Assign the top 7 columns as NCB columns
            if len(potential_ncb_cols) >= 7:
                ncb_columns = {
                    'Agent NCB Fee': potential_ncb_cols[0]['column'],
                    'Dealer NCB Fee': potential_ncb_cols[1]['column'],
                    'Agent NCB Offset': potential_ncb_cols[2]['column'],
                    'Agent NCB Offset Bucket': potential_ncb_cols[3]['column'],
                    'Dealer NCB Offset Bucket': potential_ncb_cols[4]['column'],
                    'Agent NCB Offset': potential_ncb_cols[5]['column'],
                    'Dealer NCB Offset Bucket': potential_ncb_cols[6]['column']
                }
                
                st.write(f"✅ **NCB columns assigned by content analysis:**")
                for ncb_type, col_name in ncb_columns.items():
                    st.write(f"  {ncb_type}: {col_name}")
            else:
                st.error(f"❌ Not enough potential NCB columns found. Need 7, found {len(potential_ncb_cols)}")
    
    return ncb_columns

def find_required_columns(df):
    """Find the required columns for each dataset based on Karen 2.0 specifications."""
    required_cols = {}
    
    # Data Set 1 & 2 (NB & R) - Common columns
    common_cols = [
        'B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U',  # Basic info
        'AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC'  # NCB amounts
    ]
    
    # Data Set 3 (C) - Additional columns
    cancellation_cols = common_cols + ['Z', 'AE', 'AB', 'AA']
    
    # Try to find columns by position or name
    for col in df.columns:
        col_str = str(col)
        
        # Map column positions to names
        if col_str == 'B' or 'insurer' in col_str.lower():
            required_cols['B'] = col
        elif col_str == 'C' or 'product' in col_str.lower():
            required_cols['C'] = col
        elif col_str == 'D' or 'coverage' in col_str.lower():
            required_cols['D'] = col
        elif col_str == 'E' or 'dealer' in col_str.lower() and 'number' in col_str.lower():
            required_cols['E'] = col
        elif col_str == 'F' or 'dealer' in col_str.lower() and 'name' in col_str.lower():
            required_cols['F'] = col
        elif col_str == 'H' or 'contract' in col_str.lower() and 'number' in col_str.lower():
            required_cols['H'] = col
        elif col_str == 'L' or 'sale' in col_str.lower() and 'date' in col_str.lower():
            required_cols['L'] = col
        elif col_str == 'J' or 'transaction' in col_str.lower() and 'date' in col_str.lower():
            required_cols['J'] = col
        elif col_str == 'M' or 'transaction' in col_str.lower() and 'type' in col_str.lower():
            required_cols['M'] = col
        elif col_str == 'U' or 'last' in col_str.lower() and 'name' in col_str.lower():
            required_cols['U'] = col
        elif col_str == 'Z' or 'term' in col_str.lower():
            required_cols['Z'] = col
        elif col_str == 'AE' or 'cancellation' in col_str.lower() and 'date' in col_str.lower():
            required_cols['AE'] = col
        elif col_str == 'AB' or 'reason' in col_str.lower():
            required_cols['AB'] = col
        elif col_str == 'AA' or 'factor' in col_str.lower():
            required_cols['AA'] = col
    
    # If we can't find columns by name, try to map by position
    if len(required_cols) < len(common_cols):
        st.warning("⚠️ Could not identify all required columns by name, using position-based mapping...")
        
        # Map by position (assuming first few columns are in order)
        for i, col in enumerate(df.columns[:len(common_cols)]):
            if i < len(common_cols):
                col_letter = chr(66 + i)  # Start from 'B'
                required_cols[col_letter] = col
    
    return required_cols

def process_excel_data_karen_2_0(uploaded_file):
    """Process Excel data using smart column detection for Karen 2.0."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Create expandable section for processing details
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen 2.0 Processing Started**")
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
                
                # Find NCB columns
                ncb_columns = find_ncb_columns(df, column_mapping)
                st.write("🔍 **NCB Columns Found:**")
                for ncb_type, col_name in ncb_columns.items():
                    st.write(f"  {ncb_type}: {col_name}")
                
                # Find required columns
                required_cols = find_required_columns(df)
                st.write("🔍 **Required Columns Found:**")
                for col_letter, col_name in required_cols.items():
                    st.write(f"  {col_letter}: {col_name}")
                
                if len(ncb_columns) >= 7 and len(required_cols) >= len(['B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U']):
                    # Process the data
                    results = process_transaction_data_karen_2_0(df, ncb_columns, required_cols)
                    
                    if results:
                        results['column_mapping'] = column_mapping
                        results['ncb_columns'] = ncb_columns
                        results['required_cols'] = required_cols
                        
                        return results
                    else:
                        return None
                else:
                    st.error(f"❌ Need at least 7 NCB columns and 10 required columns, found {len(ncb_columns)} NCB and {len(required_cols)} required")
                    return None
            else:
                st.error("❌ No suitable data sheet found")
                return None
                
    except Exception as e:
        st.error(f"❌ Error in Karen 2.0 processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def process_transaction_data_karen_2_0(df, ncb_columns, required_cols):
    """Process transaction data according to Karen 2.0 specifications."""
    try:
        # Get the NCB column names
        ncb_cols = list(ncb_columns.values())
        
        if len(ncb_cols) < 7:
            st.error(f"❌ Need exactly 7 NCB columns, found {len(ncb_cols)}")
            return None
        
        st.write(f"✅ **Processing with NCB columns:** {ncb_cols}")
        
        # Convert NCB columns to numeric
        for col in ncb_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            else:
                st.error(f"❌ Column {col} not found in data")
                return None
        
        # Calculate sum of NCB amounts
        df['NCB_Sum'] = df[ncb_cols].sum(axis=1)
        
        # Find transaction type column
        transaction_col = required_cols.get('M')
        if not transaction_col:
            st.error("❌ No transaction type column found")
            return None
        
        st.write(f"✅ **Transaction column found:** {transaction_col}")
        
        # Show unique values in transaction column
        unique_transactions = df[transaction_col].astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")
        
        # Filter by transaction type and NCB amounts according to Karen 2.0 rules
        nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
        c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
        r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        
        # Apply Karen 2.0 filtering rules
        # Data Set 1 (NB): sum > 0
        nb_df = df[nb_mask].copy()
        nb_filtered = nb_df[nb_df['NCB_Sum'] > 0]
        
        # Data Set 2 (R): sum > 0
        r_df = df[r_mask].copy()
        r_filtered = r_df[r_df['NCB_Sum'] > 0]
        
        # Data Set 3 (C): sum < 0
        c_df = df[c_mask].copy()
        c_filtered = c_df[c_df['NCB_Sum'] < 0]
        
        st.write(f"📊 **Filtered results (Karen 2.0 rules):**")
        st.write(f"  New Business (sum > 0): {len(nb_filtered)}")
        st.write(f"  Reinstatements (sum > 0): {len(r_filtered)}")
        st.write(f"  Cancellations (sum < 0): {len(c_filtered)}")
        
        # Create output dataframes with required columns in correct order
        # Data Set 1: New Business (NB)
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), ncb_columns.get('Agent NCB Fee'),
            ncb_columns.get('Dealer NCB Fee'), ncb_columns.get('Agent NCB Offset'),
            ncb_columns.get('Agent NCB Offset Bucket'), ncb_columns.get('Dealer NCB Offset Bucket'),
            ncb_columns.get('Agent NCB Offset'), ncb_columns.get('Dealer NCB Offset Bucket')
        ]
        
        # Data Set 2: Reinstatements (R) - same columns as NB
        r_output_cols = nb_output_cols.copy()
        
        # Data Set 3: Cancellations (C) - additional columns
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), required_cols.get('Z'), required_cols.get('AE'),
            required_cols.get('AB'), required_cols.get('AA'), ncb_columns.get('Agent NCB Fee'),
            ncb_columns.get('Dealer NCB Fee'), ncb_columns.get('Agent NCB Offset'),
            ncb_columns.get('Agent NCB Offset Bucket'), ncb_columns.get('Dealer NCB Offset Bucket'),
            ncb_columns.get('Agent NCB Offset'), ncb_columns.get('Dealer NCB Offset Bucket')
        ]
        
        # Filter dataframes to only include available columns
        nb_output_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        r_output_cols = [col for col in r_output_cols if col is not None and col in df.columns]
        c_output_cols = [col for col in c_output_cols if col is not None and col in df.columns]
        
        # Create output dataframes
        nb_output = nb_filtered[nb_output_cols].copy()
        r_output = r_filtered[r_output_cols].copy()
        c_output = c_filtered[c_output_cols].copy()
        
        # Set column headers based on Karen 2.0 specifications
        nb_headers = [
            'Insurer Code', 'Product Type Code', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Customer Last Name',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
            'Admin 8 Amount (Dealer NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)'
        ]
        
        r_headers = nb_headers.copy()
        
        c_headers = [
            'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Last Name',
            'Contract Term', 'Cancellation Date', 'Cancellation Reason', 'Cancellation Factor',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
            'Admin 8 Amount (Dealer NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)'
        ]
        
        # Apply headers
        if len(nb_headers) == len(nb_output.columns):
            nb_output.columns = nb_headers
        if len(r_headers) == len(r_output.columns):
            r_output.columns = r_headers
        if len(c_headers) == len(c_output.columns):
            c_output.columns = c_headers
        
        st.write(f"🔍 **Output dataframe shapes:**")
        st.write(f"  NB: {nb_output.shape}")
        st.write(f"  R: {r_output.shape}")
        st.write(f"  C: {c_output.shape}")
        
        return {
            'nb_data': nb_output,
            'reinstatement_data': r_output,
            'cancellation_data': c_output,
            'ncb_columns': ncb_columns,
            'total_records': len(nb_output) + len(r_output) + len(c_output)
        }
        
    except Exception as e:
        st.error(f"❌ Error processing transaction data: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def create_excel_download_karen_2_0(results):
    """Create Excel download with three separate worksheets as per Karen 2.0."""
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Data Set 1: New Business (NB)
        if len(results['nb_data']) > 0:
            results['nb_data'].to_excel(writer, sheet_name='Data Set 1 - New Business', index=False)
        
        # Data Set 2: Reinstatements (R)
        if len(results['reinstatement_data']) > 0:
            results['reinstatement_data'].to_excel(writer, sheet_name='Data Set 2 - Reinstatements', index=False)
        
        # Data Set 3: Cancellations (C)
        if len(results['cancellation_data']) > 0:
            results['cancellation_data'].to_excel(writer, sheet_name='Data Set 3 - Cancellations', index=False)
    
    return output.getvalue()

def main():
    st.title("📊 Karen 2.0 NCB Data Processor")
    st.markdown("**NCB Transaction Summarization with specific column requirements**")
    
    # Sidebar
    with st.sidebar:
        st.header("🔧 Karen 2.0 Processing")
        st.markdown("**This app will:**")
        st.markdown("• Extract NB, R, and C transaction types")
        st.markdown("• Apply Karen 2.0 filtering rules")
        st.markdown("• Create 3 separate worksheets")
        st.markdown("• Output 2,000-2,500 rows total")
        st.markdown("• Preserve exact column order")
    
    # File upload
    uploaded_file = st.file_uploader("Choose Excel file for processing", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        st.success(f"✅ File uploaded: {uploaded_file.name}")
        st.write(f"📁 Size: {uploaded_file.size / 1024:.1f} KB")
        
        if st.button("📊 Start Karen 2.0 Processing", type="primary", use_container_width=True):
            with st.spinner("Karen 2.0 processing in progress..."):
                results = process_excel_data_karen_2_0(uploaded_file)
                
                if results:
                    st.success("✅ Karen 2.0 processing complete!")
                    
                    # Show results
                    st.header("📊 Processing Results - Karen 2.0")
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("New Business", len(results['nb_data']))
                    with col2:
                        st.metric("Reinstatements", len(results['reinstatement_data']))
                    with col3:
                        st.metric("Cancellations", len(results['cancellation_data']))
                    with col4:
                        st.metric("Total Records", results['total_records'])
                    
                    # Show NCB columns used
                    st.subheader("🔍 NCB Columns Used")
                    for ncb_type, col_name in results['ncb_columns'].items():
                        st.write(f"  **{ncb_type}:** {col_name}")
                    
                    # Show expected output information
                    st.subheader("📋 Expected Output - Karen 2.0")
                    st.write("**Data Set 1 - New Business (NB):** New contracts with NCB sum > 0")
                    st.write("**Data Set 2 - Reinstatements (R):** Reinstated records with NCB sum > 0")
                    st.write("**Data Set 3 - Cancellations (C):** Cancelled records with NCB sum < 0")
                    st.write(f"**Total expected rows:** 2,000-2,500 (Current: {results['total_records']})")
                    
                    # Show sample data
                    if len(results['nb_data']) > 0:
                        st.subheader("🆕 Data Set 1 - New Business (NB)")
                        st.dataframe(results['nb_data'].head(10))
                    
                    if len(results['reinstatement_data']) > 0:
                        st.subheader("🔄 Data Set 2 - Reinstatements (R)")
                        st.dataframe(results['reinstatement_data'].head(10))
                    
                    if len(results['cancellation_data']) > 0:
                        st.subheader("❌ Data Set 3 - Cancellations (C)")
                        st.dataframe(results['cancellation_data'].head(10))
                    
                    # Download button for all three datasets
                    st.download_button(
                        label="📥 Download All Datasets (Karen 2.0 Format)",
                        data=create_excel_download_karen_2_0(results),
                        file_name=f"NCB_Karen_2_0_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Individual download buttons
                    if len(results['nb_data']) > 0:
                        st.download_button(
                            label="📥 Download New Business Data",
                            data=create_excel_download_karen_2_0({'nb_data': results['nb_data'], 'reinstatement_data': pd.DataFrame(), 'cancellation_data': pd.DataFrame()}),
                            file_name=f"Data_Set_1_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['reinstatement_data']) > 0:
                        st.download_button(
                            label="📥 Download Reinstatement Data",
                            data=create_excel_download_karen_2_0({'nb_data': pd.DataFrame(), 'reinstatement_data': results['reinstatement_data'], 'cancellation_data': pd.DataFrame()}),
                            file_name=f"Data_Set_2_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    if len(results['cancellation_data']) > 0:
                        st.download_button(
                            label="📥 Download Cancellation Data",
                            data=create_excel_download_karen_2_0({'nb_data': pd.DataFrame(), 'reinstatement_data': pd.DataFrame(), 'cancellation_data': results['cancellation_data']}),
                            file_name=f"Data_Set_3_Cancellations_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("❌ Karen 2.0 processing failed. Check the error messages above.")

if __name__ == "__main__":
    main()
