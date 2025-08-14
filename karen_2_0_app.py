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

def process_excel_data_karen_2_0(uploaded_file):
    """Process Excel data from the Data tab for Karen 2.0."""
    try:
        excel_data = pd.ExcelFile(uploaded_file)
        
        # Create expandable section for processing details
        with st.expander("🔍 **Processing Details** (Click to expand)", expanded=False):
            st.write("🔍 **Karen 2.0 Processing Started**")
            st.write(f"📊 Sheets found: {excel_data.sheet_names}")
            
            # Process ONLY the Data sheet (main transaction data)
            data_sheet_name = 'Data'
            
            if data_sheet_name in excel_data.sheet_names:
                st.write(f"📋 **Processing {data_sheet_name} Sheet**")
                
                # Read the Data sheet - the first row contains the actual headers
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name, header=None)
                st.write(f"📏 Data shape: {df.shape}")
                
                # The first row (index 0) contains the actual column headers
                # Let's use that row as our column names, but handle nan values
                header_row = df.iloc[0]
                st.write(f"🔍 **Header row sample:** {header_row[:10].tolist()}")
                
                # Clean up the header row - replace nan with meaningful names
                clean_headers = []
                for i, header in enumerate(header_row):
                    if pd.isna(header) or str(header).strip() == '':
                        clean_headers.append(f'Column_{i}')
                    else:
                        clean_headers.append(str(header).strip())
                
                df.columns = clean_headers
                df = df.iloc[1:].reset_index(drop=True)
                st.write(f"📏 Data shape after header fix: {df.shape}")
                
                # Show first few column names to understand structure
                st.write(f"🔍 **First 20 columns:** {list(df.columns[:20])}")
                
                # Find NCB columns by looking for Admin columns (AO, AQ, AU, AW, AY, BA, BC)
                ncb_columns = find_ncb_columns_simple(df)
                st.write("🔍 **NCB Columns Found:**")
                for ncb_type, col_name in ncb_columns.items():
                    st.write(f"  {ncb_type}: {col_name}")
                
                # Find required columns by position mapping
                required_cols = find_required_columns_simple(df)
                st.write("🔍 **Required Columns Found:**")
                for col_letter, col_name in required_cols.items():
                    st.write(f"  {col_letter}: {col_name}")
                
                if len(ncb_columns) >= 7 and len(required_cols) >= len(['B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U']):
                    # Process the data
                    results = process_transaction_data_karen_2_0(df, ncb_columns, required_cols)
                    
                    if results:
                        results['ncb_columns'] = ncb_columns
                        results['required_cols'] = required_cols
                        
                        return results
                    else:
                        return None
                else:
                    st.error(f"❌ Need at least 7 NCB columns and 10 required columns, found {len(ncb_columns)} NCB and {len(required_cols)} required")
                    return None
            else:
                st.error(f"❌ '{data_sheet_name}' sheet not found. Available sheets: {excel_data.sheet_names}")
                return None
                
    except Exception as e:
        st.error(f"❌ Error in Karen 2.0 processing: {e}")
        import traceback
        st.code(traceback.format_exc())
        return None

def examine_data_structure(df):
    """Examine the data structure to understand what we're working with."""
    st.write("🔍 **Data Structure Analysis:**")
    
    # Show data types
    st.write(f"📊 **Data types:**")
    for i, col in enumerate(df.columns[:20]):
        try:
            sample_values = df.iloc[:5, i].dropna()
            if len(sample_values) > 0:
                st.write(f"  Column {i} ({col}): {sample_values.iloc[0]} (type: {type(sample_values.iloc[0])})")
            else:
                st.write(f"  Column {i} ({col}): No data")
        except:
            st.write(f"  Column {i} ({col}): Error reading")
    
    # Look for any columns that might contain transaction types
    st.write(f"🔍 **Looking for transaction type columns...**")
    for i, col in enumerate(df.columns):
        try:
            values = df.iloc[:20, i].dropna()
            if len(values) > 0:
                # Check if this looks like transaction types
                str_values = [str(v).upper() for v in values]
                if any(t in str_values for t in ['NB', 'C', 'R', 'NEW', 'CANCEL', 'REINSTATE']):
                    st.write(f"  ✅ **Potential transaction column:** {col} (Position {i})")
                    st.write(f"     Sample values: {str_values[:10]}")
        except:
            continue
    
    # Look for any numeric columns that might be amounts
    st.write(f"🔍 **Looking for numeric/amount columns...**")
    numeric_columns = []
    for i, col in enumerate(df.columns):
        try:
            values = df.iloc[:50, i].dropna()
            if len(values) > 0:
                numeric_vals = pd.to_numeric(values, errors='coerce')
                non_null_numeric = numeric_vals.dropna()
                if len(non_null_numeric) > 0:
                    non_zero = (non_null_numeric != 0).sum()
                    if non_zero > 0:
                        numeric_columns.append((i, col, non_zero, len(non_null_numeric)))
        except:
            continue
    
    # Sort by number of non-zero values
    numeric_columns.sort(key=lambda x: x[2], reverse=True)
    
    st.write(f"🔍 **Found {len(numeric_columns)} numeric columns with non-zero values:**")
    for i, col, non_zero, total in numeric_columns[:15]:
        st.write(f"  Position {i}: {col} - {non_zero}/{total} non-zero values")
    
    return numeric_columns

def find_ncb_columns_simple(df):
    """Find NCB columns by looking for Admin columns based on actual file structure."""
    ncb_columns = {}
    
    # First, let's examine the data structure
    numeric_columns = examine_data_structure(df)
    
    # First, let's show all columns that might be Admin/Amount related for debugging
    st.write("🔍 **Searching for Admin/Amount columns...**")
    admin_candidates = []
    for i, col in enumerate(df.columns):
        col_str = str(col).upper()
        if any(keyword in col_str for keyword in ['ADMIN', 'AMOUNT', 'NCB', 'FEE', 'OFFSET']):
            admin_candidates.append((i, col, col_str))
    
    st.write(f"🔍 **Found {len(admin_candidates)} potential Admin/Amount columns:**")
    for pos, col, col_str in admin_candidates[:15]:  # Show first 15
        st.write(f"  Position {pos}: {col} (Content: {col_str})")
    
    # Let's also show all column names to understand the structure
    st.write("🔍 **All column names (first 30):**")
    for i, col in enumerate(df.columns[:30]):
        st.write(f"  Position {i}: {col}")
    
    # Let's be more flexible and find Admin columns by content, not just position
    # Look for columns that contain "ADMIN" and "Amount" in their names
    for i, col in enumerate(df.columns):
        col_str = str(col).upper()
        if 'ADMIN' in col_str and 'AMOUNT' in col_str:
            # Check if this column has numeric values
            try:
                # Check for numeric values
                values = df.iloc[:, i].dropna()
                if len(values) > 0:
                    numeric_vals = pd.to_numeric(values, errors='coerce')
                    non_zero = (numeric_vals != 0).sum()
                    if non_zero > 0:  # Has some non-zero values
                        # Map to the required NCB column types
                        if len(ncb_columns) == 0:
                            ncb_columns['AO'] = col  # Admin 3 Amount (Agent NCB Fee)
                        elif len(ncb_columns) == 1:
                            ncb_columns['AQ'] = col  # Admin 4 Amount (Dealer NCB Fee)
                        elif len(ncb_columns) == 2:
                            ncb_columns['AU'] = col  # Admin 6 Amount (Agent NCB Offset)
                        elif len(ncb_columns) == 3:
                            ncb_columns['AW'] = col  # Admin 7 Amount (Agent NCB Offset Bucket)
                        elif len(ncb_columns) == 4:
                            ncb_columns['AY'] = col  # Admin 8 Amount (Dealer NCB Offset Bucket)
                        elif len(ncb_columns) == 5:
                            ncb_columns['BA'] = col  # Admin 9 Amount (Agent NCB Offset)
                        elif len(ncb_columns) == 6:
                            ncb_columns['BC'] = col  # Admin 10 Amount (Dealer NCB Offset Bucket)
                        
                        st.write(f"✅ **Found NCB column:** {list(ncb_columns.keys())[-1]} → {col} (Position {i})")
                        
                        if len(ncb_columns) >= 7:
                            break
            except Exception as e:
                st.write(f"⚠️ **Skipping column {col} due to error:** {e}")
                continue
    
    # If we still don't have enough, try to find additional Admin columns by position
    if len(ncb_columns) < 7:
        st.warning(f"⚠️ **Only found {len(ncb_columns)} NCB columns, need 7. Trying position-based mapping...**")
        
        # Try some common Admin column positions
        potential_positions = [40, 42, 46, 48, 50, 52, 54, 56, 58, 60, 62, 64, 66, 68, 70]
        
        for pos in potential_positions:
            if pos < len(df.columns) and len(ncb_columns) < 7:
                col = df.columns[pos]
                col_str = str(col).upper()
                
                # Check if this looks like an Admin column
                if 'ADMIN' in col_str or 'AMOUNT' in col_str or 'NCB' in col_str:
                    # Map to the next available NCB column type
                    if 'AO' not in ncb_columns:
                        ncb_columns['AO'] = col
                    elif 'AQ' not in ncb_columns:
                        ncb_columns['AQ'] = col
                    elif 'AU' not in ncb_columns:
                        ncb_columns['AU'] = col
                    elif 'AW' not in ncb_columns:
                        ncb_columns['AW'] = col
                    elif 'AY' not in ncb_columns:
                        ncb_columns['AY'] = col
                    elif 'BA' not in ncb_columns:
                        ncb_columns['BA'] = col
                    elif 'BC' not in ncb_columns:
                        ncb_columns['BC'] = col
                    
                    st.write(f"✅ **Found NCB column by position:** {list(ncb_columns.keys())[-1]} → {col} (Position {pos})")
    
    # If we still don't have enough, let's try to find any numeric columns that might be Admin amounts
    if len(ncb_columns) < 7:
        st.warning(f"⚠️ **Still only found {len(ncb_columns)} NCB columns. Looking for any numeric columns...**")
        
        # Use the numeric columns we found earlier
        for i, col, non_zero, total in numeric_columns:
            if len(ncb_columns) >= 7:
                break
                
            # Map to the next available NCB column type
            if 'AO' not in ncb_columns:
                ncb_columns['AO'] = col
            elif 'AQ' not in ncb_columns:
                ncb_columns['AQ'] = col
            elif 'AU' not in ncb_columns:
                ncb_columns['AU'] = col
            elif 'AW' not in ncb_columns:
                ncb_columns['AW'] = col
            elif 'AY' not in ncb_columns:
                ncb_columns['AY'] = col
            elif 'BA' not in ncb_columns:
                ncb_columns['BA'] = col
            elif 'BC' not in ncb_columns:
                ncb_columns['BC'] = col
            
            st.write(f"✅ **Found potential NCB column by numeric analysis:** {list(ncb_columns.keys())[-1]} → {col} (Position {i}) - {non_zero}/{total} non-zero")
    
    return ncb_columns

def find_required_columns_simple(df):
    """Find required columns by simple position mapping based on actual file structure."""
    required_cols = {}
    
    # Let's be more flexible and find columns by content, not just fixed positions
    for i, col in enumerate(df.columns):
        col_str = str(col).upper()
        
        # Map columns based on content and common patterns
        if 'INSURER' in col_str or 'INSURANCE' in col_str:
            required_cols['B'] = col
        elif 'PRODUCT' in col_str and 'TYPE' in col_str:
            required_cols['C'] = col
        elif 'COVERAGE' in col_str:
            required_cols['D'] = col
        elif 'DEALER' in col_str and 'NUMBER' in col_str:
            required_cols['E'] = col
        elif 'DEALER' in col_str and 'NAME' in col_str:
            required_cols['F'] = col
        elif 'CONTRACT' in col_str and 'NUMBER' in col_str:
            required_cols['H'] = col
        elif 'SALE' in col_str and 'DATE' in col_str:
            required_cols['L'] = col
        elif 'TRANSACTION' in col_str and 'DATE' in col_str:
            required_cols['J'] = col
        elif 'TRANSACTION' in col_str and 'TYPE' in col_str:
            required_cols['M'] = col
        elif 'CUSTOMER' in col_str and 'LAST' in col_str:
            required_cols['U'] = col
        elif 'TERM' in col_str:
            required_cols['Z'] = col
        elif 'CANCELLATION' in col_str and 'FACTOR' in col_str:
            required_cols['AA'] = col
        elif 'CANCELLATION' in col_str and 'REASON' in col_str:
            required_cols['AB'] = col
        elif 'CANCELLATION' in col_str and 'DATE' in col_str:
            required_cols['AE'] = col
    
    # If we still don't have enough, try position-based mapping as fallback
    if len(required_cols) < 10:
        st.warning(f"⚠️ **Only found {len(required_cols)} required columns, trying position-based mapping...**")
        
        # Fallback to position-based mapping
        position_mapping = {
            1: 'B',   # Insurer Code - Column 1
            2: 'C',   # Product Type Code - Column 2
            3: 'D',   # Coverage Code - Column 3
            4: 'E',   # Dealer Number - Column 4
            5: 'F',   # Dealer Name - Column 5
            6: 'H',   # Contract Number - Column 6
            7: 'L',   # Contract Sale Date - Column 7
            8: 'J',   # Transaction Date - Column 8
            9: 'M',   # Transaction Type - Column 9
            10: 'U',  # Customer Last Name - Column 10
            20: 'Z',  # Contract Term - Column 20
            26: 'AA', # Cancellation Factor - Column 26
            27: 'AB', # Cancellation Reason - Column 27
            28: 'AE', # Cancellation Date - Column 28
        }
        
        for pos, col_letter in position_mapping.items():
            if pos < len(df.columns) and col_letter not in required_cols:
                col = df.columns[pos]
                required_cols[col_letter] = col
                st.write(f"✅ **Found required column by position:** {col_letter} → {col} (Position {pos})")
    
    return required_cols

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
        transaction_col = required_cols.get('M') # Transaction Type column
        if not transaction_col:
            st.error("❌ No transaction type column found")
            return None
        
        st.write(f"✅ **Transaction column found:** {transaction_col}")
        
        # Show unique values in transaction column
        unique_transactions = df[transaction_col].astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")
        
        # More flexible transaction type filtering - look for variations
        # Filter by transaction type and NCB amounts according to Karen 2.0 rules
        nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW', 'NEW BUSINESS', 'N'])
        c_mask = df[transaction_col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL', 'CANCELLED', 'CANC'])
        r_mask = df[transaction_col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE', 'REINSTATED'])
        
        # Also check for any other transaction types that might exist
        other_mask = ~(nb_mask | c_mask | r_mask)
        other_transactions = df[other_mask][transaction_col].astype(str).unique()
        if len(other_transactions) > 0:
            st.write(f"🔍 **Other transaction types found:** {other_transactions[:10]}")
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        st.write(f"  Other/Unknown: {other_mask.sum()}")
        
        # Show sample values from each transaction type
        if nb_mask.sum() > 0:
            st.write(f"🔍 **Sample NB transactions:** {df[nb_mask][transaction_col].astype(str).unique()[:5]}")
        if c_mask.sum() > 0:
            st.write(f"🔍 **Sample C transactions:** {df[c_mask][transaction_col].astype(str).unique()[:5]}")
        if r_mask.sum() > 0:
            st.write(f"🔍 **Sample R transactions:** {df[r_mask][transaction_col].astype(str).unique()[:5]}")
        
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
        
        # Debug: Show NCB sum distribution for each transaction type
        st.write(f"🔍 **NCB Sum Analysis:**")
        if len(nb_df) > 0:
            st.write(f"  NB NCB Sum range: {nb_df['NCB_Sum'].min():.2f} to {nb_df['NCB_Sum'].max():.2f}")
            st.write(f"  NB NCB Sum > 0: {(nb_df['NCB_Sum'] > 0).sum()}")
        if len(r_df) > 0:
            st.write(f"  R NCB Sum range: {r_df['NCB_Sum'].min():.2f} to {r_df['NCB_Sum'].max():.2f}")
            st.write(f"  R NCB Sum > 0: {(r_df['NCB_Sum'] > 0).sum()}")
        if len(c_df) > 0:
            st.write(f"  C NCB Sum range: {c_df['NCB_Sum'].min():.2f} to {c_df['NCB_Sum'].max():.2f}")
            st.write(f"  C NCB Sum < 0: {(c_df['NCB_Sum'] < 0).sum()}")
        
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
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY'),
            ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        # Data Set 2: Reinstatements (R) - same columns as NB
        r_output_cols = nb_output_cols.copy()
        
        # Data Set 3: Cancellations (C) - additional columns
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), required_cols.get('Z'), required_cols.get('AA'),
            required_cols.get('AB'), required_cols.get('AE'), ncb_columns.get('AO'),
            ncb_columns.get('AQ'), ncb_columns.get('AU'), ncb_columns.get('AW'),
            ncb_columns.get('AY'), ncb_columns.get('BA'), ncb_columns.get('BC')
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
            'Admin 6 Amount (Agent NCB Offset Bucket)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
            'Admin 8 Amount (Dealer NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)'
        ]
        
        # Data Set 2 (R) - Different headers as per Karen 2.0 specifications
        r_headers = [
            'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Last Name',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset Bucket)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
            'Admin 8 Amount (Dealer NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)'
        ]
        
        # Data Set 3 (C) - Different headers as per Karen 2.0 specifications
        c_headers = [
            'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Last Name',
            'Contract Term', 'Cancellation Date', 'Cancellation Reason', 'Cancellation Factor',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset Bucket)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
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
