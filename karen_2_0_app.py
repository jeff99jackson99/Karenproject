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
                
                # Read the Data sheet - row 12 (index 12) contains the actual headers
                df = pd.read_excel(uploaded_file, sheet_name=data_sheet_name, header=None)
                st.write(f"📏 Data shape: {df.shape}")
                
                # Row 12 (index 12) contains the actual column headers with Admin columns
                header_row = df.iloc[12]
                st.write(f"🔍 **Header row (Row 12):** {header_row[:20].tolist()}")
                
                # Use row 12 as column names
                df.columns = header_row
                # Remove rows 0-12 (they're not data) and reset index
                df = df.iloc[13:].reset_index(drop=True)
                st.write(f"📏 Data shape after header fix: {df.shape}")
                
                # Clean up any duplicate column names in the original data
                df = clean_duplicate_columns(df)
                
                # Clean up any empty/None/NaN column names
                df = clean_empty_column_names(df)
                
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
                    # Run comprehensive data structure check
                    structure_check = check_data_structure(df, ncb_columns, required_cols)
                    
                    # Process the data
                    results = process_transaction_data_karen_2_0(df, ncb_columns, required_cols)
                    
                    if results:
                        results['ncb_columns'] = ncb_columns
                        results['required_cols'] = required_cols
                        results['structure_check'] = structure_check
                        
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

def find_ncb_columns_simple(df):
    """Find NCB columns by looking for Admin columns based on actual file structure."""
    ncb_columns = {}
    
    # Look for Admin columns by their actual names (not positions)
    # The actual file has ADMIN 1-10 Amount, we need to map specific ones to Karen 2.0 requirements
    admin_mapping = {
        'ADMIN 3 Amount': 'AO',  # Agent NCB Fee
        'ADMIN 4 Amount': 'AQ',  # Dealer NCB Fee
        'ADMIN 6 Amount': 'AU',  # Agent NCB Offset
        'ADMIN 7 Amount': 'AW',  # Agent NCB Offset Bucket
        'ADMIN 8 Amount': 'AY',  # Dealer NCB Offset Bucket
        'ADMIN 9 Amount': 'BA',  # Agent NCB Offset
        'ADMIN 10 Amount': 'BC', # Dealer NCB Offset Bucket
    }
    
    # Map Admin columns by their actual names
    for col in df.columns:
        if col in admin_mapping:
            ncb_type = admin_mapping[col]
            ncb_columns[ncb_type] = col
            st.write(f"✅ **Found NCB column:** {ncb_type} → {col}")
    
    # If we still don't have enough, try to find additional Admin columns
    if len(ncb_columns) < 7:
        st.warning(f"⚠️ **Only found {len(ncb_columns)} NCB columns, need 7. Looking for additional Admin columns...**")
        
        # Look for additional Admin columns that might contain NCB amounts
        for col in df.columns:
            try:
                if 'ADMIN' in str(col) and 'Amount' in str(col) and col not in ncb_columns.values():
                    # Check if this column has meaningful values
                    values = df[col].dropna()
                    if len(values) > 0:
                        numeric_vals = pd.to_numeric(values, errors='coerce')
                        non_zero = (numeric_vals != 0).sum()
                        if non_zero > 0 and non_zero < len(values) * 0.9:  # Not all zeros, not all non-zero
                            # Find a unique name for this Admin column
                            admin_num = f"Admin_{len(ncb_columns) + 1}"
                            ncb_columns[admin_num] = col
                            st.write(f"✅ **Found additional Admin column:** {admin_num} → {col}")
                            
                            if len(ncb_columns) >= 7:
                                break
            except:
                pass
    
    return ncb_columns

def find_required_columns_simple(df):
    """Find required columns by label matching as required by Karen 2.0."""
    required_cols = {}
    
    # Use label matching instead of fixed positions as required by Karen 2.0
    # Look for columns that match the required labels
    for i, col in enumerate(df.columns):
        col_str = str(col).upper()
        
        # Map columns based on content and common patterns
        if 'INSURER' in col_str and 'CODE' in col_str:
            required_cols['B'] = col
        elif 'PRODUCT' in col_str and 'TYPE' in col_str and 'CODE' in col_str:
            required_cols['C'] = col
        elif 'COVERAGE' in col_str and 'CODE' in col_str:
            required_cols['D'] = col
        elif 'DEALER' in col_str and 'NUMBER' in col_str:
            required_cols['E'] = col
        elif 'DEALER' in col_str and 'NAME' in col_str:
            required_cols['F'] = col
        elif 'CONTRACT' in col_str and 'NUMBER' in col_str:
            required_cols['H'] = col
        elif 'CONTRACT' in col_str and 'SALE' in col_str and 'DATE' in col_str:
            required_cols['L'] = col
        elif 'TRANSACTION' in col_str and 'DATE' in col_str:
            required_cols['J'] = col
        elif 'TRANSACTION' in col_str and 'TYPE' in col_str:
            required_cols['M'] = col
        elif 'CUSTOMER' in col_str and 'LAST' in col_str and 'NAME' in col_str:
            required_cols['U'] = col
        elif 'TERM' in col_str and 'MONTHS' in col_str:
            required_cols['Z'] = col
        elif 'CANCELLATION' in col_str and 'FACTOR' in col_str:
            required_cols['AA'] = col
        elif 'CANCELLATION' in col_str and 'REASON' in col_str:
            required_cols['AB'] = col
        elif 'CANCELLATION' in col_str and 'DATE' in col_str:
            required_cols['AE'] = col
    
    # If we still don't have enough, try position-based mapping as fallback
    if len(required_cols) < 10:
        st.warning(f"⚠️ **Only found {len(required_cols)} required columns by label matching, trying position-based mapping...**")
        
        # Fallback to position-based mapping
        position_mapping = {
            1: 'B',   # Insurer Code - Column 1
            2: 'C',   # Product Type Code - Column 2
            3: 'D',   # Coverage Code - Column 3
            4: 'E',   # Dealer Number - Column 4
            5: 'F',   # Dealer Name - Column 5
            7: 'H',   # Contract Number - Column 7
            11: 'L',  # Contract Sale Date - Column 11
            8: 'J',   # Transaction Date - Column 8
            9: 'M',   # Transaction Type - Column 9
            12: 'U',  # Customer Last Name - Column 12
            20: 'Z',  # Term Months (Contract Term) - Column 20
            27: 'AA', # Cancellation Factor - Column 27
            30: 'AB', # Cancellation Reason - Column 30
            25: 'AE', # Cancellation Date - Column 25
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
        st.write(f"🔍 **Transaction column type:** {type(df[transaction_col])}")
        st.write(f"🔍 **Transaction column shape:** {df[transaction_col].shape if hasattr(df[transaction_col], 'shape') else 'No shape'}")
        
        # Ensure we're working with a Series, not a DataFrame
        if isinstance(df[transaction_col], pd.DataFrame):
            # If it's a DataFrame, select the first column
            transaction_series = df[transaction_col].iloc[:, 0]
            st.write(f"🔍 **Converted DataFrame to Series:** {type(transaction_series)}")
        else:
            # If it's already a Series, use it directly
            transaction_series = df[transaction_col]
            st.write(f"🔍 **Using Series directly:** {type(transaction_series)}")
        
        # Show unique values in transaction column
        unique_transactions = transaction_series.astype(str).unique()
        st.write(f"🔍 **Unique transaction values found:** {unique_transactions[:20]}")
        
        # More flexible transaction type filtering - look for variations
        # Filter by transaction type and NCB amounts according to Karen 2.0 rules
        nb_mask = transaction_series.astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW', 'N'])
        c_mask = transaction_series.astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL', 'CANCELLED', 'CANC'])
        r_mask = transaction_series.astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE', 'REINSTATED'])
        
        # Also check for any other transaction types that might exist
        other_mask = ~(nb_mask | c_mask | r_mask)
        other_transactions = transaction_series.astype(str)[other_mask].unique()
        if len(other_transactions) > 0:
            st.write(f"🔍 **Other transaction types found:** {other_transactions[:10]}")
        
        st.write(f"📊 **Transaction counts:**")
        st.write(f"  New Business: {nb_mask.sum()}")
        st.write(f"  Cancellations: {c_mask.sum()}")
        st.write(f"  Reinstatements: {r_mask.sum()}")
        st.write(f"  Other/Unknown: {other_mask.sum()}")
        
        # Show sample values from each transaction type
        if nb_mask.sum() > 0:
            sample_nb = transaction_series[nb_mask].astype(str).unique()[:5]
            st.write(f"🔍 **Sample NB transactions:** {sample_nb}")
        if c_mask.sum() > 0:
            sample_c = transaction_series[c_mask].astype(str).unique()[:5]
            st.write(f"🔍 **Sample C transactions:** {sample_c}")
        if r_mask.sum() > 0:
            sample_r = transaction_series[r_mask].astype(str).unique()[:5]
            st.write(f"🔍 **Sample R transactions:** {sample_r}")
        
        # Create filtered dataframes
        nb_df = df[nb_mask].copy()
        c_df = df[c_mask].copy()
        r_df = df[r_mask].copy()
        
        st.write(f"📊 **Filtered dataframe sizes:**")
        st.write(f"  NB DataFrame: {nb_df.shape}")
        st.write(f"  C DataFrame: {c_df.shape}")
        st.write(f"  R DataFrame: {r_df.shape}")
        
        # Calculate NCB sum for each transaction type
        ncb_cols = list(ncb_columns.values())
        
        if len(nb_df) > 0:
            nb_df['NCB_Sum'] = nb_df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
            st.write(f"🔍 **NB NCB sums sample:** {nb_df['NCB_Sum'].head().tolist()}")
        
        if len(c_df) > 0:
            c_df['NCB_Sum'] = c_df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
            st.write(f"🔍 **C NCB sums sample:** {c_df['NCB_Sum'].head().tolist()}")
        
        if len(r_df) > 0:
            r_df['NCB_Sum'] = r_df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
            st.write(f"🔍 **R NCB sums sample:** {r_df['NCB_Sum'].head().tolist()}")
        
        # Apply Karen 2.0 filtering rules
        # Data Set 1 (NB): sum >= 0 (include 0 values)
        nb_filtered = nb_df[nb_df['NCB_Sum'] >= 0]
        
        # Data Set 2 (R): sum >= 0 (include 0 values)
        r_filtered = r_df[r_df['NCB_Sum'] >= 0]
        
        # Data Set 3 (C): sum <= 0 (include 0 values)
        c_filtered = c_df[c_df['NCB_Sum'] <= 0]
        
        st.write(f"📊 **Final filtered results (Karen 2.0 rules):**")
        st.write(f"  New Business (sum >= 0): {len(nb_filtered)}")
        st.write(f"  Reinstatements (sum >= 0): {len(r_filtered)}")
        st.write(f"  Cancellations (sum <= 0): {len(c_filtered)}")
        
        # Show distribution of NCB sums for debugging
        if len(nb_df) > 0:
            st.write(f"🔍 **NB NCB sum distribution:**")
            st.write(f"  Sum > 0: {(nb_df['NCB_Sum'] > 0).sum()}")
            st.write(f"  Sum = 0: {(nb_df['NCB_Sum'] == 0).sum()}")
            st.write(f"  Sum < 0: {(nb_df['NCB_Sum'] < 0).sum()}")
        
        if len(c_df) > 0:
            st.write(f"🔍 **C NCB sum distribution:**")
            st.write(f"  Sum > 0: {(c_df['NCB_Sum'] > 0).sum()}")
            st.write(f"  Sum = 0: {(c_df['NCB_Sum'] == 0).sum()}")
            st.write(f"  Sum < 0: {(c_df['NCB_Sum'] < 0).sum()}")
        
        if len(r_df) > 0:
            st.write(f"🔍 **R NCB sum distribution:**")
            st.write(f"  Sum > 0: {(r_df['NCB_Sum'] > 0).sum()}")
            st.write(f"  Sum = 0: {(r_df['NCB_Sum'] == 0).sum()}")
            st.write(f"  Sum < 0: {(r_df['NCB_Sum'] < 0).sum()}")
        
        # Create output dataframes with required columns in correct order
        # Data Set 1: New Business (NB) - 17 columns in exact order
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY'),
            ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        # Data Set 2: Reinstatements (R) - 17 columns in exact order
        r_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY'),
            ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        # Data Set 3: Cancellations (C) - 21 columns in exact order
        c_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), required_cols.get('Z'), required_cols.get('AA'),
            required_cols.get('AB'), required_cols.get('AE'), ncb_columns.get('AO'),
            ncb_columns.get('AQ'), ncb_columns.get('AU'), ncb_columns.get('AW'),
            ncb_columns.get('AY'), ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        # Debug: Show selected columns for each dataset
        st.write("🔍 **DEBUG: Selected columns for each dataset:**")
        st.write(f"  NB columns: {nb_output_cols}")
        st.write(f"  R columns: {r_output_cols}")
        st.write(f"  C columns: {c_output_cols}")
        
        # Check for None values in column selection
        nb_none_count = sum(1 for col in nb_output_cols if col is None)
        r_none_count = sum(1 for col in r_output_cols if col is None)
        c_none_count = sum(1 for col in c_output_cols if col is None)
        
        if nb_none_count > 0:
            st.warning(f"⚠️ **NB dataset has {nb_none_count} None columns**")
        if r_none_count > 0:
            st.warning(f"⚠️ **R dataset has {r_none_count} None columns**")
        if c_none_count > 0:
            st.warning(f"⚠️ **C dataset has {c_none_count} None columns**")
        
        # Filter dataframes to only include available columns
        nb_output_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        r_output_cols = [col for col in r_output_cols if col is not None and col in df.columns]
        c_output_cols = [col for col in c_output_cols if col is not None and col in df.columns]
        
        # Debug: Show filtered columns
        st.write("🔍 **DEBUG: Filtered columns (None removed):**")
        st.write(f"  NB filtered: {len(nb_output_cols)} columns")
        st.write(f"  R filtered: {len(r_output_cols)} columns")
        st.write(f"  C filtered: {len(c_output_cols)} columns")
        
        # Check for duplicates in column selection
        nb_duplicates = [col for col in set(nb_output_cols) if nb_output_cols.count(col) > 1]
        r_duplicates = [col for col in set(r_output_cols) if r_output_cols.count(col) > 1]
        c_duplicates = [col for col in set(c_output_cols) if c_output_cols.count(col) > 1]
        
        if nb_duplicates:
            st.error(f"❌ **NB dataset has duplicate columns:** {nb_duplicates}")
        if r_duplicates:
            st.error(f"❌ **R dataset has duplicate columns:** {r_duplicates}")
        if c_duplicates:
            st.error(f"❌ **C dataset has duplicate columns:** {c_duplicates}")
        
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
        
        # Data Set 2 (R) - Different headers as per Karen 2.0 specifications
        r_headers = [
            'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
            'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Transaction Type', 'Last Name',
            'Admin 3 Amount (Agent NCB Fee)', 'Admin 4 Amount (Dealer NCB Fee)',
            'Admin 6 Amount (Agent NCB Offset)', 'Admin 7 Amount (Agent NCB Offset Bucket)',
            'Admin 8 Amount (Dealer NCB Offset Bucket)', 'Admin 9 Amount (Agent NCB Offset)',
            'Admin 10 Amount (Dealer NCB Offset Bucket)'
        ]
        
        # Data Set 3 (C) - Different headers as per Karen 2.0 specifications
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
        
        # Fix any duplicate column names
        nb_output = fix_duplicate_columns(nb_output, "New Business (NB)")
        r_output = fix_duplicate_columns(r_output, "Reinstatements (R)")
        c_output = fix_duplicate_columns(c_output, "Cancellations (C)")
        
        # Validate the output dataframes
        validation_results = validate_output_dataframes(nb_output, r_output, c_output, ncb_columns, required_cols)
        
        if validation_results['errors']:
            st.warning("⚠️ **Validation Issues Found:**")
            for error in validation_results['errors']:
                st.write(error)
        
        st.write(f"🔍 **Output dataframe shapes:**")
        st.write(f"  NB: {nb_output.shape}")
        st.write(f"  R: {r_output.shape}")
        st.write(f"  C: {c_output.shape}")
        
        return {
            'nb_data': nb_output,
            'reinstatement_data': r_output,
            'cancellation_data': c_output,
            'validation_results': validation_results,
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

def validate_output_dataframes(nb_output, r_output, c_output, ncb_columns, required_cols):
    """Validate output dataframes for common issues."""
    validation_results = {
        'nb_valid': True,
        'r_valid': True,
        'c_valid': True,
        'errors': []
    }
    
    # Check for duplicate column names
    for name, df, expected_cols in [
        ('New Business (NB)', nb_output, 17),
        ('Reinstatements (R)', r_output, 17),
        ('Cancellations (C)', c_output, 21)
    ]:
        if df is None or df.empty:
            validation_results[f'{name.lower().replace(" ", "_").replace("(", "").replace(")", "")}_valid'] = False
            validation_results['errors'].append(f"❌ {name}: DataFrame is None or empty")
            continue
            
        # Check column count
        if len(df.columns) != expected_cols:
            validation_results['errors'].append(f"⚠️ {name}: Expected {expected_cols} columns, got {len(df.columns)}")
        
        # Check for duplicate column names
        duplicate_cols = df.columns[df.columns.duplicated()].tolist()
        if duplicate_cols:
            validation_results[f'{name.lower().replace(" ", "_").replace("(", "").replace(")", "")}_valid'] = False
            validation_results['errors'].append(f"❌ {name}: Duplicate columns found: {duplicate_cols}")
        
        # Check for None or empty column names
        empty_cols = [col for col in df.columns if col is None or str(col).strip() == '' or str(col) == 'nan']
        if empty_cols:
            validation_results['errors'].append(f"⚠️ {name}: Empty/None column names: {empty_cols}")
        
        # Check for missing required columns
        if name == 'New Business (NB)' or name == 'Reinstatements (R)':
            required_cols_list = ['B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U', 'AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC']
        else:  # Cancellations
            required_cols_list = ['B', 'C', 'D', 'E', 'F', 'H', 'L', 'J', 'M', 'U', 'Z', 'AE', 'AB', 'AA', 'AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC']
        
        missing_cols = []
        for col_letter in required_cols_list:
            if col_letter in ['AO', 'AQ', 'AU', 'AW', 'AY', 'BA', 'BC']:
                col_name = ncb_columns.get(col_letter)
            else:
                col_name = required_cols.get(col_letter)
            
            if col_name and col_name not in df.columns:
                missing_cols.append(f"{col_letter} ({col_name})")
        
        if missing_cols:
            validation_results['errors'].append(f"⚠️ {name}: Missing required columns: {missing_cols}")
    
    return validation_results

def fix_duplicate_columns(df, dataset_name):
    """Fix duplicate column names by adding suffixes."""
    if df is None or df.empty:
        return df
    
    # Check for duplicate columns
    duplicate_cols = df.columns[df.columns.duplicated()].tolist()
    if not duplicate_cols:
        return df
    
    # Create a new DataFrame with unique column names
    new_columns = []
    seen_columns = {}
    
    for col in df.columns:
        if col in seen_columns:
            seen_columns[col] += 1
            new_col = f"{col}_{seen_columns[col]}"
        else:
            seen_columns[col] = 0
            new_col = col
        
        new_columns.append(new_col)
    
    # Create new DataFrame with fixed column names
    fixed_df = df.copy()
    fixed_df.columns = new_columns
    
    return fixed_df

def check_data_structure(df, ncb_columns, required_cols):
    """Comprehensive check of data structure to identify issues."""
    st.write("🔍 **COMPREHENSIVE DATA STRUCTURE CHECK**")
    
    # Check column names
    st.write(f"📊 **Total columns:** {len(df.columns)}")
    st.write(f"📊 **Total rows:** {len(df)}")
    
    # Check for duplicate column names in original data
    duplicate_original = df.columns[df.columns.duplicated()].tolist()
    if duplicate_original:
        st.error(f"❌ **DUPLICATE COLUMNS IN ORIGINAL DATA:** {duplicate_original}")
    else:
        st.write("✅ **No duplicate columns in original data**")
    
    # Check for None/empty column names
    empty_cols = [col for col in df.columns if col is None or str(col).strip() == '' or str(col) == 'nan']
    if empty_cols:
        st.warning(f"⚠️ **Empty/None column names:** {empty_cols}")
    
    # Check NCB columns
    st.write("🔍 **NCB Column Mapping Check:**")
    for ncb_type, col_name in ncb_columns.items():
        if col_name in df.columns:
            st.write(f"  ✅ {ncb_type}: {col_name} (exists)")
            # Check data type and sample values
            sample_vals = df[col_name].dropna().head(3).tolist()
            st.write(f"    Sample values: {sample_vals}")
        else:
            st.error(f"  ❌ {ncb_type}: {col_name} (MISSING)")
    
    # Check required columns
    st.write("🔍 **Required Column Mapping Check:**")
    for col_letter, col_name in required_cols.items():
        if col_name in df.columns:
            st.write(f"  ✅ {col_letter}: {col_name} (exists)")
        else:
            st.error(f"  ❌ {col_letter}: {col_name} (MISSING)")
    
    # Check transaction types
    transaction_col = required_cols.get('M')
    if transaction_col and transaction_col in df.columns:
        st.write("🔍 **Transaction Type Analysis:**")
        
        # Ensure we're working with a Series, not a DataFrame
        if isinstance(df[transaction_col], pd.DataFrame):
            # If it's a DataFrame, select the first column
            transaction_series = df[transaction_col].iloc[:, 0]
        else:
            # If it's already a Series, use it directly
            transaction_series = df[transaction_col]
        
        unique_transactions = transaction_series.astype(str).unique()
        st.write(f"  Unique values: {unique_transactions}")
        
        # Count each transaction type
        for trans_type in ['NB', 'C', 'R']:
            count = transaction_series.astype(str).str.upper().isin([trans_type]).sum()
            st.write(f"  {trans_type}: {count}")
    
    # Check NCB sum calculation
    st.write("🔍 **NCB Sum Calculation Check:**")
    ncb_cols = list(ncb_columns.values())
    if all(col in df.columns for col in ncb_cols):
        # Calculate NCB sum for sample rows
        sample_df = df.head(5)
        for idx, row in sample_df.iterrows():
            ncb_sum = sum(pd.to_numeric(row[col], errors='coerce') or 0 for col in ncb_cols)
            st.write(f"  Row {idx}: NCB Sum = {ncb_sum}")
    else:
        st.error("❌ **Cannot calculate NCB sum - missing columns**")
    
    return {
        'duplicate_columns': duplicate_original,
        'empty_columns': empty_cols,
        'missing_ncb': [col for col in ncb_columns.values() if col not in df.columns],
        'missing_required': [col for col in required_cols.values() if col not in df.columns]
    }

def clean_duplicate_columns(df):
    """Clean up duplicate column names in the original DataFrame."""
    if df is None or df.empty:
        return df
    
    # Check for duplicate column names
    duplicate_cols = df.columns[df.columns.duplicated()].tolist()
    if not duplicate_cols:
        return df
    
    st.warning(f"⚠️ **Found duplicate columns in original data: {duplicate_cols}**")
    st.write("🔧 **Cleaning duplicate columns...**")
    
    # Create new column names with suffixes for duplicates
    new_columns = []
    seen_columns = {}
    
    for col in df.columns:
        if col in seen_columns:
            seen_columns[col] += 1
            new_col = f"{col}_{seen_columns[col]}"
        else:
            seen_columns[col] = 0
            new_col = col
        
        new_columns.append(new_col)
    
    # Create new DataFrame with cleaned column names
    cleaned_df = df.copy()
    cleaned_df.columns = new_columns
    
    st.write(f"✅ **Cleaned DataFrame shape: {cleaned_df.shape}**")
    st.write(f"✅ **Duplicate columns resolved**")
    
    return cleaned_df

def clean_empty_column_names(df):
    """Clean up empty, None, or NaN column names."""
    if df is None or df.empty:
        return df
    
    # Check for empty/None/NaN column names
    empty_cols = [col for col in df.columns if col is None or str(col).strip() == '' or str(col) == 'nan']
    if not empty_cols:
        return df
    
    st.warning(f"⚠️ **Found empty/None column names: {empty_cols}**")
    st.write("🔧 **Cleaning empty column names...**")
    
    # Create new column names, replacing empty ones with descriptive names
    new_columns = []
    empty_count = 0
    
    for col in df.columns:
        if col is None or str(col).strip() == '' or str(col) == 'nan':
            new_col = f"Unnamed_Column_{empty_count}"
            empty_count += 1
        else:
            new_col = col
        
        new_columns.append(new_col)
    
    # Create new DataFrame with cleaned column names
    cleaned_df = df.copy()
    cleaned_df.columns = new_columns
    
    st.write(f"✅ **Cleaned DataFrame shape: {cleaned_df.shape}**")
    st.write(f"✅ **Empty column names resolved**")
    
    return cleaned_df

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
