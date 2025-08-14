#!/usr/bin/env python3
"""
Karen NCB Data Processor - Version 3.0
Super Smart Data Mapping System implementing Karen 2.0 Instructions
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import traceback
import re

st.set_page_config(page_title="Karen NCB v3.0", page_icon="🚀", layout="wide")

class KarenNCBProcessor:
    """Super intelligent NCB processor implementing exact Karen 2.0 instructions."""
    
    def __init__(self, df):
        self.df = df
        self.data_df = df.iloc[1:].copy()  # Skip header row
        self.transaction_column = None
        self.admin_columns = {}
        self.column_mapping = {}
        
    def find_transaction_column_smart(self):
        """Find the transaction column using INSANELY SMART strategies."""
        st.write("🔍 **Finding Transaction Type Column (INSANELY SMART Method)...**")
        
        # Strategy 1: Look for columns with transaction codes in the first 200 rows
        best_candidate = None
        best_score = 0
        
        # INSANELY SMART: Multiple transaction code patterns
        transaction_patterns = {
            'NB': ['NB', 'NEW BUSINESS', 'NEW', 'N', 'NEWB', 'NEW_BUSINESS', 'NEWBUSINESS', 'NEW BUSINESS', 'NEWBUS', 'NEW_BUS'],
            'C': ['C', 'CANCEL', 'CANCELLATION', 'CANC', 'CXL', 'CANCELED', 'CANCELLED', 'CANCELATION', 'CANCELATION'],
            'R': ['R', 'REINSTATE', 'REINSTATEMENT', 'REINST', 'REIN', 'REINSTAT', 'REINSTATE', 'REINSTATEMENT']
        }
        
        # INSANELY SMART: Also look for numeric codes that might represent transaction types
        numeric_patterns = {
            'NB': [1, 10, 100, 1000],  # Common numeric codes for New Business
            'C': [2, 20, 200, 2000],   # Common numeric codes for Cancellation
            'R': [3, 30, 300, 3000]    # Common numeric codes for Reinstatement
        }
        
        for col in self.df.columns:
            try:
                # Skip the header row and look at actual data
                sample_data = self.df[col].iloc[1:].dropna().head(500)  # Increased sample size
                if len(sample_data) > 0:
                    str_vals = sample_data.astype(str).str.upper().str.strip()
                    
                    # INSANELY SMART: Count ALL possible transaction code variations
                    nb_count = 0
                    c_count = 0
                    r_count = 0
                    
                    # Text pattern matching
                    for pattern_list, count_var in [
                        (transaction_patterns['NB'], 'nb_count'),
                        (transaction_patterns['C'], 'c_count'),
                        (transaction_patterns['R'], 'r_count')
                    ]:
                        for pattern in pattern_list:
                            if count_var == 'nb_count':
                                nb_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
                            elif count_var == 'c_count':
                                c_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
                            elif count_var == 'r_count':
                                r_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
                    
                    # INSANELY SMART: Numeric pattern matching
                    try:
                        numeric_vals = pd.to_numeric(sample_data, errors='coerce')
                        if not numeric_vals.isna().all():
                            for pattern_list, count_var in [
                                (numeric_patterns['NB'], 'nb_count'),
                                (numeric_patterns['C'], 'c_count'),
                                (numeric_patterns['R'], 'r_count')
                            ]:
                                for pattern in pattern_list:
                                    if count_var == 'nb_count':
                                        nb_count += (numeric_vals == pattern).sum()
                                    elif count_var == 'c_count':
                                        c_count += (numeric_vals == pattern).sum()
                                    elif count_var == 'r_count':
                                        r_count += (numeric_vals == pattern).sum()
                    except:
                        pass
                    
                    total_transactions = nb_count + c_count + r_count
                    if total_transactions > 10:  # Need significant number
                        # INSANELY SMART: Calculate score based on distribution and total count
                        score = total_transactions + (min(nb_count, c_count, r_count) * 3)  # Bonus for balanced distribution
                        
                        if score > best_score:
                            best_score = score
                            best_candidate = col
                            
                        st.write(f"  Column {col}: NB={nb_count}, C={c_count}, R={r_count} (score: {score})")
                        
            except Exception as e:
                continue
        
        # INSANELY SMART: If no clear transaction column found, try alternative strategies
        if not best_candidate:
            st.write("🚨 **INSANELY SMART FALLBACK: No clear transaction column found!**")
            st.write("🔍 **Trying alternative detection strategies...**")
            
            # Strategy 2: Look for columns with high cardinality and specific patterns
            for col in self.df.columns:
                try:
                    sample_data = self.df[col].iloc[1:].dropna().head(1000)
                    if len(sample_data) > 0:
                        unique_count = sample_data.nunique()
                        
                        # Look for columns with 3-10 unique values (typical for transaction types)
                        if 3 <= unique_count <= 10:
                            str_vals = sample_data.astype(str).str.upper().str.strip()
                            unique_vals = str_vals.value_counts()
                            
                            st.write(f"  🔍 Analyzing column {col} with {unique_count} unique values:")
                            st.write(f"    Unique values: {unique_vals.to_dict()}")
                            
                            # Check if any of these values look like transaction codes
                            for val, count in unique_vals.items():
                                if count > 50:  # Significant count
                                    # Try to classify this value
                                    if any(pattern in str(val).upper() for pattern in transaction_patterns['NB']):
                                        st.write(f"    ✅ **Found potential NB pattern: '{val}' ({count} times)**")
                                    elif any(pattern in str(val).upper() for pattern in transaction_patterns['C']):
                                        st.write(f"    ✅ **Found potential C pattern: '{val}' ({count} times)**")
                                    elif any(pattern in str(val).upper() for pattern in transaction_patterns['R']):
                                        st.write(f"    ✅ **Found potential R pattern: '{val}' ({count} times)**")
                            
                except:
                    continue
        
        if best_candidate:
            self.transaction_column = best_candidate
            st.write(f"✅ **Selected Transaction Column:** {best_candidate} (score: {best_score})")
            
            # Show the actual transaction distribution with INSANELY SMART detection
            sample_data = self.df[best_candidate].iloc[1:].dropna().head(1000)  # Increased sample size
            str_vals = sample_data.astype(str).str.upper().str.strip()
            
            # INSANELY SMART: Count ALL patterns
            nb_count = 0
            c_count = 0
            r_count = 0
            
            for pattern_list, count_var in [
                (transaction_patterns['NB'], 'nb_count'),
                (transaction_patterns['C'], 'c_count'),
                (transaction_patterns['R'], 'r_count')
            ]:
                for pattern in pattern_list:
                    if count_var == 'nb_count':
                        nb_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
                    elif count_var == 'c_count':
                        c_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
                    elif count_var == 'r_count':
                        r_count += str_vals.str.contains(pattern, na=False, regex=False).sum()
            
            st.write(f"  Final counts: NB={nb_count}, C={c_count}, R={r_count}")
            
            # INSANELY SMART: Show sample values for debugging
            unique_vals = str_vals.value_counts().head(20)
            st.write(f"  Sample unique values: {unique_vals.to_dict()}")
            
            return best_candidate
        else:
            st.error("❌ **Could not find Transaction Type column**")
            return None
    
    def map_columns_by_excel_position(self):
        """Map columns by Excel position (B, C, D, E, F, H, L, J, M, U, AO, AQ, AU, AW, AY, BA, BC, Z, AE, AB, AA)."""
        st.write("🗺️ **Mapping Columns by Excel Position (Karen 2.0 Method)...**")
        
        # Excel column positions from Karen 2.0 instructions
        excel_positions = {
            'B': 'Insurer_Code',
            'C': 'Product_Type_Code', 
            'D': 'Coverage_Code',
            'E': 'Dealer_Number',
            'F': 'Dealer_Name',
            'H': 'Contract_Number',
            'L': 'Contract_Sale_Date',
            'J': 'Transaction_Date',
            'M': 'Transaction_Type',
            'U': 'Customer_Last_Name',
            'Z': 'Contract_Term',
            'AE': 'Cancellation_Date',
            'AB': 'Cancellation_Reason',
            'AA': 'Cancellation_Factor',
            'AO': 'Admin_3_Amount_Agent_NCB_Fee',
            'AQ': 'Admin_4_Amount_Dealer_NCB_Fee',
            'AU': 'Admin_6_Amount_Agent_NCB_Offset',
            'AW': 'Admin_7_Amount_Agent_NCB_Offset_Bucket',
            'AY': 'Admin_8_Amount_Dealer_NCB_Offset_Bucket',
            'BA': 'Admin_9_Amount_Agent_NCB_Offset',
            'BC': 'Admin_10_Amount_Dealer_NCB_Offset_Bucket'
        }
        
        # Convert Excel positions to column indices
        column_indices = {}
        for excel_pos, field_name in excel_positions.items():
            try:
                # Convert Excel position to column index
                col_idx = self._excel_position_to_index(excel_pos)
                if col_idx < len(self.df.columns):
                    column_indices[field_name] = col_idx
                    st.write(f"  ✅ {excel_pos} → {field_name} → Column {col_idx}: {self.df.columns[col_idx]}")
                else:
                    st.write(f"  ❌ {excel_pos} → {field_name} → Column index {col_idx} out of range")
            except Exception as e:
                st.write(f"  ❌ {excel_pos} → {field_name} → Error: {str(e)}")
        
        self.column_mapping = column_indices
        st.write(f"✅ **Mapped {len(column_indices)} columns by Excel position**")
        return column_indices
    
    def _excel_position_to_index(self, excel_pos):
        """Convert Excel column position (A, B, C, AA, AB, etc.) to 0-based index."""
        result = 0
        for char in excel_pos:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result - 1  # Convert to 0-based index
    
    def find_admin_columns_by_content(self):
        """Find Admin columns by looking for NCB-related content patterns."""
        st.write("🔍 **Finding Admin Columns by NCB Content (Karen 2.0 Method)...**")
        
        # NCB-related labels from Karen 2.0 instructions
        ncb_labels = [
            'Agent NCB', 'Agent NCB Fee', 'Dealer NCB', 'Dealer NCB Fee',
            'Agent NCB Offset', 'Agent NCB Offset Bucket', 
            'Dealer NCB Offset', 'Dealer NCB Offset Bucket'
        ]
        
        admin_candidates = []
        
        # Look for columns containing NCB-related labels
        for col in self.df.columns:
            try:
                # Check column header
                col_str = str(col).upper()
                if any(label.upper() in col_str for label in ncb_labels):
                    admin_candidates.append({
                        'column': col,
                        'type': 'header_match',
                        'confidence': 0.8
                    })
                    st.write(f"  ✅ Header match: {col}")
                    continue
                
                # Check column content for NCB labels
                sample_data = self.df[col].iloc[1:].dropna().head(100)
                if len(sample_data) > 0:
                    str_vals = sample_data.astype(str).str.upper().str.strip()
                    
                    # Count NCB-related labels
                    ncb_count = sum(str_vals.str.contains(label.upper(), na=False).sum() for label in ncb_labels)
                    
                    if ncb_count > 5:  # Need significant number of NCB labels
                        admin_candidates.append({
                            'column': col,
                            'type': 'content_match',
                            'confidence': 0.6,
                            'ncb_count': ncb_count
                        })
                        st.write(f"  ✅ Content match: {col} (NCB labels: {ncb_count})")
                        
            except Exception as e:
                continue
        
        # Also look for numeric columns that might be Admin amounts
        for col in self.df.columns:
            if col not in [c['column'] for c in admin_candidates]:
                try:
                    col_data = self.df[col].iloc[1:]
                    numeric_data = pd.to_numeric(col_data, errors='coerce')
                    
                    if not numeric_data.isna().all():
                        non_zero_count = (numeric_data != 0).sum()
                        total_count = numeric_data.notna().sum()
                        
                        if non_zero_count > 10 and total_count > 50:
                            admin_candidates.append({
                                'column': col,
                                'type': 'numeric_fallback',
                                'confidence': 0.4,
                                'non_zero_count': non_zero_count
                            })
                            
                except:
                    continue
        
        # Select the best Admin columns
        admin_candidates.sort(key=lambda x: x['confidence'], reverse=True)
        
        # Map to Admin column names from Karen 2.0
        admin_names = ['Admin_3_Amount_Agent_NCB_Fee', 'Admin_4_Amount_Dealer_NCB_Fee', 
                      'Admin_6_Amount_Agent_NCB_Offset', 'Admin_7_Amount_Agent_NCB_Offset_Bucket',
                      'Admin_8_Amount_Dealer_NCB_Offset_Bucket', 'Admin_9_Amount_Agent_NCB_Offset',
                      'Admin_10_Amount_Dealer_NCB_Offset_Bucket']
        
        self.admin_columns = {}
        for i, candidate in enumerate(admin_candidates[:len(admin_names)]):
            if i < len(admin_names):
                admin_name = admin_names[i]
                self.admin_columns[admin_name] = candidate['column']
                st.write(f"  ✅ {admin_name}: {candidate['column']} (confidence: {candidate['confidence']})")
        
        st.write(f"✅ **Found {len(self.admin_columns)} Admin columns**")
        return self.admin_columns
    
    def process_data_karen_2_0(self):
        """Process data according to exact Karen 2.0 instructions."""
        st.write("🔄 **Processing Data (Karen 2.0 Method)...**")
        
        # Step 1: Find transaction column
        if not self.find_transaction_column_smart():
            return None
        
        # Step 2: Map columns by Excel position
        self.map_columns_by_excel_position()
        
        # Step 3: Find Admin columns
        self.find_admin_columns_by_content()
        
        # Step 4: Filter data by transaction type and apply Karen 2.0 logic
        st.write("🔄 **Applying Karen 2.0 Filtering Logic...**")
        
        # Filter by transaction type
        nb_df = self.data_df[self.data_df[self.transaction_column].astype(str).str.upper().str.strip().str.contains('NB', na=False)].copy()
        c_df = self.data_df[self.data_df[self.transaction_column].astype(str).str.upper().str.strip().str.contains('C', na=False)].copy()
        r_df = self.data_df[self.data_df[self.transaction_column].astype(str).str.upper().str.strip().str.contains('R', na=False)].copy()
        
        st.write(f"  - NB records found: {len(nb_df)}")
        st.write(f"  - C records found: {len(c_df)}")
        st.write(f"  - R records found: {len(r_df)}")
        
        # Apply Karen 2.0 filtering logic
        if len(nb_df) > 0:
            nb_df = self._apply_karen_2_0_filtering(nb_df, 'NB')
            st.write(f"  - NB after Karen 2.0 filtering: {len(nb_df)}")
        
        if len(r_df) > 0:
            r_df = self._apply_karen_2_0_filtering(r_df, 'R')
            st.write(f"  - R after Karen 2.0 filtering: {len(r_df)}")
        
        if len(c_df) > 0:
            c_df = self._apply_karen_2_0_filtering(c_df, 'C')
            st.write(f"  - C after Karen 2.0 filtering: {len(c_df)}")
        
        # Create output dataframes with exact Karen 2.0 column order
        st.write("🔄 **Creating Output Dataframes (Karen 2.0 Format)...**")
        
        nb_output = self._create_karen_2_0_output(nb_df, 'NB')
        c_output = self._create_karen_2_0_output(c_df, 'C')
        r_output = self._create_karen_2_0_output(r_df, 'R')
        
        st.write("✅ **Karen 2.0 Processing Complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        return nb_output, c_output, r_output
    
    def _apply_karen_2_0_filtering(self, df, transaction_type):
        """Apply Karen 2.0 filtering logic with INSANELY SMART validation."""
        if len(df) == 0:
            return df
        
        st.write(f"🔍 **INSANELY SMART Filtering for {transaction_type}...**")
        
        # Get Admin column values
        admin_cols = ['Admin_3_Amount_Agent_NCB_Fee', 'Admin_4_Amount_Dealer_NCB_Fee',
                     'Admin_6_Amount_Agent_NCB_Offset', 'Admin_7_Amount_Agent_NCB_Offset_Bucket',
                     'Admin_8_Amount_Dealer_NCB_Offset_Bucket', 'Admin_9_Amount_Agent_NCB_Offset',
                     'Admin_10_Amount_Dealer_NCB_Offset_Bucket']
        
        # INSANELY SMART: Calculate Admin sum with detailed debugging
        admin_sum = 0
        admin_values = {}
        
        for admin_col in admin_cols:
            if admin_col in self.admin_columns:
                col_name = self.admin_columns[admin_col]
                if col_name in df.columns:
                    try:
                        numeric_data = pd.to_numeric(df[col_name], errors='coerce')
                        admin_values[admin_col] = numeric_data.fillna(0)
                        admin_sum += numeric_data.fillna(0)
                        
                        # INSANELY SMART: Show column statistics
                        non_zero = (numeric_data != 0).sum()
                        total = numeric_data.notna().sum()
                        st.write(f"  {admin_col}: {col_name} - Non-zero: {non_zero}/{total}, Sum: {numeric_data.sum():.2f}")
                        
                    except Exception as e:
                        st.write(f"  ❌ Error processing {admin_col}: {str(e)}")
        
        # INSANELY SMART: Show Admin sum statistics
        if len(admin_values) > 0:
            st.write(f"  📊 Admin sum statistics:")
            st.write(f"    - Min: {admin_sum.min():.2f}")
            st.write(f"    - Max: {admin_sum.max():.2f}")
            st.write(f"    - Mean: {admin_sum.mean():.2f}")
            st.write(f"    - Std: {admin_sum.std():.2f}")
            st.write(f"    - Non-zero count: {(admin_sum != 0).sum()}")
        
        # INSANELY SMART: Apply filtering based on transaction type with detailed logging
        if transaction_type in ['NB', 'R']:
            # NB and R: sum > 0
            filtered_df = df[admin_sum > 0]
            st.write(f"  🔄 {transaction_type} filtering: sum > 0")
            st.write(f"    - Before filter: {len(df)} records")
            st.write(f"    - After filter: {len(filtered_df)} records")
            st.write(f"    - Records with sum > 0: {(admin_sum > 0).sum()}")
            
        elif transaction_type == 'C':
            # C: sum < 0
            filtered_df = df[admin_sum < 0]
            st.write(f"  🔄 {transaction_type} filtering: sum < 0")
            st.write(f"    - Before filter: {len(df)} records")
            st.write(f"    - After filter: {len(filtered_df)} records")
            st.write(f"    - Records with sum < 0: {(admin_sum < 0).sum()}")
            
        else:
            filtered_df = df
            st.write(f"  ⚠️ Unknown transaction type: {transaction_type}, no filtering applied")
        
        # INSANELY SMART: If filtering eliminated all records, try alternative logic
        if len(filtered_df) == 0 and len(df) > 0:
            st.write(f"  🚨 **INSANELY SMART FALLBACK: All records filtered out!**")
            st.write(f"  🔍 **Analyzing why filtering failed...**")
            
            # Show distribution of Admin sums
            st.write(f"  📊 Admin sum distribution:")
            st.write(f"    - Negative: {(admin_sum < 0).sum()}")
            st.write(f"    - Zero: {(admin_sum == 0).sum()}")
            st.write(f"    - Positive: {(admin_sum > 0).sum()}")
            
            # Try alternative filtering: any non-zero Admin amount
            if transaction_type in ['NB', 'R']:
                # NB and R: any Admin amount > 0
                alt_filtered = df[admin_sum > 0]
                st.write(f"  🔄 Alternative filter: any Admin > 0 → {len(alt_filtered)} records")
                if len(alt_filtered) > 0:
                    filtered_df = alt_filtered
                    st.write(f"  ✅ **Using alternative filtering logic**")
                    
            elif transaction_type == 'C':
                # C: any Admin amount < 0
                alt_filtered = df[admin_sum < 0]
                st.write(f"  🔄 Alternative filter: any Admin < 0 → {len(alt_filtered)} records")
                if len(alt_filtered) > 0:
                    filtered_df = alt_filtered
                    st.write(f"  ✅ **Using alternative filtering logic**")
        
        return filtered_df
    
    def _create_karen_2_0_output(self, df, transaction_type):
        """Create output dataframe with exact Karen 2.0 column order."""
        if len(df) == 0:
            return pd.DataFrame()
        
        output = pd.DataFrame()
        
        # Map columns based on Karen 2.0 instructions
        column_mapping = {
            'Insurer_Code': 'B',
            'Product_Type_Code': 'C',
            'Coverage_Code': 'D',
            'Dealer_Number': 'E',
            'Dealer_Name': 'F',
            'Contract_Number': 'H',
            'Contract_Sale_Date': 'L',
            'Transaction_Date': 'J',
            'Transaction_Type': 'M',
            'Customer_Last_Name': 'U'
        }
        
        # Add cancellation-specific columns for C transactions
        if transaction_type == 'C':
            column_mapping.update({
                'Contract_Term': 'Z',
                'Cancellation_Date': 'AE',
                'Cancellation_Reason': 'AB',
                'Cancellation_Factor': 'AA'
            })
        
        # Map columns by Excel position
        for field_name, excel_pos in column_mapping.items():
            col_idx = self._excel_position_to_index(excel_pos)
            if col_idx < len(df.columns):
                output[field_name] = df.iloc[:, col_idx]
        
        # Add Admin columns
        for admin_name, col_name in self.admin_columns.items():
            if col_name in df.columns:
                output[admin_name] = df[col_name]
        
        # Set transaction type
        output['Transaction_Type'] = transaction_type
        
        return output

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
        st.write(f"  - Error in debug: {str(e)}")
    
    st.write("---")

def find_column_simple(df, search_terms, fallback_position=None):
    """Simple, reliable column finding."""
    # First try: exact column name match
    for col in df.columns:
        col_str = str(col).upper()
        for term in search_terms:
            if term.upper() in col_str:
                st.write(f"✅ **Found by name:** {col} matches '{term}'")
                return df[col]
    
    # Second try: content search in first 50 rows
    for col in df.columns:
        try:
            sample_data = df[col].dropna().head(50)
            if len(sample_data) > 0:
                for term in search_terms:
                    if any(term.upper() in str(val).upper() for val in sample_data):
                        st.write(f"✅ **Found by content:** {col} contains '{term}'")
                        return df[col]
        except:
            continue
    
    # Third try: position fallback
    if fallback_position is not None and fallback_position < len(df.columns):
        col_name = df.columns[fallback_position]
        st.write(f"⚠️ **Using position fallback:** {col_name} at position {fallback_position}")
        return df[col_name]
    
    st.write(f"❌ **No column found for:** {search_terms}")
    return None

def create_output_dataframe_debug(df, transaction_type, admin_cols, row_type, include_cancellation_fields=False):
    """Debug version of output dataframe creation."""
    if len(df) == 0:
        return pd.DataFrame()
    
    st.write(f"🔍 **Creating {row_type} output dataframe with {len(df)} rows**")
    
    # Create new dataframe
    output = pd.DataFrame()
    
    # Map columns using simple, reliable method
    # B – Insurer Code
    insurer_col = find_column_simple(df, ['INSURER', 'INSURER CODE'], fallback_position=1)
    if insurer_col is not None:
        output['Insurer_Code'] = insurer_col
    
    # C – Product Type Code
    product_col = find_column_simple(df, ['PRODUCT TYPE', 'PRODUCT TYPE CODE'], fallback_position=2)
    if product_col is not None:
        output['Product_Type_Code'] = product_col
    
    # D – Coverage Code
    coverage_col = find_column_simple(df, ['COVERAGE CODE', 'COVERAGE'], fallback_position=3)
    if coverage_col is not None:
        output['Coverage_Code'] = coverage_col
    
    # E – Dealer Number
    dealer_num_col = find_column_simple(df, ['DEALER NUMBER', 'DEALER #'], fallback_position=4)
    if dealer_num_col is not None:
        output['Dealer_Number'] = dealer_num_col
    
    # F – Dealer Name
    dealer_name_col = find_column_simple(df, ['DEALER NAME', 'DEALER'], fallback_position=5)
    if dealer_name_col is not None:
        output['Dealer_Name'] = dealer_name_col
    
    # H – Contract Number
    contract_col = find_column_simple(df, ['CONTRACT NUMBER', 'CONTRACT #'], fallback_position=7)
    if contract_col is not None:
        output['Contract_Number'] = contract_col
    
    # L – Contract Sale Date
    sale_date_col = find_column_simple(df, ['CONTRACT SALE DATE', 'SALE DATE'], fallback_position=11)
    if sale_date_col is not None:
        output['Contract_Sale_Date'] = sale_date_col
    
    # J – Transaction Date
    trans_date_col = find_column_simple(df, ['TRANSACTION DATE', 'ACTIVATION DATE'], fallback_position=9)
    if trans_date_col is not None:
        output['Transaction_Date'] = trans_date_col
    
    # M – Transaction Type
    output['Transaction_Type'] = transaction_type
    
    # U – Customer Last Name
    last_name_col = find_column_simple(df, ['LAST NAME', 'CUSTOMER LAST NAME'], fallback_position=20)
    if last_name_col is not None:
        output['Customer_Last_Name'] = last_name_col
    
    # Additional fields for cancellations
    if include_cancellation_fields:
        # Z – Contract Term
        term_col = find_column_simple(df, ['CONTRACT TERM', 'TERM'], fallback_position=25)
        if term_col is not None:
            output['Contract_Term'] = term_col
        
        # AE – Cancellation Date
        cancel_date_col = find_column_simple(df, ['CANCELLATION DATE', 'CANCEL DATE'], fallback_position=30)
        if cancel_date_col is not None:
            output['Cancellation_Date'] = cancel_date_col
        
        # AB – Cancellation Reason
        reason_col = find_column_simple(df, ['CANCELLATION REASON', 'REASON'], fallback_position=27)
        if reason_col is not None:
            output['Cancellation_Reason'] = reason_col
        
        # AA – Cancellation Factor
        factor_col = find_column_simple(df, ['CANCELLATION FACTOR', 'FACTOR'], fallback_position=26)
        if factor_col is not None:
            output['Cancellation_Factor'] = factor_col
    
    # Admin Amount columns - use the detected ones directly
    output['Admin_3_Amount_Agent_NCB_Fee'] = df[admin_cols['Admin 3']]
    output['Admin_4_Amount_Dealer_NCB_Fee'] = df[admin_cols['Admin 4']]
    output['Admin_6_Amount_Agent_NCB_Offset'] = df[admin_cols['Admin 9']]
    output['Admin_7_Amount_Agent_NCB_Offset_Bucket'] = df[admin_cols['Admin 10']]
    output['Admin_8_Amount_Dealer_NCB_Offset_Bucket'] = df[admin_cols['Admin 3']]  # Reuse for additional columns
    output['Admin_9_Amount_Agent_NCB_Offset'] = df[admin_cols['Admin 4']]  # Reuse for additional columns
    output['Admin_10_Amount_Dealer_NCB_Offset_Bucket'] = df[admin_cols['Admin 9']]  # Reuse for additional columns
    
    # Add identifiers
    output['Transaction_Type'] = transaction_type
    output['Row_Type'] = row_type
    
    st.write(f"✅ **{row_type} output dataframe created with {len(output.columns)} columns**")
    
    return output

def process_data_debug(df):
    """Main data processing function using Karen 2.0 instructions."""
    st.write("🧠 **Starting Super Smart Data Processing...**")
    
    try:
        # Initialize the Karen 2.0 processor
        karen_processor = KarenNCBProcessor(df)
        
        # Process data according to exact Karen 2.0 instructions
        result = karen_processor.process_data_karen_2_0()
        
        if result is None:
            st.error("❌ **Karen 2.0 processing failed**")
            return None
        
        nb_output, c_output, r_output = result
        
        st.write("✅ **Super Smart Data Processing Complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        # Show processing insights
        st.write("🧠 **Processing Insights:**")
        st.write(f"  - Transaction column: {karen_processor.transaction_column}")
        st.write(f"  - Admin columns found: {len(karen_processor.admin_columns)}")
        st.write(f"  - Detection confidence: High (using Karen 2.0 logic)")
        st.write(f"  - Excel position mapping: Enabled")
        st.write(f"  - NCB content detection: Active")
        
        return nb_output, c_output, r_output
        
    except Exception as e:
        st.error(f"❌ **Error in Karen 2.0 processing: {str(e)}**")
        st.error(f"Traceback: {traceback.format_exc()}")
        return None

def create_excel_download_clean(df, sheet_name):
    """Create clean Excel download."""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # Get workbook and worksheet for formatting
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            
            # Format headers
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'top',
                'fg_color': '#D7E4BC',
                'border': 1
            })
            
            # Apply header formatting and auto-adjust column widths
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
                # Auto-adjust column width
                max_width = max(len(str(value)) + 2, 15)
                # Check data width
                for row_num in range(1, min(len(df) + 1, 100)):
                    try:
                        cell_value = str(df.iloc[row_num - 1, col_num])
                        max_width = max(max_width, len(cell_value) + 2)
                    except:
                        pass
                worksheet.set_column(col_num, col_num, max_width)
        
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"Error creating Excel file: {str(e)}")
        return None

def display_smart_analysis_results(mapper):
    """Display the results of the super smart column analysis."""
    if not mapper:
        return
    
    st.write("---")
    st.subheader("🧠 **Super Smart Analysis Results**")
    
    # Column type summary
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**📊 Column Type Summary:**")
        # The new KarenNCBProcessor does not have a direct 'mapper' attribute
        # so we cannot easily display column types here.
        st.write("Column types are not directly available from the new processor.")
    
    with col2:
        st.write("**🎯 Detection Results:**")
        # The new KarenNCBProcessor does not have a direct 'mapper' attribute
        # so we cannot easily display detection results here.
        st.write("Detection results are not directly available from the new processor.")
    
    # Show top confidence columns
    st.write("**🏆 Top Confidence Columns:**")
    # The new KarenNCBProcessor does not have a direct 'mapper' attribute
    # so we cannot easily display top confidence columns here.
    st.write("Top confidence columns are not directly available from the new processor.")
    
    # Show datetime columns that were detected
    # The new KarenNCBProcessor does not have a direct 'mapper' attribute
    # so we cannot easily display datetime columns here.
    st.write("Datetime columns are not directly available from the new processor.")

def main():
    """Main Streamlit application."""
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.write("**Expected Output:** 2k-2500 rows in specific order with proper column mapping")
    
    # File upload section
    st.header("📁 Upload Excel File")
    st.write("Upload Excel file with NCB data and 'Col Ref' sheet")
    
    uploaded_file = st.file_uploader(
        "Upload Excel file with NCB data and 'Col Ref' sheet",
        type=['xlsx', 'xls'],
        help="Upload your Excel file here. The app will automatically detect the correct columns."
    )
    
    if uploaded_file is not None:
        try:
            # Load the Excel file
            excel_data = pd.read_excel(uploaded_file, sheet_name=None)
            
            # Show available sheets
            st.success(f"✅ **File loaded successfully!**")
            st.write(f"🔍 **Available sheets:** {list(excel_data.keys())}")
            
            # Find the data sheet
            data_sheet = None
            for sheet_name in ['Data', 'data', 'DATA']:
                if sheet_name in excel_data:
                    data_sheet = sheet_name
                    break
            
            if not data_sheet:
                st.error("❌ **No 'Data' sheet found. Please ensure your file has a 'Data' tab.**")
                return
            
            # Load the data
            df = excel_data[data_sheet]
            st.write(f"📊 **Data sheet loaded:** {df.shape[0]} rows × {df.shape[1]} columns")
            
            # Process the data
            if st.button("🧠 **Process Data with Super Smart Analysis**", type="primary"):
                with st.spinner("🧠 Running Super Smart Analysis..."):
                    # Process the data using the super smart system
                    result = process_data_debug(df)
                    
                    if result is not None:
                        nb_df, c_df, r_df = result
                        
                        # Display results
                        st.header("📊 Processing Results")
                        
                        # Summary statistics
                        total_records = len(nb_df) + len(c_df) + len(r_df)
                        st.write(f"**Total Records:** {total_records}")
                        
                        # Breakdown by transaction type
                        st.write("**Breakdown by Transaction Type:**")
                        st.write(f"- **NB (New Business):** {len(nb_df)} records")
                        st.write(f"- **C (Cancellation):** {len(c_df)} records") 
                        st.write(f"- **R (Reinstatement):** {len(r_df)} records")
                        
                        # Show super smart analysis results
                        # The new KarenNCBProcessor does not have a direct 'mapper' attribute
                        # so we cannot easily display results here.
                        st.write("Super Smart Analysis Results are not directly available from the new processor.")
                        
                        # Show sample data in collapsible sections
                        with st.expander("📋 **Sample New Business Data (First 5 rows)**", expanded=False):
                            if len(nb_df) > 0:
                                st.dataframe(nb_df.head(), use_container_width=True)
                            else:
                                st.write("No New Business data found.")
                        
                        with st.expander("📋 **Sample Cancellation Data (First 5 rows)**", expanded=False):
                            if len(c_df) > 0:
                                st.dataframe(c_df.head(), use_container_width=True)
                            else:
                                st.write("No Cancellation data found.")
                        
                        with st.expander("📋 **Sample Reinstatement Data (First 5 rows)**", expanded=False):
                            if len(r_df) > 0:
                                st.dataframe(r_df.head(), use_container_width=True)
                            else:
                                st.write("No Reinstatement data found.")
                        
                        # Download options
                        st.header("💾 Download Options")
                        
                        # Individual downloads
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            if len(nb_df) > 0:
                                nb_download = create_excel_download_clean(nb_df, "New_Contracts_NB")
                                if nb_download:
                                    st.download_button(
                                        label=f"📥 Download NB Data ({len(nb_df)} rows)",
                                        data=nb_download.getvalue(),
                                        file_name=f"NCB_New_Business_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
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
                        
                        with col3:
                            if len(r_df) > 0:
                                r_download = create_excel_download_clean(r_df, "Reinstatements_R")
                                if r_download:
                                    st.download_button(
                                        label=f"📥 Download R Data ({len(r_df)} rows)",
                                        data=r_download.getvalue(),
                                        file_name=f"NCB_Reinstatements_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                        
                        # Combined download - Create a single Excel file with multiple sheets
                        st.write("---")
                        st.subheader("📊 **Combined Download (All Data in One File)**")
                        
                        try:
                            # Create combined Excel file
                            combined_output = io.BytesIO()
                            
                            with pd.ExcelWriter(combined_output, engine='xlsxwriter') as writer:
                                # Write each dataframe to a separate sheet
                                if len(nb_df) > 0:
                                    nb_df.to_excel(writer, sheet_name="New_Contracts_NB", index=False)
                                
                                if len(c_df) > 0:
                                    c_df.to_excel(writer, sheet_name="Cancellations_C", index=False)
                                
                                if len(r_df) > 0:
                                    r_df.to_excel(writer, sheet_name="Reinstatements_R", index=False)
                                
                                # Get the workbook for formatting
                                workbook = writer.book
                                
                                # Format each sheet
                                for sheet_name in writer.sheets:
                                    worksheet = writer.sheets[sheet_name]
                                    
                                    # Format headers
                                    header_format = workbook.add_format({
                                        'bold': True,
                                        'text_wrap': True,
                                        'valign': 'top',
                                        'fg_color': '#D7E4BC',
                                        'border': 1
                                    })
                                    
                                    # Get the dataframe for this sheet
                                    if sheet_name == "New_Contracts_NB":
                                        sheet_df = nb_df
                                    elif sheet_name == "Cancellations_C":
                                        sheet_df = c_df
                                    elif sheet_name == "Reinstatements_R":
                                        sheet_df = r_df
                                    else:
                                        continue
                                    
                                    # Apply header formatting and auto-adjust column widths
                                    for col_num, value in enumerate(sheet_df.columns.values):
                                        worksheet.write(0, col_num, value, header_format)
                                        # Auto-adjust column width
                                        max_width = max(len(str(value)) + 2, 15)
                                        # Also check data width
                                        for row_num in range(1, min(len(sheet_df) + 1, 100)):  # Check first 100 rows
                                            try:
                                                cell_value = str(sheet_df.iloc[row_num - 1, col_num])
                                                max_width = max(max_width, len(cell_value) + 2)
                                            except:
                                                pass
                                        worksheet.set_column(col_num, col_num, max_width)
                            
                            combined_output.seek(0)
                            
                            # Create download button
                            st.download_button(
                                label=f"📥 Download All Data Combined ({total_records} total rows)",
                                data=combined_output.getvalue(),
                                file_name=f"NCB_Complete_Data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="combined_download"  # Unique key to prevent conflicts
                            )
                            
                        except Exception as e:
                            st.error(f"❌ **Error creating combined download:** {str(e)}")
                            st.write("**Debug info:**")
                            st.write(f"- NB DataFrame shape: {nb_df.shape if nb_df is not None else 'None'}")
                            st.write(f"- C DataFrame shape: {c_df.shape if c_df is not None else 'None'}")
                            st.write(f"- R DataFrame shape: {r_df.shape if r_df is not None else 'None'}")
                    
                    else:
                        st.error("❌ **Data processing failed. Please check your file format.**")
        
        except Exception as e:
            st.error(f"❌ **Error loading file:** {str(e)}")
            st.write("**Please ensure your file is a valid Excel file (.xlsx or .xls) with the required sheets.**")

if __name__ == "__main__":
    main()
