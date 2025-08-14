#!/usr/bin/env python3
"""
Karen NCB Data Processor - Version 3.0
Super Smart Data Mapping System with Intelligent Column Discovery
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import traceback
import re

st.set_page_config(page_title="Karen NCB v3.0", page_icon="🚀", layout="wide")

class SmartColumnMapper:
    """Super intelligent column mapping system that learns from data patterns."""
    
    def __init__(self, df):
        self.df = df
        self.data_df = df.iloc[1:].copy()  # Skip header row
        self.column_analysis = {}
        self.admin_columns = {}
        self.transaction_column = None
        
    def analyze_all_columns(self):
        """Comprehensive analysis of all columns to understand their nature."""
        st.write("🧠 **Starting Super Smart Column Analysis...**")
        
        for col in self.df.columns:
            self.analyze_single_column(col)
        
        st.write(f"✅ **Analyzed {len(self.df.columns)} columns**")
        return self.column_analysis
    
    def analyze_single_column(self, col):
        """Deep analysis of a single column to determine its type and purpose."""
        try:
            col_data = self.data_df[col]
            
            analysis = {
                'column_name': col,
                'data_type': str(col_data.dtype),
                'total_rows': len(col_data),
                'non_null_count': col_data.count(),
                'null_percentage': (col_data.isna().sum() / len(col_data)) * 100,
                'unique_values': col_data.nunique(),
                'is_datetime': False,
                'is_numeric': False,
                'is_categorical': False,
                'is_transaction_type': False,
                'is_admin_amount': False,
                'confidence_score': 0.0,
                'sample_values': [],
                'patterns': [],
                'recommendations': []
            }
            
            # Check if it's a datetime column
            if self._is_datetime_column(col, col_data):
                analysis['is_datetime'] = True
                analysis['confidence_score'] += 0.1
            
            # Check if it's numeric
            if self._is_numeric_column(col_data):
                analysis['is_numeric'] = True
                analysis['confidence_score'] += 0.2
                
                # Check if it looks like financial data
                if self._looks_like_financial_data(col_data):
                    analysis['is_admin_amount'] = True
                    analysis['confidence_score'] += 0.3
                    analysis['recommendations'].append("Potential Admin amount column")
            
            # Check if it's categorical
            if self._is_categorical_column(col_data):
                analysis['is_categorical'] = True
                analysis['confidence_score'] += 0.1
            
            # Check if it's transaction type
            if self._is_transaction_type_column(col_data):
                analysis['is_transaction_type'] = True
                analysis['confidence_score'] += 0.4
                analysis['recommendations'].append("Transaction Type column detected")
            
            # Get sample values
            analysis['sample_values'] = col_data.dropna().head(5).tolist()
            
            # Detect patterns
            analysis['patterns'] = self._detect_patterns(col_data)
            
            self.column_analysis[col] = analysis
            
        except Exception as e:
            st.write(f"❌ **Error analyzing column {col}: {str(e)}**")
    
    def _is_datetime_column(self, col_name, col_data):
        """Smart datetime detection that avoids false positives."""
        # Check column name
        col_str = str(col_name).lower()
        if any(pattern in col_str for pattern in ['datetime', 'timestamp', 'date', 'time']):
            return True
        
        # Check for actual datetime objects
        if col_data.dtype == 'datetime64[ns]':
            return True
        
        # Check for datetime-like patterns in data
        sample_str = str(col_data.head(10).tolist())
        if any(pattern in sample_str for pattern in ['2025-', '2024-', '2023-', '2022-', '2021-', '2020-']):
            return True
        
        return False
    
    def _is_numeric_column(self, col_data):
        """Smart numeric detection with tolerance for mixed data."""
        try:
            # Try to convert to numeric
            numeric_data = pd.to_numeric(col_data, errors='coerce')
            non_null_count = numeric_data.notna().sum()
            
            # Consider it numeric if at least 70% of values convert successfully
            return (non_null_count / len(col_data)) > 0.7
        except:
            return False
    
    def _looks_like_financial_data(self, col_data):
        """Detect if numeric column contains financial data patterns."""
        try:
            numeric_data = pd.to_numeric(col_data, errors='coerce')
            
            # Check for financial patterns
            has_decimals = (numeric_data % 1 != 0).any()
            has_negative = (numeric_data < 0).any()
            has_large_values = (abs(numeric_data) > 1000).any()
            
            # Financial data typically has these characteristics
            return has_decimals or has_negative or has_large_values
        except:
            return False
    
    def _is_categorical_column(self, col_data):
        """Detect categorical columns."""
        unique_ratio = col_data.nunique() / len(col_data)
        return unique_ratio < 0.1  # Less than 10% unique values
    
    def _is_transaction_type_column(self, col_data):
        """Smart detection of transaction type columns."""
        try:
            # Convert to string and look for transaction codes
            str_vals = col_data.astype(str).str.upper().str.strip()
            
            # Count transaction codes
            nb_count = str_vals.str.contains('NB', na=False).sum()
            c_count = str_vals.str.contains('C', na=False).sum()
            r_count = str_vals.str.contains('R', na=False).sum()
            
            total_transactions = nb_count + c_count + r_count
            total_rows = len(str_vals)
            
            # Consider it a transaction column if it has significant transaction codes
            return total_transactions > 0 and (total_transactions / total_rows) > 0.05
        except:
            return False
    
    def _detect_patterns(self, col_data):
        """Detect various patterns in the data."""
        patterns = []
        
        try:
            # Check for constant values
            if col_data.nunique() == 1:
                patterns.append("Constant value")
            
            # Check for sequential numbers
            if self._is_sequential(col_data):
                patterns.append("Sequential numbers")
            
            # Check for currency patterns
            if self._has_currency_patterns(col_data):
                patterns.append("Currency patterns")
            
            # Check for percentage patterns
            if self._has_percentage_patterns(col_data):
                patterns.append("Percentage patterns")
                
        except:
            pass
        
        return patterns
    
    def _is_sequential(self, col_data):
        """Check if column contains sequential numbers."""
        try:
            numeric_data = pd.to_numeric(col_data, errors='coerce')
            if numeric_data.isna().all():
                return False
            
            # Check if values are roughly sequential
            sorted_vals = numeric_data.dropna().sort_values()
            if len(sorted_vals) < 3:
                return False
            
            # Check if differences are roughly constant
            diffs = sorted_vals.diff().dropna()
            if len(diffs) == 0:
                return False
            
            # If 80% of differences are the same, consider it sequential
            most_common_diff = diffs.mode().iloc[0] if not diffs.mode().empty else 0
            same_diff_count = (diffs == most_common_diff).sum()
            
            return (same_diff_count / len(diffs)) > 0.8
        except:
            return False
    
    def _has_currency_patterns(self, col_data):
        """Check for currency-like patterns."""
        try:
            str_data = col_data.astype(str)
            has_dollar_sign = str_data.str.contains('$', na=False).any()
            has_commas = str_data.str.contains(',', na=False).any()
            has_parentheses = str_data.str.contains(r'[()]', na=False).any()
            
            return has_dollar_sign or has_commas or has_parentheses
        except:
            return False
    
    def _has_percentage_patterns(self, col_data):
        """Check for percentage-like patterns."""
        try:
            str_data = col_data.astype(str)
            has_percent = str_data.str.contains('%', na=False).any()
            
            if has_percent:
                return True
            
            # Check if values are between 0 and 1 (potential percentages)
            numeric_data = pd.to_numeric(col_data, errors='coerce')
            if not numeric_data.isna().all():
                in_range = ((numeric_data >= 0) & (numeric_data <= 1)).sum()
                return (in_range / numeric_data.notna().sum()) > 0.5
            
            return False
        except:
            return False
    
    def find_transaction_column(self):
        """Find the transaction type column using multiple strategies."""
        st.write("🔍 **Finding Transaction Type Column...**")
        
        # Strategy 1: Look for columns marked as transaction type
        transaction_candidates = []
        for col, analysis in self.column_analysis.items():
            if analysis['is_transaction_type']:
                transaction_candidates.append((col, analysis['confidence_score']))
        
        if transaction_candidates:
            # Sort by confidence score
            transaction_candidates.sort(key=lambda x: x[1], reverse=True)
            best_candidate = transaction_candidates[0]
            
            st.write(f"✅ **Found Transaction Type column:** {best_candidate[0]} (confidence: {best_candidate[1]:.2f})")
            self.transaction_column = best_candidate[0]
            return best_candidate[0]
        
        # Strategy 2: Look for columns with transaction codes in sample data
        st.write("🔄 **Strategy 1 failed, trying content-based detection...**")
        
        for col in self.df.columns:
            try:
                sample_data = self.data_df[col].dropna().head(100)
                str_vals = sample_data.astype(str).str.upper().str.strip()
                
                nb_count = str_vals.str.contains('NB', na=False).sum()
                c_count = str_vals.str.contains('C', na=False).sum()
                r_count = str_vals.str.contains('R', na=False).sum()
                
                total_transactions = nb_count + c_count + r_count
                if total_transactions > 10:  # Need significant number of transaction codes
                    st.write(f"✅ **Found Transaction Type column by content:** {col}")
                    st.write(f"   Sample counts: NB={nb_count}, C={c_count}, R={r_count}")
                    self.transaction_column = col
                    return col
                    
            except Exception as e:
                continue
        
        st.error("❌ **Could not find Transaction Type column**")
        return None
    
    def find_admin_columns(self):
        """Find Admin columns using multiple intelligent strategies."""
        st.write("🔍 **Finding Admin Columns with Super Smart Detection...**")
        
        # Strategy 1: Look for columns marked as admin amounts
        admin_candidates = []
        for col, analysis in self.column_analysis.items():
            if analysis['is_admin_amount'] and analysis['is_numeric']:
                admin_candidates.append({
                    'column': col,
                    'confidence': analysis['confidence_score'],
                    'non_null_count': analysis['non_null_count'],
                    'unique_values': analysis['unique_values']
                })
        
        # Strategy 2: Look for columns with financial patterns
        if len(admin_candidates) < 4:
            st.write("🔄 **Strategy 1 found insufficient columns, trying pattern detection...**")
            
            for col, analysis in self.column_analysis.items():
                if analysis['is_numeric'] and col not in [c['column'] for c in admin_candidates]:
                    # Check for financial characteristics
                    col_data = self.data_df[col]
                    numeric_data = pd.to_numeric(col_data, errors='coerce')
                    
                    if not numeric_data.isna().all():
                        # Calculate financial metrics
                        non_zero_count = (numeric_data != 0).sum()
                        total_count = numeric_data.notna().sum()
                        has_negative = (numeric_data < 0).any()
                        has_decimals = (numeric_data % 1 != 0).any()
                        
                        # Score based on financial characteristics
                        financial_score = 0
                        if non_zero_count > 5:
                            financial_score += 0.3
                        if has_negative:
                            financial_score += 0.2
                        if has_decimals:
                            financial_score += 0.2
                        if total_count > 100:
                            financial_score += 0.3
                        
                        if financial_score > 0.5:  # Only include if it looks financial
                            admin_candidates.append({
                                'column': col,
                                'confidence': financial_score,
                                'non_null_count': total_count,
                                'unique_values': analysis['unique_values']
                            })
        
        # Strategy 3: Look for columns by name/content if still insufficient
        if len(admin_candidates) < 4:
            st.write("🔄 **Strategy 2 found insufficient columns, trying name-based detection...**")
            
            for col, analysis in self.column_analysis.items():
                if analysis['is_numeric'] and col not in [c['column'] for c in admin_candidates]:
                    col_str = str(col).lower()
                    
                    # Look for financial keywords in column names
                    financial_keywords = ['admin', 'amount', 'fee', 'ncb', 'agent', 'dealer', 'cost', 'price', 'value']
                    if any(keyword in col_str for keyword in financial_keywords):
                        admin_candidates.append({
                            'column': col,
                            'confidence': 0.4,  # Lower confidence for name-based detection
                            'non_null_count': analysis['non_null_count'],
                            'unique_values': analysis['unique_values']
                        })
        
        # Sort by confidence and select the best 4
        admin_candidates.sort(key=lambda x: x['confidence'], reverse=True)
        
        if len(admin_candidates) < 4:
            st.error(f"❌ **Only found {len(admin_candidates)} Admin columns, need 4**")
            return None
        
        # Select the top 4
        selected_admin_cols = admin_candidates[:4]
        
        # Map to Admin column names
        admin_names = ['Admin 3', 'Admin 4', 'Admin 9', 'Admin 10']
        self.admin_columns = {}
        
        for i, candidate in enumerate(selected_admin_cols):
            admin_name = admin_names[i]
            self.admin_columns[admin_name] = candidate['column']
            
            st.write(f"✅ **{admin_name}:** {candidate['column']} (confidence: {candidate['confidence']:.2f})")
            st.write(f"   Non-null: {candidate['non_null_count']}, Unique: {candidate['unique_values']}")
        
        return self.admin_columns
    
    def get_column_summary(self):
        """Get a summary of all column analysis."""
        st.write("📊 **Column Analysis Summary**")
        
        # Group columns by type
        datetime_cols = [col for col, analysis in self.column_analysis.items() if analysis['is_datetime']]
        numeric_cols = [col for col, analysis in self.column_analysis.items() if analysis['is_numeric']]
        categorical_cols = [col for col, analysis in self.column_analysis.items() if analysis['is_categorical']]
        admin_cols = [col for col, analysis in self.column_analysis.items() if analysis['is_admin_amount']]
        
        st.write(f"  - Datetime columns: {len(datetime_cols)}")
        st.write(f"  - Numeric columns: {len(numeric_cols)}")
        st.write(f"  - Categorical columns: {len(categorical_cols)}")
        st.write(f"  - Admin amount columns: {len(admin_cols)}")
        
        if datetime_cols:
            st.write(f"  - Datetime columns: {datetime_cols[:3]}...")
        if admin_cols:
            st.write(f"  - Admin columns: {admin_cols[:3]}...")

def analyze_data_structure_smart(df):
    """Super smart data structure analysis using the intelligent mapper."""
    st.write("🧠 **Starting Super Smart Data Structure Analysis...**")
    
    # Initialize the smart mapper
    mapper = SmartColumnMapper(df)
    
    # Analyze all columns
    mapper.analyze_all_columns()
    
    # Get column summary
    mapper.get_column_summary()
    
    # Find transaction column
    transaction_col = mapper.find_transaction_column()
    if not transaction_col:
        return None
    
    # Find admin columns
    admin_columns = mapper.find_admin_columns()
    if not admin_columns:
        return None
    
    st.write("✅ **Super Smart Analysis Complete!**")
    
    return {
        'transaction_col': transaction_col,
        'admin_columns': admin_columns,
        'mapper': mapper
    }

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
    """Debug version of data processing."""
    try:
        st.write("🔍 **Starting DEBUG data processing...**")
        
        # Analyze data structure
        structure_info = analyze_data_structure_smart(df)
        if not structure_info:
            st.error("❌ **Could not analyze data structure**")
            return None
        
        transaction_col = structure_info['transaction_col']
        admin_cols = structure_info['admin_columns']
        
        st.write(f"✅ **Data structure analysis complete**")
        st.write(f"  - Transaction column: {transaction_col}")
        st.write(f"  - Admin columns: {admin_cols}")
        
        # DEBUG: Show final Admin column details
        st.write("🔍 **DEBUG: Final Admin column validation before processing**")
        for admin_type, col_name in admin_cols.items():
            debug_column_info(df, col_name, f"Pre-Processing - {admin_type}")
        
        # Filter by transaction type - skip the header row
        st.write("🔄 **Filtering data by transaction type...**")
        
        # Skip the first row (header) and work with actual data
        data_df = df.iloc[1:].copy()
        
        # New Business (NB)
        nb_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'NB'].copy()
        st.write(f"  - NB records found: {len(nb_df)}")
        
        # Cancellations (C)
        c_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'C'].copy()
        st.write(f"  - C records found: {len(c_df)}")
        
        # Reinstatements (R)
        r_df = data_df[data_df[transaction_col].astype(str).str.upper().str.strip() == 'R'].copy()
        st.write(f"  - R records found: {len(r_df)}")
        
        # Apply Admin amount filtering with DEBUG
        st.write("🔄 **Applying Admin amount filtering with DEBUG...**")
        
        # For NB: sum > 0 AND all individual amounts > 0
        if len(nb_df) > 0:
            admin_cols_list = list(admin_cols.values())
            
            st.write(f"🔍 **DEBUG: Admin columns for NB filtering:**")
            for i, col in enumerate(admin_cols_list):
                st.write(f"  Admin {i+1}: {col}")
                debug_column_info(nb_df, col, f"NB Filtering - Admin {i+1}")
            
            # Convert to numeric and check for errors
            st.write("🔄 **Converting Admin columns to numeric...**")
            numeric_admin_cols = []
            
            for col in admin_cols_list:
                try:
                    numeric_col = pd.to_numeric(nb_df[col], errors='coerce')
                    if not numeric_col.isna().all():
                        numeric_admin_cols.append(numeric_col)
                        st.write(f"✅ **Successfully converted {col} to numeric**")
                    else:
                        st.write(f"❌ **Failed to convert {col} to numeric**")
                except Exception as e:
                    st.write(f"❌ **Error converting {col}: {str(e)}**")
            
            if len(numeric_admin_cols) < 4:
                st.error(f"❌ **Only {len(numeric_admin_cols)} Admin columns could be converted to numeric**")
                return None
            
            # Calculate sum
            st.write("🔄 **Calculating Admin sum...**")
            try:
                nb_df['Admin_Sum'] = pd.concat(numeric_admin_cols, axis=1).sum(axis=1)
                st.write(f"✅ **Admin sum calculated successfully**")
            except Exception as e:
                st.write(f"❌ **Error calculating Admin sum: {str(e)}**")
                st.write(f"  - Numeric columns: {len(numeric_admin_cols)}")
                st.write(f"  - Column shapes: {[col.shape for col in numeric_admin_cols]}")
                return None
            
            # First filter: sum > 0
            nb_filtered = nb_df[nb_df['Admin_Sum'] > 0]
            st.write(f"  - NB after Admin sum > 0 filter: {len(nb_filtered)} records")
            
            # Second filter: all individual amounts > 0
            nb_final = nb_filtered[
                (numeric_admin_cols[0] > 0) & 
                (numeric_admin_cols[1] > 0) & 
                (numeric_admin_cols[2] > 0) & 
                (numeric_admin_cols[3] > 0)
            ]
            st.write(f"  - NB after individual Admin > 0 filter: {len(nb_final)} records")
            nb_df = nb_final
        
        # For R: sum > 0
        if len(r_df) > 0:
            admin_cols_list = list(admin_cols.values())
            r_df['Admin_Sum'] = r_df[admin_cols_list].sum(axis=1)
            r_filtered = r_df[r_df['Admin_Sum'] > 0]
            st.write(f"  - R after Admin sum > 0 filter: {len(r_filtered)} records")
            r_df = r_filtered
        
        # For C: sum < 0
        if len(c_df) > 0:
            admin_cols_list = list(admin_cols.values())
            c_df['Admin_Sum'] = c_df[admin_cols_list].sum(axis=1)
            c_filtered = c_df[c_df['Admin_Sum'] < 0]
            st.write(f"  - C after Admin sum < 0 filter: {len(c_filtered)} records")
            c_df = c_filtered
        
        # Create output dataframes
        st.write("🔄 **Creating output dataframes...**")
        
        nb_output = create_output_dataframe_debug(nb_df, 'NB', admin_cols, 'New Business', False)
        c_output = create_output_dataframe_debug(c_df, 'C', admin_cols, 'Cancellation', True)
        r_output = create_output_dataframe_debug(r_df, 'R', admin_cols, 'Reinstatement', False)
        
        st.write("✅ **Data processing complete!**")
        st.write(f"  - Final NB records: {len(nb_output)}")
        st.write(f"  - Final C records: {len(c_output)}")
        st.write(f"  - Final R records: {len(r_output)}")
        st.write(f"  - Total records: {len(nb_output) + len(c_output) + len(r_output)}")
        
        return nb_output, c_output, r_output
        
    except Exception as e:
        st.error(f"❌ **Error in DEBUG data processing:** {str(e)}")
        st.write("**Full error details:**")
        st.code(traceback.format_exc())
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

def main():
    st.set_page_config(
        page_title="Karen NCB Data Processor - Version 3.0",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("🚀 Karen NCB Data Processor - Version 3.0")
    st.markdown("**Expected Output:** 2k-2500 rows in specific order with proper column mapping")
    
    # File upload section
    st.header("📁 Upload Excel File")
    uploaded_file = st.file_uploader(
        "Upload Excel file with NCB data and 'Col Ref' sheet",
        type=['xlsx', 'xls'],
        help="Your file should have a 'Data' tab and a reference sheet (like 'Col Ref' or 'xref')"
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
            if st.button("🔄 **Process Data**", type="primary"):
                with st.spinner("Processing data..."):
                    # Process the data using the enhanced functions
                    nb_df, c_df, r_df = process_data_debug(df)
                    
                    if nb_df is not None and c_df is not None and r_df is not None:
                        
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
