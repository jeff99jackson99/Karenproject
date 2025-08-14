#!/usr/bin/env python3
"""
Comprehensive Test Script for Karen 2.0 NCB Data Processor
Tests all functions for errors before deployment
"""

import pandas as pd
import numpy as np
import sys
import os

# Add the current directory to Python path to import functions
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_dataframe_creation():
    """Test basic DataFrame creation and operations."""
    print("🔍 Testing DataFrame creation...")
    
    try:
        # Create test data similar to the Excel structure
        test_data = {
            'Row Number': [1, 2, 3],
            'Insurer Code': ['INS1', 'INS2', 'INS3'],
            'Product Type Code': ['PROD1', 'PROD2', 'PROD3'],
            'Coverage Code': ['COV1', 'COV2', 'COV3'],
            'Dealer Number': ['D1', 'D2', 'D3'],
            'Dealer Name': ['Dealer1', 'Dealer2', 'Dealer3'],
            'Dealer Insured State': ['CA', 'NY', 'TX'],
            'Contract Number': ['C1', 'C2', 'C3'],
            'Transaction Date': ['2025-01-01', '2025-01-02', '2025-01-03'],
            'Transaction Type': ['NB', 'C', 'R'],
            'Billed Date': ['2025-01-01', '2025-01-02', '2025-01-03'],
            'Contract Sale Date': ['2025-01-01', '2025-01-02', '2025-01-03'],
            'Customer Last Name': ['Smith', 'Jones', 'Brown'],
            'Customer First Name': ['John', 'Jane', 'Bob'],
            'Vehicle Age Type': ['New', 'Used', 'New'],
            'VIN': ['VIN1', 'VIN2', 'VIN3'],
            'Vehicle Model Year': [2025, 2024, 2025],
            'Vehicle Make': ['Toyota', 'Honda', 'Ford'],
            'Vehicle Model': ['Camry', 'Civic', 'F-150'],
            'Vehicle Class': ['Sedan', 'Sedan', 'Truck'],
            'Term Months': [36, 48, 60],
            'Cancellation Factor': [0.8, 0.9, 0.7],
            'Cancellation Reason': ['None', 'Customer Request', 'None'],
            'Cancellation Date': ['', '2025-01-15', ''],
            'ADMIN 3 Amount': [100.0, -50.0, 75.0],
            'ADMIN 4 Amount': [200.0, -100.0, 150.0],
            'ADMIN 6 Amount': [50.0, -25.0, 37.5],
            'ADMIN 7 Amount': [25.0, -12.5, 18.75],
            'ADMIN 8 Amount': [75.0, -37.5, 56.25],
            'ADMIN 9 Amount': [125.0, -62.5, 93.75],
            'ADMIN 10 Amount': [150.0, -75.0, 112.5]
        }
        
        df = pd.DataFrame(test_data)
        print(f"✅ DataFrame created successfully: {df.shape}")
        print(f"✅ Columns: {len(df.columns)}")
        return df
        
    except Exception as e:
        print(f"❌ DataFrame creation failed: {e}")
        return None

def test_clean_duplicate_columns(df):
    """Test the clean_duplicate_columns function."""
    print("\n🔍 Testing clean_duplicate_columns function...")
    
    try:
        # Create duplicate columns by renaming existing columns
        df_with_duplicates = df.copy()
        # Rename the last column to create a duplicate
        new_columns = list(df_with_duplicates.columns)
        new_columns[-1] = new_columns[0]  # Make last column same as first
        df_with_duplicates.columns = new_columns
        
        # Import and test the function
        from karen_2_0_app import clean_duplicate_columns
        
        cleaned_df = clean_duplicate_columns(df_with_duplicates)
        
        if cleaned_df is not None:
            print(f"✅ Duplicate columns cleaned: {cleaned_df.shape}")
            duplicate_cols = cleaned_df.columns[cleaned_df.columns.duplicated()].tolist()
            if not duplicate_cols:
                print("✅ No duplicate columns remaining")
                return True  # Just return success, don't modify original df
            else:
                print(f"❌ Still have duplicate columns: {duplicate_cols}")
                return False
        else:
            print("❌ Function returned None")
            return False
            
    except Exception as e:
        print(f"❌ clean_duplicate_columns test failed: {e}")
        return False

def test_clean_empty_column_names(df):
    """Test the clean_empty_column_names function."""
    print("\n🔍 Testing clean_empty_column_names function...")
    
    try:
        # Create columns with empty names by renaming existing columns
        df_with_empty = df.copy()
        new_columns = list(df_with_empty.columns)
        # Replace some column names with empty values
        new_columns[0] = np.nan  # First column becomes NaN
        new_columns[1] = ''      # Second column becomes empty string
        df_with_empty.columns = new_columns
        
        # Import and test the function
        from karen_2_0_app import clean_empty_column_names
        
        cleaned_df = clean_empty_column_names(df_with_empty)
        
        if cleaned_df is not None:
            print(f"✅ Empty column names cleaned: {cleaned_df.shape}")
            empty_cols = [col for col in cleaned_df.columns if col is None or str(col).strip() == '' or str(col) == 'nan']
            if not empty_cols:
                print("✅ No empty column names remaining")
                return True  # Just return success, don't modify original df
            else:
                print(f"❌ Still have empty column names: {empty_cols}")
                return False
        else:
            print("❌ Function returned None")
            return False
            
    except Exception as e:
        print(f"❌ clean_empty_column_names test failed: {e}")
        return False

def test_find_ncb_columns_simple(df):
    """Test the find_ncb_columns_simple function."""
    print("\n🔍 Testing find_ncb_columns_simple function...")
    
    try:
        # Import and test the function
        from karen_2_0_app import find_ncb_columns_simple
        
        # Mock st.write function for testing
        import karen_2_0_app
        mock_st = type('MockStreamlit', (), {
            'write': lambda x, *args, **kwargs: None,
            'warning': lambda x, *args, **kwargs: None,
            'error': lambda x, *args, **kwargs: None,
            'info': lambda x, *args, **kwargs: None,
            'success': lambda x, *args, **kwargs: None
        })()
        karen_2_0_app.st = mock_st
        
        # Debug: Show all columns in test data
        print(f"  Test data columns: {list(df.columns)}")
        
        # Debug: Show what we're looking for
        admin_mapping = {
            'ADMIN 3 Amount': 'AO',
            'ADMIN 4 Amount': 'AQ', 
            'ADMIN 6 Amount': 'AU',
            'ADMIN 7 Amount': 'AW',
            'ADMIN 8 Amount': 'AY',
            'ADMIN 9 Amount': 'BA',
            'ADMIN 10 Amount': 'BC',
        }
        print(f"  Looking for Admin columns: {list(admin_mapping.keys())}")
        
        ncb_columns = find_ncb_columns_simple(df)
        
        if ncb_columns and len(ncb_columns) >= 7:
            print(f"✅ Found {len(ncb_columns)} NCB columns")
            for ncb_type, col_name in ncb_columns.items():
                print(f"  {ncb_type}: {col_name}")
            return ncb_columns
        else:
            print(f"❌ Expected 7+ NCB columns, got {len(ncb_columns) if ncb_columns else 0}")
            if ncb_columns:
                print("  Found columns:")
                for ncb_type, col_name in ncb_columns.items():
                    print(f"    {ncb_type}: {col_name}")
            
            # Debug: Check what Admin columns exist in test data
            admin_cols = [col for col in df.columns if 'ADMIN' in str(col)]
            print(f"  Admin columns in test data: {admin_cols}")
            return None
            
    except Exception as e:
        print(f"❌ find_ncb_columns_simple test failed: {e}")
        return None

def test_find_required_columns_simple(df):
    """Test the find_required_columns_simple function."""
    print("\n🔍 Testing find_required_columns_simple function...")
    
    try:
        # Import and test the function
        from karen_2_0_app import find_required_columns_simple
        
        # Mock st.write function for testing
        import karen_2_0_app
        mock_st = type('MockStreamlit', (), {
            'write': lambda x, *args, **kwargs: None,
            'warning': lambda x, *args, **kwargs: None,
            'error': lambda x, *args, **kwargs: None,
            'info': lambda x, *args, **kwargs: None,
            'success': lambda x, *args, **kwargs: None
        })()
        karen_2_0_app.st = mock_st
        
        required_cols = find_required_columns_simple(df)
        
        if required_cols and len(required_cols) >= 10:
            print(f"✅ Found {len(required_cols)} required columns")
            for col_letter, col_name in required_cols.items():
                print(f"  {col_letter}: {col_name}")
            return required_cols
        else:
            print(f"❌ Expected 10+ required columns, got {len(required_cols) if required_cols else 0}")
            return None
            
    except Exception as e:
        print(f"❌ find_required_columns_simple test failed: {e}")
        return None

def test_transaction_processing(df, ncb_columns, required_cols):
    """Test the transaction processing logic."""
    print("\n🔍 Testing transaction processing logic...")
    
    try:
        # Test transaction column handling
        transaction_col = required_cols.get('M')
        if not transaction_col:
            print("❌ No transaction column found")
            return False
        
        print(f"✅ Transaction column: {transaction_col}")
        
        # Test DataFrame vs Series handling
        if isinstance(df[transaction_col], pd.DataFrame):
            print("⚠️ Transaction column is DataFrame, converting to Series...")
            transaction_series = df[transaction_col].iloc[:, 0]
        else:
            transaction_series = df[transaction_col]
        
        print(f"✅ Transaction series type: {type(transaction_series)}")
        
        # Test unique values
        unique_transactions = transaction_series.astype(str).unique()
        print(f"✅ Unique transactions: {unique_transactions}")
        
        # Test filtering
        nb_mask = transaction_series.astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW', 'N'])
        c_mask = transaction_series.astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL', 'CANCELLED', 'CANC'])
        r_mask = transaction_series.astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE', 'REINSTATED'])
        
        print(f"✅ NB count: {nb_mask.sum()}")
        print(f"✅ C count: {c_mask.sum()}")
        print(f"✅ R count: {r_mask.sum()}")
        
        return True
        
    except Exception as e:
        print(f"❌ Transaction processing test failed: {e}")
        return False

def test_ncb_sum_calculation(df, ncb_columns):
    """Test NCB sum calculation."""
    print("\n🔍 Testing NCB sum calculation...")
    
    try:
        # Calculate NCB sum for each row
        ncb_cols = list(ncb_columns.values())
        
        if all(col in df.columns for col in ncb_cols):
            df['NCB_Sum'] = df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
            
            print(f"✅ NCB sum calculated for {len(df)} rows")
            print(f"✅ Sample NCB sums: {df['NCB_Sum'].head().tolist()}")
            
            # Test filtering logic
            nb_filtered = df[df['NCB_Sum'] > 0]
            c_filtered = df[df['NCB_Sum'] < 0]
            
            print(f"✅ NB filtered (sum > 0): {len(nb_filtered)}")
            print(f"✅ C filtered (sum < 0): {len(c_filtered)}")
            
            return True
        else:
            print(f"❌ Missing NCB columns: {[col for col in ncb_cols if col not in df.columns]}")
            return False
            
    except Exception as e:
        print(f"❌ NCB sum calculation test failed: {e}")
        return False

def test_output_creation(df, ncb_columns, required_cols):
    """Test output DataFrame creation."""
    print("\n🔍 Testing output DataFrame creation...")
    
    try:
        # Test column selection
        nb_output_cols = [
            required_cols.get('B'), required_cols.get('C'), required_cols.get('D'),
            required_cols.get('E'), required_cols.get('F'), required_cols.get('H'),
            required_cols.get('L'), required_cols.get('J'), required_cols.get('M'),
            required_cols.get('U'), ncb_columns.get('AO'), ncb_columns.get('AQ'),
            ncb_columns.get('AU'), ncb_columns.get('AW'), ncb_columns.get('AY'),
            ncb_columns.get('BA'), ncb_columns.get('BC')
        ]
        
        # Check for None values
        none_count = sum(1 for col in nb_output_cols if col is None)
        if none_count > 0:
            print(f"⚠️ {none_count} None values in column selection")
        
        # Check for duplicates
        nb_duplicates = [col for col in set(nb_output_cols) if nb_output_cols.count(col) > 1]
        if nb_duplicates:
            print(f"❌ Duplicate columns in selection: {nb_duplicates}")
            return False
        
        # Filter to available columns
        available_cols = [col for col in nb_output_cols if col is not None and col in df.columns]
        print(f"✅ Available columns for output: {len(available_cols)}")
        
        if len(available_cols) >= 17:
            print("✅ Sufficient columns for NB output")
            return True
        else:
            print(f"❌ Insufficient columns: {len(available_cols)} < 17")
            return False
            
    except Exception as e:
        print(f"❌ Output creation test failed: {e}")
        return False

def main():
    """Run all tests."""
    print("🧪 COMPREHENSIVE KAREN 2.0 TESTING SUITE")
    print("=" * 50)
    
    # Test 1: DataFrame creation
    df = test_dataframe_creation()
    if df is None:
        print("❌ CRITICAL: DataFrame creation failed. Stopping tests.")
        return False
    
    # Test 2: Duplicate column cleaning
    if not test_clean_duplicate_columns(df):
        print("❌ CRITICAL: Duplicate column cleaning failed. Stopping tests.")
        return False
    
    # Test 3: Empty column name cleaning
    if not test_clean_empty_column_names(df):
        print("❌ CRITICAL: Empty column name cleaning failed. Stopping tests.")
        return False
    
    # Test 4: NCB column finding
    ncb_columns = test_find_ncb_columns_simple(df)
    if ncb_columns is None:
        print("❌ CRITICAL: NCB column finding failed. Stopping tests.")
        return False
    
    # Test 5: Required column finding
    required_cols = test_find_required_columns_simple(df)
    if required_cols is None:
        print("❌ CRITICAL: Required column finding failed. Stopping tests.")
        return False
    
    # Test 6: Transaction processing
    if not test_transaction_processing(df, ncb_columns, required_cols):
        print("❌ CRITICAL: Transaction processing failed. Stopping tests.")
        return False
    
    # Test 7: NCB sum calculation
    if not test_ncb_sum_calculation(df, ncb_columns):
        print("❌ CRITICAL: NCB sum calculation failed. Stopping tests.")
        return False
    
    # Test 8: Output creation
    if not test_output_creation(df, ncb_columns, required_cols):
        print("❌ CRITICAL: Output creation failed. Stopping tests.")
        return False
    
    print("\n" + "=" * 50)
    print("🎉 ALL TESTS PASSED! Karen 2.0 is ready for deployment.")
    print("=" * 50)
    return True

if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)
