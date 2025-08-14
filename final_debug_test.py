#!/usr/bin/env python3
"""
Final Debug Test for Karen 2.0 NCB Data Processor
Tests edge cases and ensures complete error-free operation
"""

import pandas as pd
import numpy as np
import sys
import os

# Add the current directory to Python path to import functions
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def test_edge_cases():
    """Test edge cases that could cause errors."""
    print("🧪 FINAL EDGE CASE TESTING")
    print("=" * 50)
    
    # Test 1: Empty DataFrame
    print("\n🔍 Test 1: Empty DataFrame handling")
    try:
        empty_df = pd.DataFrame()
        from karen_2_0_app import clean_duplicate_columns, clean_empty_column_names
        
        # Test duplicate cleaning on empty DataFrame
        result = clean_duplicate_columns(empty_df)
        if result is not None:
            print("✅ Empty DataFrame duplicate cleaning: PASSED")
        else:
            print("❌ Empty DataFrame duplicate cleaning: FAILED")
            
        # Test empty column cleaning on empty DataFrame
        result = clean_empty_column_names(empty_df)
        if result is not None:
            print("✅ Empty DataFrame empty column cleaning: PASSED")
        else:
            print("❌ Empty DataFrame empty column cleaning: FAILED")
            
    except Exception as e:
        print(f"❌ Empty DataFrame test failed: {e}")
    
    # Test 2: DataFrame with all NaN columns
    print("\n🔍 Test 2: All NaN columns handling")
    try:
        nan_df = pd.DataFrame({
            'col1': [np.nan, np.nan, np.nan],
            'col2': [np.nan, np.nan, np.nan]
        })
        
        result = clean_empty_column_names(nan_df)
        if result is not None:
            print("✅ All NaN columns cleaning: PASSED")
        else:
            print("❌ All NaN columns cleaning: FAILED")
            
    except Exception as e:
        print(f"❌ All NaN columns test failed: {e}")
    
    # Test 3: DataFrame with mixed data types
    print("\n🔍 Test 3: Mixed data types handling")
    try:
        mixed_df = pd.DataFrame({
            'text': ['a', 'b', 'c'],
            'number': [1, 2, 3],
            'float': [1.1, 2.2, 3.3],
            'bool': [True, False, True],
            'ADMIN 3 Amount': [100, 200, 300],
            'ADMIN 4 Amount': [150, 250, 350],
            'ADMIN 6 Amount': [50, 100, 150],
            'ADMIN 7 Amount': [25, 50, 75],
            'ADMIN 8 Amount': [75, 150, 225],
            'ADMIN 9 Amount': [125, 250, 375],
            'ADMIN 10 Amount': [175, 350, 525]
        })
        
        from karen_2_0_app import find_ncb_columns_simple, find_required_columns_simple
        
        # Mock Streamlit
        import karen_2_0_app
        mock_st = type('MockStreamlit', (), {
            'write': lambda x, *args, **kwargs: None,
            'warning': lambda x, *args, **kwargs: None,
            'error': lambda x, *args, **kwargs: None,
            'info': lambda x, *args, **kwargs: None,
            'success': lambda x, *args, **kwargs: None
        })()
        karen_2_0_app.st = mock_st
        
        # Test NCB column finding
        ncb_columns = find_ncb_columns_simple(mixed_df)
        if ncb_columns and len(ncb_columns) >= 7:
            print("✅ Mixed data types NCB finding: PASSED")
        else:
            print(f"❌ Mixed data types NCB finding: FAILED - found {len(ncb_columns) if ncb_columns else 0}")
            
        # Test required column finding
        required_cols = find_required_columns_simple(mixed_df)
        if required_cols and len(required_cols) >= 5:
            print("✅ Mixed data types required column finding: PASSED")
        else:
            print(f"❌ Mixed data types required column finding: FAILED - found {len(required_cols) if required_cols else 0}")
            
    except Exception as e:
        print(f"❌ Mixed data types test failed: {e}")
    
    # Test 4: Large DataFrame performance
    print("\n🔍 Test 4: Large DataFrame performance")
    try:
        # Create a larger DataFrame
        large_df = pd.DataFrame({
            'col_' + str(i): np.random.randn(1000) for i in range(50)
        })
        
        # Add Admin columns
        for i in range(3, 11):
            large_df[f'ADMIN {i} Amount'] = np.random.randn(1000)
        
        # Test performance
        import time
        start_time = time.time()
        
        result = clean_duplicate_columns(large_df)
        
        end_time = time.time()
        processing_time = end_time - start_time
        
        if result is not None and processing_time < 5.0:  # Should complete in under 5 seconds
            print(f"✅ Large DataFrame performance: PASSED ({processing_time:.2f}s)")
        else:
            print(f"❌ Large DataFrame performance: FAILED ({processing_time:.2f}s)")
            
    except Exception as e:
        print(f"❌ Large DataFrame test failed: {e}")
    
    # Test 5: Special characters in column names
    print("\n🔍 Test 5: Special characters in column names")
    try:
        special_df = pd.DataFrame({
            'Column with spaces': [1, 2, 3],
            'Column-with-dashes': [4, 5, 6],
            'Column_with_underscores': [7, 8, 9],
            'Column.with.dots': [10, 11, 12],
            'Column!@#$%^&*()': [13, 14, 15],
            'ADMIN 3 Amount': [100, 200, 300],
            'ADMIN 4 Amount': [150, 250, 350],
            'ADMIN 6 Amount': [50, 100, 150],
            'ADMIN 7 Amount': [25, 50, 75],
            'ADMIN 8 Amount': [75, 150, 225],
            'ADMIN 9 Amount': [125, 250, 375],
            'ADMIN 10 Amount': [175, 350, 525]
        })
        
        # Test NCB column finding with special characters
        ncb_columns = find_ncb_columns_simple(special_df)
        if ncb_columns and len(ncb_columns) >= 7:
            print("✅ Special characters NCB finding: PASSED")
        else:
            print(f"❌ Special characters NCB finding: FAILED - found {len(ncb_columns) if ncb_columns else 0}")
            
    except Exception as e:
        print(f"❌ Special characters test failed: {e}")
    
    # Test 6: Memory usage check
    print("\n🔍 Test 6: Memory usage check")
    try:
        import psutil
        import gc
        
        # Get initial memory usage
        process = psutil.Process()
        initial_memory = process.memory_info().rss / 1024 / 1024  # MB
        
        # Create and process a medium DataFrame
        medium_df = pd.DataFrame({
            'col_' + str(i): np.random.randn(100) for i in range(100)
        })
        
        # Add Admin columns
        for i in range(3, 11):
            medium_df[f'ADMIN {i} Amount'] = np.random.randn(100)
        
        # Process the DataFrame
        result = clean_duplicate_columns(medium_df)
        result = clean_empty_column_names(result)
        
        # Force garbage collection
        gc.collect()
        
        # Get final memory usage
        final_memory = process.memory_info().rss / 1024 / 1024  # MB
        memory_increase = final_memory - initial_memory
        
        if memory_increase < 100:  # Should not increase by more than 100MB
            print(f"✅ Memory usage: PASSED (increase: {memory_increase:.1f}MB)")
        else:
            print(f"⚠️ Memory usage: WARNING (increase: {memory_increase:.1f}MB)")
            
    except ImportError:
        print("⚠️ psutil not available, skipping memory test")
    except Exception as e:
        print(f"❌ Memory usage test failed: {e}")

def test_karen_2_0_compliance():
    """Test Karen 2.0 specification compliance."""
    print("\n🔍 KAREN 2.0 COMPLIANCE TEST")
    print("=" * 50)
    
    try:
        # Create test data that matches Karen 2.0 requirements
        test_data = {
            'Insurer Code': ['INS1', 'INS2', 'INS3', 'INS4', 'INS5'],
            'Product Type Code': ['PROD1', 'PROD2', 'PROD3', 'PROD4', 'PROD5'],
            'Coverage Code': ['COV1', 'COV2', 'COV3', 'COV4', 'COV5'],
            'Dealer Number': ['D1', 'D2', 'D3', 'D4', 'D5'],
            'Dealer Name': ['Dealer1', 'Dealer2', 'Dealer3', 'Dealer4', 'Dealer5'],
            'Contract Number': ['C1', 'C2', 'C3', 'C4', 'C5'],
            'Contract Sale Date': ['2025-01-01', '2025-01-02', '2025-01-03', '2025-01-04', '2025-01-05'],
            'Transaction Date': ['2025-01-01', '2025-01-02', '2025-01-03', '2025-01-04', '2025-01-05'],
            'Transaction Type': ['NB', 'C', 'R', 'NB', 'C'],
            'Customer Last Name': ['Smith', 'Jones', 'Brown', 'Wilson', 'Davis'],
            'Term Months': [36, 48, 60, 72, 84],
            'Cancellation Factor': [0.8, 0.9, 0.7, 0.6, 0.5],
            'Cancellation Reason': ['None', 'Customer Request', 'None', 'None', 'None'],
            'Cancellation Date': ['', '2025-01-15', '', '', ''],
            'ADMIN 3 Amount': [100.0, -50.0, 75.0, 200.0, -100.0],
            'ADMIN 4 Amount': [200.0, -100.0, 150.0, 400.0, -200.0],
            'ADMIN 6 Amount': [50.0, -25.0, 37.5, 100.0, -50.0],
            'ADMIN 7 Amount': [25.0, -12.5, 18.75, 50.0, -25.0],
            'ADMIN 8 Amount': [75.0, -37.5, 56.25, 150.0, -75.0],
            'ADMIN 9 Amount': [125.0, -62.5, 93.75, 250.0, -125.0],
            'ADMIN 10 Amount': [150.0, -75.0, 112.5, 300.0, -150.0]
        }
        
        df = pd.DataFrame(test_data)
        
        # Mock Streamlit
        import karen_2_0_app
        mock_st = type('MockStreamlit', (), {
            'write': lambda x, *args, **kwargs: None,
            'warning': lambda x, *args, **kwargs: None,
            'error': lambda x, *args, **kwargs: None,
            'info': lambda x, *args, **kwargs: None,
            'success': lambda x, *args, **kwargs: None
        })()
        karen_2_0_app.st = mock_st
        
        # Test NCB column finding
        from karen_2_0_app import find_ncb_columns_simple
        ncb_columns = find_ncb_columns_simple(df)
        
        if ncb_columns and len(ncb_columns) >= 7:
            print("✅ NCB column finding: PASSED")
            
            # Test transaction processing
            from karen_2_0_app import find_required_columns_simple
            required_cols = find_required_columns_simple(df)
            
            if required_cols and len(required_cols) >= 10:
                print("✅ Required column finding: PASSED")
                
                # Test transaction filtering
                transaction_col = required_cols.get('M')
                if transaction_col:
                    # Test NB filtering (sum > 0)
                    nb_mask = df[transaction_col].astype(str).str.upper().isin(['NB'])
                    nb_df = df[nb_mask]
                    
                    # Calculate NCB sum for NB transactions
                    ncb_cols = list(ncb_columns.values())
                    nb_df['NCB_Sum'] = nb_df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                    nb_filtered = nb_df[nb_df['NCB_Sum'] > 0]
                    
                    if len(nb_filtered) >= 1:
                        print("✅ NB filtering (sum > 0): PASSED")
                    else:
                        print("❌ NB filtering (sum > 0): FAILED")
                    
                    # Test C filtering (sum < 0)
                    c_mask = df[transaction_col].astype(str).str.upper().isin(['C'])
                    c_df = df[c_mask]
                    c_df['NCB_Sum'] = c_df[ncb_cols].apply(pd.to_numeric, errors='coerce').sum(axis=1)
                    c_filtered = c_df[c_df['NCB_Sum'] < 0]
                    
                    if len(c_filtered) >= 1:
                        print("✅ C filtering (sum < 0): PASSED")
                    else:
                        print("❌ C filtering (sum < 0): FAILED")
                        
                else:
                    print("❌ Transaction column not found")
            else:
                print(f"❌ Required column finding: FAILED - found {len(required_cols) if required_cols else 0}")
        else:
            print(f"❌ NCB column finding: FAILED - found {len(ncb_columns) if ncb_columns else 0}")
            
    except Exception as e:
        print(f"❌ Karen 2.0 compliance test failed: {e}")

def main():
    """Run all final debug tests."""
    print("🧪 FINAL COMPREHENSIVE DEBUG TESTING")
    print("=" * 60)
    
    # Test edge cases
    test_edge_cases()
    
    # Test Karen 2.0 compliance
    test_karen_2_0_compliance()
    
    print("\n" + "=" * 60)
    print("🎯 FINAL DEBUG TESTING COMPLETE")
    print("=" * 60)

if __name__ == "__main__":
    main()
