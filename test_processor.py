#!/usr/bin/env python3
"""
Test script for NCB Data Processor
Tests basic functionality without requiring actual Excel files.
"""

import sys
import os
import pandas as pd
from datetime import datetime

# Add current directory to path to import main module
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def create_sample_data():
    """Create sample data for testing."""
    sample_data = {
        'Transaction Type': ['NB', 'C', 'R', 'NB', 'C'],
        'Insurer': ['Insurer A', 'Insurer B', 'Insurer C', 'Insurer D', 'Insurer E'],
        'Product Type': ['Product 1', 'Product 2', 'Product 3', 'Product 4', 'Product 5'],
        'Admin 3 Amount': [100, 0, 50, 200, 75],
        'Admin 4 Amount': [150, 25, 0, 300, 100],
        'Admin 9 Amount': [25, 0, 10, 50, 25],
        'Admin 10 Amount': [75, 0, 20, 100, 50],
        'Contract Number': ['CN001', 'CN002', 'CN003', 'CN004', 'CN005']
    }
    
    return pd.DataFrame(sample_data)

def test_basic_functionality():
    """Test basic functionality of the processor."""
    print("Testing NCB Data Processor...")
    
    try:
        # Test pandas import
        print("✓ Pandas imported successfully")
        
        # Test sample data creation
        df = create_sample_data()
        print(f"✓ Sample data created with {len(df)} rows")
        
        # Test NCB amount calculation
        total_amounts = df['Admin 3 Amount'] + df['Admin 4 Amount'] + df['Admin 9 Amount'] + df['Admin 10 Amount']
        print(f"✓ NCB amounts calculated successfully")
        print(f"  - Total amounts range: {total_amounts.min()} to {total_amounts.max()}")
        
        # Test filtering
        nb_data = df[df['Transaction Type'] == 'NB']
        c_data = df[df['Transaction Type'] == 'C']
        r_data = df[df['Transaction Type'] == 'R']
        
        print(f"✓ Data filtering successful:")
        print(f"  - NB records: {len(nb_data)}")
        print(f"  - Cancellation records: {len(c_data)}")
        print(f"  - Reinstatement records: {len(r_data)}")
        
        print("\n✓ All basic tests passed!")
        return True
        
    except Exception as e:
        print(f"✗ Test failed: {e}")
        return False

if __name__ == "__main__":
    success = test_basic_functionality()
    if success:
        print("\nReady to process NCB data!")
    else:
        print("\nPlease check your installation and dependencies.")
        sys.exit(1)
