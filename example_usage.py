#!/usr/bin/env python3
"""
Example usage script for NCB Data Processor
Demonstrates how to use the processor programmatically.
"""

import os
import sys
from datetime import datetime

# Add current directory to path
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

def example_usage():
    """Example of how to use the NCB Data Processor."""
    print("NCB Data Processor - Example Usage")
    print("=" * 40)
    
    print("\n1. Command Line Usage:")
    print("   python main.py 'path/to/your/excel/file.xlsx' --output-dir output")
    
    print("\n2. GitHub Upload Usage:")
    print("   python github_uploader.py --token YOUR_TOKEN --repo owner/repo --path output/")
    
    print("\n3. Programmatic Usage:")
    print("   from main import NCBDataProcessor")
    print("   processor = NCBDataProcessor()")
    print("   processor.process_data('your_file.xlsx')")
    print("   processor.generate_output_files('output')")
    
    print("\n4. Required Excel Columns:")
    print("   - Transaction Type (NB, C, R)")
    print("   - Admin 3 Amount (Agent NCB)")
    print("   - Admin 4 Amount (Dealer NCB)")
    print("   - Admin 9 Amount (Agent NCB Offset)")
    print("   - Admin 10 Amount (Dealer NCB Offset)")
    
    print("\n5. Output Files Generated:")
    print("   - NB_Data_[timestamp].xlsx")
    print("   - Cancellation_Data_[timestamp].xlsx")
    print("   - Reinstatement_Data_[timestamp].xlsx")
    print("   - Processing_Summary_[timestamp].xlsx")
    
    print("\n6. Setup Steps:")
    print("   - Run: python setup.py")
    print("   - Copy .env.example to .env and configure")
    print("   - Place your Excel file in the project directory")
    print("   - Run the processor")

if __name__ == "__main__":
    example_usage()
