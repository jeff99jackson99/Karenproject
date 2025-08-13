#!/usr/bin/env python3
"""
NCB Data Processor
Processes New Business (NB), Cancellation (C), and Reinstatement (R) data from Excel spreadsheets
and generates filtered results based on specific criteria.
"""

import pandas as pd
import os
import sys
from datetime import datetime
import logging
import argparse

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ncb_processing.log'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

class NCBDataProcessor:
    """Main class for processing NCB data from Excel spreadsheets."""
    
    def __init__(self, input_file_path=None):
        self.input_file_path = input_file_path
        self.nb_data = None
        self.cancellation_data = None
        self.reinstatement_data = None
        
    def load_excel_file(self, file_path):
        """Load Excel file and return all sheets as a dictionary."""
        try:
            logger.info(f"Loading Excel file: {file_path}")
            excel_file = pd.ExcelFile(file_path)
            sheets = {}
            
            for sheet_name in excel_file.sheet_names:
                logger.info(f"Processing sheet: {sheet_name}")
                sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name)
                
            return sheets
        except Exception as e:
            logger.error(f"Error loading Excel file: {e}")
            raise
    
    def identify_transaction_types(self, df):
        """Identify transaction types in the dataframe."""
        transaction_columns = []
        for col in df.columns:
            col_lower = str(col).lower()
            if any(keyword in col_lower for keyword in ['transaction', 'type', 'trans_type', 'tran_type']):
                transaction_columns.append(col)
        return transaction_columns
    
    def filter_nb_data(self, df):
        """Filter data for New Business (NB) transactions."""
        logger.info("Filtering New Business (NB) data...")
        
        transaction_cols = self.identify_transaction_types(df)
        
        if not transaction_cols:
            logger.warning("No transaction type column found. Processing all data as potential NB.")
            return df
        
        nb_data = df.copy()
        for col in transaction_cols:
            if col in df.columns:
                mask = df[col].astype(str).str.upper().isin(['NB', 'NEW BUSINESS', 'NEW'])
                nb_data = df[mask]
                if len(nb_data) > 0:
                    logger.info(f"Found {len(nb_data)} NB records using column: {col}")
                    break
        
        return nb_data
    
    def filter_cancellation_data(self, df):
        """Filter data for Cancellation (C) transactions."""
        logger.info("Filtering Cancellation (C) data...")
        
        transaction_cols = self.identify_transaction_types(df)
        
        if not transaction_cols:
            logger.warning("No transaction type column found. Cannot filter cancellations.")
            return pd.DataFrame()
        
        cancellation_data = pd.DataFrame()
        for col in transaction_cols:
            if col in df.columns:
                mask = df[col].astype(str).str.upper().isin(['C', 'CANCELLATION', 'CANCEL'])
                temp_data = df[mask]
                if len(temp_data) > 0:
                    cancellation_data = temp_data
                    logger.info(f"Found {len(cancellation_data)} cancellation records using column: {col}")
                    break
        
        return cancellation_data
    
    def filter_reinstatement_data(self, df):
        """Filter data for Reinstatement (R) transactions."""
        logger.info("Filtering Reinstatement (R) data...")
        
        transaction_cols = self.identify_transaction_types(df)
        
        if not transaction_cols:
            logger.warning("No transaction type column found. Cannot filter reinstatements.")
            return pd.DataFrame()
        
        reinstatement_data = pd.DataFrame()
        for col in transaction_cols:
            if col in df.columns:
                mask = df[col].astype(str).str.upper().isin(['R', 'REINSTATEMENT', 'REINSTATE'])
                temp_data = df[mask]
                if len(temp_data) > 0:
                    reinstatement_data = temp_data
                    logger.info(f"Found {len(reinstatement_data)} reinstatement records using column: {col}")
                    break
        
        return reinstatement_data
    
    def check_ncb_amounts(self, df):
        """Check if the sum of Admin 3, 4, 9, and 10 amounts is greater than 0."""
        logger.info("Checking NCB amounts...")
        
        admin_columns = {
            'Admin 3 Amount': ['admin 3 amount', 'admin3 amount', 'admin_3_amount'],
            'Admin 4 Amount': ['admin 4 amount', 'admin4 amount', 'admin_4_amount'],
            'Admin 9 Amount': ['admin 9 amount', 'admin9 amount', 'admin_9_amount'],
            'Admin 10 Amount': ['admin 10 amount', 'admin10 amount', 'admin_10_amount']
        }
        
        found_columns = {}
        for admin_name, possible_names in admin_columns.items():
            for possible_name in possible_names:
                matching_cols = [col for col in df.columns if possible_name in str(col).lower()]
                if matching_cols:
                    found_columns[admin_name] = matching_cols[0]
                    break
        
        if len(found_columns) != 4:
            logger.warning(f"Only found {len(found_columns)} out of 4 required admin amount columns")
            logger.info(f"Found columns: {list(found_columns.keys())}")
            return df
        
        try:
            for col_name in found_columns.values():
                df[col_name] = pd.to_numeric(df[col_name], errors='coerce').fillna(0)
            
            total_amount = df[found_columns['Admin 3 Amount']] + \
                          df[found_columns['Admin 4 Amount']] + \
                          df[found_columns['Admin 9 Amount']] + \
                          df[found_columns['Admin 10 Amount']]
            
            filtered_df = df[total_amount > 0].copy()
            
            logger.info(f"Filtered from {len(df)} to {len(filtered_df)} records with NCB amounts > 0")
            return filtered_df
            
        except Exception as e:
            logger.error(f"Error calculating NCB amounts: {e}")
            return df
    
    def process_data(self, file_path):
        """Main method to process the Excel file."""
        logger.info("Starting NCB data processing...")
        
        try:
            sheets = self.load_excel_file(file_path)
            
            all_nb_data = []
            all_cancellation_data = []
            all_reinstatement_data = []
            
            for sheet_name, sheet_data in sheets.items():
                logger.info(f"Processing sheet: {sheet_name}")
                
                nb_data = self.filter_nb_data(sheet_data)
                if len(nb_data) > 0:
                    nb_data['Source_Sheet'] = sheet_name
                    all_nb_data.append(nb_data)
                
                cancellation_data = self.filter_cancellation_data(sheet_data)
                if len(cancellation_data) > 0:
                    cancellation_data['Source_Sheet'] = sheet_name
                    all_cancellation_data.append(cancellation_data)
                
                reinstatement_data = self.filter_reinstatement_data(sheet_data)
                if len(reinstatement_data) > 0:
                    reinstatement_data['Source_Sheet'] = sheet_name
                    all_reinstatement_data.append(reinstatement_data)
            
            if all_nb_data:
                self.nb_data = pd.concat(all_nb_data, ignore_index=True)
                self.nb_data = self.check_ncb_amounts(self.nb_data)
            
            if all_cancellation_data:
                self.cancellation_data = pd.concat(all_cancellation_data, ignore_index=True)
                self.cancellation_data = self.check_ncb_amounts(self.cancellation_data)
            
            if all_reinstatement_data:
                self.reinstatement_data = pd.concat(all_reinstatement_data, ignore_index=True)
                self.reinstatement_data = self.check_ncb_amounts(self.reinstatement_data)
            
            logger.info("Data processing completed successfully")
            
        except Exception as e:
            logger.error(f"Error processing data: {e}")
            raise
    
    def generate_output_files(self, output_dir="output"):
        """Generate output Excel files with processed data."""
        logger.info("Generating output files...")
        
        os.makedirs(output_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        if self.nb_data is not None and len(self.nb_data) > 0:
            nb_filename = f"{output_dir}/NB_Data_{timestamp}.xlsx"
            self.nb_data.to_excel(nb_filename, index=False, sheet_name="New Business Data")
            logger.info(f"Generated NB data file: {nb_filename}")
        
        if self.cancellation_data is not None and len(self.cancellation_data) > 0:
            cancellation_filename = f"{output_dir}/Cancellation_Data_{timestamp}.xlsx"
            self.cancellation_data.to_excel(cancellation_filename, index=False, sheet_name="Cancellation Data")
            logger.info(f"Generated cancellation data file: {cancellation_filename}")
        
        if self.reinstatement_data is not None and len(self.reinstatement_data) > 0:
            reinstatement_filename = f"{output_dir}/Reinstatement_Data_{timestamp}.xlsx"
            self.reinstatement_data.to_excel(reinstatement_filename, index=False, sheet_name="Reinstatement Data")
            logger.info(f"Generated reinstatement data file: {reinstatement_filename}")
        
        self.generate_summary_report(output_dir, timestamp)
        
        logger.info("Output files generated successfully")
    
    def generate_summary_report(self, output_dir, timestamp):
        """Generate a summary report of the processing results."""
        summary_data = {
            'Data Type': ['New Business', 'Cancellation', 'Reinstatement'],
            'Record Count': [
                len(self.nb_data) if self.nb_data is not None else 0,
                len(self.cancellation_data) if self.cancellation_data is not None else 0,
                len(self.reinstatement_data) if self.reinstatement_data is not None else 0
            ],
            'Processing Date': [datetime.now().strftime("%Y-%m-%d %H:%M:%S")] * 3
        }
        
        summary_df = pd.DataFrame(summary_data)
        summary_filename = f"{output_dir}/Processing_Summary_{timestamp}.xlsx"
        summary_df.to_excel(summary_filename, index=False, sheet_name="Summary")
        logger.info(f"Generated summary report: {summary_filename}")

def main():
    """Main function to run the NCB data processor."""
    parser = argparse.ArgumentParser(description='Process NCB data from Excel spreadsheets')
    parser.add_argument('input_file', help='Path to the input Excel file')
    parser.add_argument('--output-dir', default='output', help='Output directory for processed files')
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_file):
        logger.error(f"Input file not found: {args.input_file}")
        sys.exit(1)
    
    try:
        processor = NCBDataProcessor()
        processor.process_data(args.input_file)
        processor.generate_output_files(args.output_dir)
        logger.info("NCB data processing completed successfully!")
        
    except Exception as e:
        logger.error(f"Processing failed: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
