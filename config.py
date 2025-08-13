#!/usr/bin/env python3
"""
Configuration file for NCB Data Processor
Contains settings and constants for the application.
"""

import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# GitHub Configuration
GITHUB_TOKEN = os.getenv('GITHUB_TOKEN')
GITHUB_REPO = os.getenv('GITHUB_REPO', 'owner/repository')

# File Processing Configuration
DEFAULT_OUTPUT_DIR = 'output'
SUPPORTED_EXCEL_EXTENSIONS = ['.xlsx', '.xls']

# NCB Amount Column Names (case-insensitive matching)
NCB_ADMIN_COLUMNS = {
    'Admin 3 Amount': ['admin 3 amount', 'admin3 amount', 'admin_3_amount', 'agent ncb', 'agent ncb fee'],
    'Admin 4 Amount': ['admin 4 amount', 'admin4 amount', 'admin_4_amount', 'dealer ncb', 'dealer ncb fee'],
    'Admin 9 Amount': ['admin 9 amount', 'admin9 amount', 'admin_9_amount', 'agent ncb offset'],
    'Admin 10 Amount': ['admin 10 amount', 'admin10 amount', 'admin_10_amount', 'dealer ncb offset']
}

# Transaction Type Values
TRANSACTION_TYPES = {
    'NEW_BUSINESS': ['NB', 'NEW BUSINESS', 'NEW'],
    'CANCELLATION': ['C', 'CANCELLATION', 'CANCEL'],
    'REINSTATEMENT': ['R', 'REINSTATEMENT', 'REINSTATE']
}

# Required Fields for Each Transaction Type
REQUIRED_FIELDS = {
    'NEW_BUSINESS': [
        'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
        'Contract Number', 'Contract Sale Date', 'Transaction Date', 'Last Name', 'Contract Term'
    ],
    'CANCELLATION': [
        'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
        'Contract Number', 'Last Name', 'Contract Sale Date', 'Cancellation Date',
        'Cancellation Reason', 'Cancellation Factor', 'Number Days in force'
    ],
    'REINSTATEMENT': [
        'Insurer', 'Product Type', 'Coverage Code', 'Dealer Number', 'Dealer Name',
        'Contract Number', 'Last Name', 'Contract Sale Date', 'Reinstatement Date'
    ]
}

# Logging Configuration
LOG_LEVEL = 'INFO'
LOG_FORMAT = '%(asctime)s - %(levelname)s - %(message)s'
LOG_FILE = 'ncb_processing.log'

# Output File Templates
OUTPUT_FILE_TEMPLATES = {
    'NEW_BUSINESS': 'NB_Data_{timestamp}.xlsx',
    'CANCELLATION': 'Cancellation_Data_{timestamp}.xlsx',
    'REINSTATEMENT': 'Reinstatement_Data_{timestamp}.xlsx',
    'SUMMARY': 'Processing_Summary_{timestamp}.xlsx'
}
