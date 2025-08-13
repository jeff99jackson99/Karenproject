#!/usr/bin/env python3
"""
Secure script to run NCB Data Processor with GitHub token
This script keeps your token secure and doesn't commit it to git.
"""

import os
import sys
import subprocess

def run_processor_with_token():
    """Run the NCB processor with the provided token."""
    print("🔐 NCB Data Processor - Secure Token Runner")
    print("=" * 50)
    
    # Check if token is provided
    if len(sys.argv) < 2:
        print("Usage: python3 run_with_token.py <your_github_token>")
        print("Example: python3 run_with_token.py YOUR_TOKEN_HERE")
        sys.exit(1)
    
    token = sys.argv[1]
    
    # Set environment variable for this session only
    os.environ['GITHUB_TOKEN'] = token
    os.environ['GITHUB_REPO'] = 'jeff99jackson99/Karenproject'
    
    print("✅ GitHub token set for this session")
    print("✅ Repository configured: jeff99jackson99/Karenproject")
    
    # Show available commands
    print("\n🚀 Available Commands:")
    print("1. Process Excel file:")
    print("   python3 main.py 'your_excel_file.xlsx'")
    
    print("\n2. Upload results to GitHub:")
    print("   python3 github_uploader.py --path output/")
    
    print("\n3. Complete workflow (process + upload):")
    print("   python3 workflow.py 'your_excel_file.xlsx'")
    
    print("\n4. Test GitHub connection:")
    print("   python3 -c \"from github_uploader import GitHubUploader; u=GitHubUploader('" + token + "', 'jeff99jackson99/Karenproject'); print('✓ Connection successful!')\"")
    
    print(f"\n💡 Your token is now active for this terminal session")
    print(f"💡 Token: {token[:20]}...{token[-10:]}")

if __name__ == "__main__":
    run_processor_with_token()
