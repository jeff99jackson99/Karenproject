#!/usr/bin/env python3
"""
Setup script for NCB Data Processor
Installs dependencies and sets up the environment.
"""

import subprocess
import sys
import os

def install_requirements():
    """Install required packages."""
    print("Installing required packages...")
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✓ Dependencies installed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ Failed to install dependencies: {e}")
        return False

def create_directories():
    """Create necessary directories."""
    print("Creating directories...")
    directories = ['output', 'logs', 'data']
    
    for directory in directories:
        os.makedirs(directory, exist_ok=True)
        print(f"✓ Created directory: {directory}")
    
    return True

def check_python_version():
    """Check if Python version is compatible."""
    print("Checking Python version...")
    version = sys.version_info
    
    if version.major >= 3 and version.minor >= 7:
        print(f"✓ Python {version.major}.{version.minor} is compatible")
        return True
    else:
        print(f"✗ Python {version.major}.{version.minor} is not compatible. Please use Python 3.7+")
        return False

def main():
    """Main setup function."""
    print("Setting up NCB Data Processor...")
    print("=" * 40)
    
    # Check Python version
    if not check_python_version():
        sys.exit(1)
    
    # Install dependencies
    if not install_requirements():
        print("\nPlease install dependencies manually:")
        print("pip install -r requirements.txt")
        sys.exit(1)
    
    # Create directories
    create_directories()
    
    print("\n" + "=" * 40)
    print("✓ Setup completed successfully!")
    print("\nNext steps:")
    print("1. Copy your Excel file to the project directory")
    print("2. Set up GitHub authentication in .env file")
    print("3. Run: python main.py your_file.xlsx")
    print("4. Upload results: python github_uploader.py --token YOUR_TOKEN --repo owner/repo --path output/")

if __name__ == "__main__":
    main()
