#!/bin/bash

# 🚀 KAREN 2.0 PROJECT - NEW COMPUTER SETUP SCRIPT
# This script will set up the project on your new, more powerful computer

echo "🎯 Setting up Karen 2.0 Project on new computer..."
echo "=================================================="

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python3 not found. Please install Python 3.8+ first."
    exit 1
fi

echo "✅ Python3 found: $(python3 --version)"

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 not found. Please install pip first."
    exit 1
fi

echo "✅ pip3 found: $(pip3 --version)"

# Install required packages
echo "📦 Installing required packages..."
pip3 install streamlit pandas openpyxl numpy

# Check if git is installed
if ! command -v git &> /dev/null; then
    echo "❌ Git not found. Please install Git first."
    exit 1
fi

echo "✅ Git found: $(git --version)"

# Clone the repository if not already present
if [ ! -d "Karenproject" ]; then
    echo "📥 Cloning repository..."
    git clone https://github.com/jeff99jackson99/Karenproject.git
    cd Karenproject
else
    echo "✅ Repository already exists, updating..."
    cd Karenproject
    git pull origin main
fi

echo ""
echo "🎉 Setup complete! Your Karen 2.0 project is ready."
echo ""
echo "📋 Next steps:"
echo "1. Place your Excel file in the project directory:"
echo "   - 2025-0731 Production Summary FINAL.xlsx"
echo ""
echo "2. Run the app:"
echo "   streamlit run karen_2_0_app.py"
echo ""
echo "3. Open your browser to the URL shown in the terminal"
echo ""
echo "4. Upload your Excel file and process it!"
echo ""
echo "📚 For detailed information, see PROJECT_SUMMARY.md"
echo "🔧 For troubleshooting, see the troubleshooting guide in PROJECT_SUMMARY.md"
echo ""
echo "🚀 Happy processing on your powerful new computer!"


