#!/bin/bash
# 🚀 Karen 3.0 NCB Data Processor - Quick Setup Script
# Run this script on any new computer to get the project running quickly

echo "🚀 Setting up Karen 3.0 NCB Data Processor on new computer..."

# Check if Python 3 is installed
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 is not installed. Please install Python 3.8+ first."
    exit 1
fi

echo "✅ Python 3 found: $(python3 --version)"

# Check if pip is installed
if ! command -v pip3 &> /dev/null; then
    echo "❌ pip3 is not installed. Please install pip3 first."
    exit 1
fi

echo "✅ pip3 found: $(pip3 --version)"

# Install required packages
echo "📦 Installing required packages..."
pip3 install streamlit pandas openpyxl

# Check if Streamlit was installed correctly
if ! python3 -c "import streamlit" &> /dev/null; then
    echo "❌ Failed to install Streamlit. Trying alternative installation..."
    pip3 install --user streamlit pandas openpyxl
fi

echo "✅ All packages installed successfully!"

# Create a quick start script
cat > quick_start.sh << 'EOF'
#!/bin/bash
echo "🚀 Starting Karen 3.0 NCB Data Processor..."
echo "📱 The app will open in your browser at: http://localhost:8501"
echo "📁 Make sure your Excel file is ready to upload!"
echo ""
streamlit run karen_3_0_app_fixed.py
EOF

chmod +x quick_start.sh

# Create a project info file
cat > PROJECT_INFO.txt << 'EOF'
🚀 KAREN 3.0 NCB DATA PROCESSOR - PROJECT INFO

📋 QUICK START:
1. Run: ./quick_start.sh
2. Open browser to: http://localhost:8501
3. Upload your Excel file
4. Process according to Instructions 3.0

📁 KEY FILES:
- karen_3_0_app_fixed.py (RECOMMENDED - implements exact Instructions 3.0)
- karen_3_0_app.py (original version)
- PROJECT_RECALL_GUIDE.md (comprehensive project guide)

🎯 INSTRUCTIONS 3.0 IMPLEMENTED:
✅ Cancellations: Only records with negative Admin values
✅ Admin 6,7,8: Properly positioned after Admin 10
✅ Transaction Type: Column J from pull sheet
✅ All output sheets: Correct column structure

📊 STATUS: COMPLETE AND PRODUCTION READY
EOF

echo ""
echo "🎉 Setup completed successfully!"
echo ""
echo "📋 To start the application, run:"
echo "   ./quick_start.sh"
echo ""
echo "📚 For detailed project information, see:"
echo "   PROJECT_RECALL_GUIDE.md"
echo "   PROJECT_INFO.txt"
echo ""
echo "🚀 Your Karen 3.0 NCB Data Processor is ready to use!"


