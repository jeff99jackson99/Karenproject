# NCB Data Processor - Complete Usage Guide

## Quick Start

1. **Setup** (already completed):
   ```bash
   python3 setup.py
   ```

2. **Configure GitHub** (optional):
   - Copy `.env.example` to `.env`
   - Add your GitHub token and repository

3. **Process your Excel file**:
   ```bash
   python3 main.py "path/to/your/excel/file.xlsx"
   ```

4. **Upload to GitHub** (if configured):
   ```bash
   python3 github_uploader.py --token YOUR_TOKEN --repo owner/repo --path output/
   ```

## Complete Workflow (One Command)

```bash
python3 workflow.py "path/to/your/excel/file.xlsx" --github-token YOUR_TOKEN --github-repo owner/repo
```

## What the Program Does

### 1. Data Processing
- **Loads** your Excel file (all sheets)
- **Identifies** transaction types (NB, C, R)
- **Filters** data based on your criteria
- **Validates** NCB amounts (Admin 3+4+9+10 > 0)

### 2. Output Generation
- **NB_Data_[timestamp].xlsx** - New Business transactions
- **Cancellation_Data_[timestamp].xlsx** - Cancellation transactions  
- **Reinstatement_Data_[timestamp].xlsx** - Reinstatement transactions
- **Processing_Summary_[timestamp].xlsx** - Summary report

### 3. GitHub Integration
- **Uploads** processed files to your repository
- **Organizes** files in a `data/` folder
- **Tracks** changes with commit messages

## Required Excel Structure

Your Excel file should contain these columns (names can vary):

### Transaction Type Column
- Look for: "Transaction Type", "Type", "Trans Type"
- Values: "NB", "C", "R" (or similar)

### NCB Amount Columns
- **Admin 3 Amount**: Agent NCB or Agent NCB Fee
- **Admin 4 Amount**: Dealer NCB or Dealer NCB Fee  
- **Admin 9 Amount**: Agent NCB Offset
- **Admin 10 Amount**: Dealer NCB Offset

### Required Fields (per transaction type)
- Insurer, Product Type, Coverage Code
- Dealer Number, Dealer Name, Contract Number
- Contract Sale Date, Transaction Date, Last Name
- Contract Term (for NB)

## Example Commands

### Basic Processing
```bash
# Process Excel file
python3 main.py "2025-0731 Production Summary FINAL.xlsx"

# Process with custom output directory
python3 main.py "your_file.xlsx" --output-dir "my_results"
```

### GitHub Upload
```bash
# Upload single file
python3 github_uploader.py --token abc123 --repo jeff/insurance-data --path output/NB_Data_20250812_143022.xlsx

# Upload entire output directory
python3 github_uploader.py --token abc123 --repo jeff/insurance-data --path output/
```

### Complete Workflow
```bash
# Process and upload in one command
python3 workflow.py "2025-0731 Production Summary FINAL.xlsx" --github-token abc123 --github-repo jeff/insurance-data
```

## Troubleshooting

### Common Issues
1. **"No transaction type column found"**
   - Check your Excel headers for transaction type information
   - The program looks for columns containing "transaction", "type", etc.

2. **"Only found X out of 4 required admin amount columns"**
   - Ensure your Excel has Admin 3, 4, 9, and 10 amount columns
   - Column names can vary (e.g., "Admin 3 Amount", "admin3 amount")

3. **"GitHub upload failed"**
   - Verify your token has `repo` permissions
   - Check repository name format: `owner/repository`

### Getting Help
- Check the logs in `ncb_processing.log`
- Run `python3 test_processor.py` to test basic functionality
- Review the README.md for detailed documentation

## Advanced Usage

### Custom Configuration
Edit `config.py` to modify:
- Column name patterns
- Transaction type values
- Required field lists

### Programmatic Usage
```python
from main import NCBDataProcessor

processor = NCBDataProcessor()
processor.process_data("your_file.xlsx")
processor.generate_output_files("output")
```

## File Structure
```
ncb_data_processor/
├── main.py              # Main processing logic
├── github_uploader.py   # GitHub integration
├── workflow.py          # Complete workflow script
├── config.py            # Configuration settings
├── setup.py             # Installation script
├── test_processor.py    # Testing script
├── requirements.txt     # Python dependencies
├── README.md            # Documentation
├── USAGE_GUIDE.md      # This file
├── .env.example         # Environment template
├── output/              # Generated files
├── logs/                # Log files
└── data/                # Data directory
```

## Support

For issues or questions:
1. Check the logs in `ncb_processing.log`
2. Run the test script: `python3 test_processor.py`
3. Review error messages for specific guidance
4. Ensure all dependencies are installed: `python3 setup.py`
