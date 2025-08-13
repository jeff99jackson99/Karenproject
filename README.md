# NCB Data Processor

A comprehensive Python application for processing New Business (NB), Cancellation (C), and Reinstatement (R) data from Excel spreadsheets and uploading results to GitHub.

## Features

- **Data Processing**: Automatically identifies and filters transaction types from Excel spreadsheets
- **NCB Amount Validation**: Ensures the sum of Admin 3, 4, 9, and 10 amounts is greater than 0
- **Multiple Output Formats**: Generates separate Excel files for each transaction type
- **GitHub Integration**: Uploads processed results directly to GitHub repositories
- **Comprehensive Logging**: Detailed logging for debugging and audit trails

## Requirements

- Python 3.7+
- pandas
- openpyxl
- PyGithub
- python-dotenv

## Installation

1. Clone or download this repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Set up GitHub authentication (see Configuration section)

## Configuration

### GitHub Setup

1. Create a GitHub Personal Access Token:
   - Go to GitHub Settings > Developer settings > Personal access tokens
   - Generate a new token with `repo` permissions
   - Copy the token

2. Create a `.env` file in the project directory:
   ```
   GITHUB_TOKEN=your_token_here
   GITHUB_REPO=owner/repository_name
   ```

## Usage

### Processing NCB Data

```bash
python main.py "path/to/your/excel/file.xlsx" --output-dir output
```

### Uploading to GitHub

```bash
python github_uploader.py --token YOUR_TOKEN --repo owner/repo --path output/
```

## Data Processing Logic

### New Business (NB) Transactions
- Filters for transaction type "NB" or "New Business"
- Extracts key fields: Insurer, Product Type, Coverage Code, Dealer Number, etc.
- Validates NCB amounts (Admin 3, 4, 9, 10)

### Cancellation (C) Transactions
- Filters for transaction type "C" or "Cancellation"
- Includes cancellation-specific fields: Cancellation Date, Reason, Factor
- Validates NCB amounts

### Reinstatement (R) Transactions
- Filters for transaction type "R" or "Reinstatement"
- Includes reinstatement-specific fields
- Validates NCB amounts

## Output Files

The processor generates:
- `NB_Data_[timestamp].xlsx` - New Business transactions
- `Cancellation_Data_[timestamp].xlsx` - Cancellation transactions
- `Reinstatement_Data_[timestamp].xlsx` - Reinstatement transactions
- `Processing_Summary_[timestamp].xlsx` - Summary report

## Logging

All operations are logged to:
- Console output
- `ncb_processing.log` file

## Error Handling

The application includes comprehensive error handling for:
- File loading issues
- Data processing errors
- GitHub upload failures
- Missing or invalid data

## Support

For issues or questions, please check the logs and ensure all dependencies are properly installed.
