# Receipt PDF to Google Sheets

A Python script that extracts itemized data from receipt PDFs and uploads it to Google Sheets or saves to CSV files. Uses `pdfplumber` for table extraction from PDF receipts.

## Features

- **PDF Table Extraction**: Automatically extracts receipt tables from PDF files
- **Google Sheets Integration**: Upload extracted data directly to Google Sheets
- **CSV Export**: Save data to CSV files for local processing
- **Flexible Output**: Create new spreadsheets or update existing sheets
- **Receipt Format Support**: Handles Swedish ICA receipt format and similar structured receipts

## Installation

1. Clone or download this repository
2. Install dependencies using uv:

```bash
uv sync
```

Or install manually:
```bash
pip install pdfplumber google-api-python-client google-auth-httplib2 google-auth-oauthlib
```

## Setup for Google Sheets Integration

1. **Create a Google Cloud Project**:
   - Go to [Google Cloud Console](https://console.cloud.google.com/)
   - Create a new project or select an existing one

2. **Enable Google Sheets API**:
   - Navigate to "APIs & Services" > "Library"
   - Search for "Google Sheets API" and enable it

3. **Create Credentials**:
   - Go to "APIs & Services" > "Credentials"
   - Click "Create Credentials" > "OAuth 2.0 Client IDs"
   - Choose "Desktop application"
   - Download the credentials JSON file

4. **Setup Credentials**:
   - Save the downloaded file as `credentials.json` in the project directory
   - On first run, you'll be prompted to authorize the application

## Usage

### Basic Usage - Save to CSV

```bash
uv run python receipt_processor.py "path/to/receipt.pdf" --to-csv output.csv
```

### Upload to Existing Google Sheet

```bash
uv run python receipt_processor.py "path/to/receipt.pdf" --spreadsheet-id "your-sheet-id"
```

### Create New Google Spreadsheet

```bash
uv run python receipt_processor.py "path/to/receipt.pdf" --create-new
```

### Command Line Options

- `pdf_path`: Path to the receipt PDF file (required)
- `--spreadsheet-id`: Google Sheets spreadsheet ID (required unless using --create-new or --to-csv)
- `--sheet-name`: Name of the sheet to update (default: "Receipt Items")
- `--create-new`: Create a new spreadsheet instead of updating existing one
- `--credentials`: Path to Google API credentials file (default: credentials.json)
- `--token`: Path to token file for storing authentication (default: token.json)
- `--to-csv`: Save extracted data to CSV file instead of Google Sheets

## Examples

### Process ICA Receipt and Save to CSV

```bash
uv run python receipt_processor.py "bills/ICA_receipt.pdf" --to-csv "grocery_items.csv"
```

### Upload to Specific Google Sheet

```bash
uv run python receipt_processor.py "bills/receipt.pdf" --spreadsheet-id "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
```

### Create New Spreadsheet for Each Receipt

```bash
uv run python receipt_processor.py "bills/receipt_2025-09-17.pdf" --create-new
```

## Supported Receipt Formats

The script is designed to work with structured PDF receipts that contain:

- Item names/descriptions
- Prices in decimal format (xx.xx or xx,xx)
- Tabular layout or structured text format

Tested formats:
-  ICA Supermarket receipts (Swedish)
-  General European receipt formats with similar structure

## Output Format

The extracted data includes:
- **Item**: Product name/description
- **Price**: Final price in decimal format (converted from comma to dot notation)

Example output:
```csv
Item,Price
Gr�nk�lsblad ICA,39.95
Gul l�k ICA,1.90
Havregurt ugnsbaka,34.95
```

## File Structure

```
bills2sheet_pdf/
 receipt_processor.py    # Main script
 credentials.json        # Google API credentials (you provide)
 token.json             # OAuth token (auto-generated)
 bills/                 # Example PDF receipts
 README.md             # This file
 pyproject.toml        # Project dependencies
```

## How It Works

1. **PDF Processing**: Uses `pdfplumber` to extract tables from PDF files
2. **Text Parsing**: Falls back to text parsing if no structured tables are found
3. **Data Cleaning**: Normalizes prices, handles special characters, validates format
4. **Output**: Uploads to Google Sheets or saves to CSV based on user preference

## Troubleshooting

### Common Issues

**"No tables found in PDF"**
- Ensure the PDF contains structured tabular data
- Try with a different PDF or check if the receipt format is supported

**Google Sheets Authentication Errors**
- Verify `credentials.json` is in the project directory
- Check that Google Sheets API is enabled in your Google Cloud project
- Delete `token.json`, run the app again -  a brwoser window asking you to authenticate again should open

**Import/Module Errors**
- Ensure all dependencies are installed: `uv sync`
- Check that you're running with `uv run python` if using uv

### Debug Mode

For detailed debugging, you can modify the script to print extracted tables:

```python
# Add after table extraction
print("Extracted table:", table)
```

## License

This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with various PDF receipts
5. Submit a pull request
