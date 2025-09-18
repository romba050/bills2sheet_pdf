#!/usr/bin/env python3
"""
Receipt PDF to Google Sheets Script
Processes receipt PDFs using tabula-py for table extraction and uploads extracted data to Google Sheets

Copyright (C) 2025 Basile

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.
"""

import subprocess
import sys
import re
import argparse
from pathlib import Path
from typing import List, Tuple
import json
import csv
import pdfplumber

try:
    from googleapiclient.discovery import build
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
except ImportError:
    print("Error: Google API client libraries not installed.")
    print("""Install with:
          uv add google-api-python-client google-auth-httplib2 google-auth-oauthlib
          or
          pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib""")
    sys.exit(1)


# Google Sheets API scope
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

class ReceiptProcessor:
    def __init__(self, credentials_file: str = 'credentials.json', token_file: str = 'token.json'):
        """Initialize the receipt processor with Google Sheets API credentials."""
        self.credentials_file = credentials_file
        self.token_file = token_file
        self.service = None
        
    def authenticate_google_sheets(self) -> None:
        """Authenticate with Google Sheets API."""
        creds = None
        
        # Load existing token if available
        if Path(self.token_file).exists():
            creds = Credentials.from_authorized_user_file(self.token_file, SCOPES)
        
        # If no valid credentials, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not Path(self.credentials_file).exists():
                    print(f"Error: {self.credentials_file} not found.")
                    print("Download it from Google Cloud Console and place it in the script directory.")
                    sys.exit(1)
                
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save credentials for next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        
        self.service = build('sheets', 'v4', credentials=creds)
    
    def extract_table_from_pdf(self, pdf_path: str) -> List[List[str]]:
        """Extract table data from PDF using pdfplumber."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                all_tables = []

                for page in pdf.pages:
                    # Try to extract tables
                    tables = page.extract_tables()
                    if tables:
                        all_tables.extend(tables)

                    # If no tables found, try to extract text and parse it
                    if not tables:
                        text = page.extract_text()
                        if text:
                            # Parse text for receipt structure
                            parsed_table = self.parse_receipt_text(text)
                            if parsed_table:
                                all_tables.append(parsed_table)

                if not all_tables:
                    raise ValueError(f"No tables found in PDF: {pdf_path}")

                # Return the largest table (likely the main receipt table)
                return max(all_tables, key=len)

        except Exception as e:
            print(f"Error extracting table from PDF: {e}")
            sys.exit(1)

    def extract_betalat_total(self, pdf_path: str) -> str:
        """Extract the 'Betalat' (paid) total from the PDF."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]

                        for line in lines:
                            # Look for "Betalat" followed by amount
                            if line.startswith('Betalat '):
                                # Extract the amount after "Betalat "
                                amount_str = line.replace('Betalat ', '').strip()
                                # Convert comma to dot for decimal
                                amount_clean = amount_str.replace(',', '.')
                                # Validate it's a proper decimal number
                                if re.match(r'^\d+\.\d{2}$', amount_clean):
                                    return amount_clean

                # If not found, return empty string
                return ""

        except Exception as e:
            print(f"Error extracting Betalat total from PDF: {e}")
            return ""

    def extract_date_from_pdf(self, pdf_path: str) -> str:
        """Extract the date from the PDF (format: YYYY-MM-DD)."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]

                        for line in lines:
                            # Look for "Datum" followed by date
                            if 'Datum' in line:
                                # Extract date in YYYY-MM-DD format
                                date_match = re.search(r'(\d{4}-\d{2}-\d{2})', line)
                                if date_match:
                                    return date_match.group(1)

                # If not found, return empty string
                return ""

        except Exception as e:
            print(f"Error extracting date from PDF: {e}")
            return ""

    def parse_receipt_text(self, text: str) -> List[List[str]]:
        """Parse receipt text to extract structured data."""
        lines = [line.strip() for line in text.split('\n') if line.strip()]

        # Find the table section - look for lines that match receipt item pattern
        table_rows = []
        header_found = False

        for line in lines:
            # Look for the header line with "Beskrivning", "Pris", etc.
            if any(keyword in line for keyword in ['Beskrivning', 'Artikelnummer', 'Pris', 'Mängd', 'Summa']):
                # Split the header line into columns
                header_parts = re.split(r'\s{2,}', line)
                table_rows.append(header_parts)
                header_found = True
                continue

            # After header, look for item lines with pattern: name, article_number, price, quantity, total
            if header_found:
                # Stop at footer indicators
                if any(keyword in line.lower() for keyword in ['betalat', 'moms', 'kort', 'totalt', 'köp']):
                    break

                # Try to parse item lines - they typically have name + number + prices
                # Pattern: Item Name + Article Number + Unit Price + Quantity + Total Price
                parts = re.split(r'\s+', line)
                if len(parts) >= 4:
                    # Find the pattern: text + long_number + price + quantity + price
                    item_name = []
                    article_num = None
                    prices = []

                    for i, part in enumerate(parts):
                        if re.match(r'^\d{4,}$', part):  # Article number (4+ digits)
                            article_num = part
                            item_name = ' '.join(parts[:i])
                            remaining = parts[i+1:]
                            # Extract prices from remaining parts
                            for p in remaining:
                                if re.match(r'^\d+[,\.]\d{2}$', p):
                                    prices.append(p.replace(',', '.'))
                            break

                    if item_name and len(prices) >= 2:
                        # Format: [item_name, article_num, unit_price, quantity, total_price]
                        table_rows.append([item_name, article_num, prices[0], '1.00 st', prices[-1]])

        return table_rows if len(table_rows) > 1 else []

    def process_receipt_table(self, table: List[List[str]]) -> List[Tuple[str, str]]:
        """Process the extracted table to get item-price pairs."""
        items_and_prices = []

        if not table or len(table) < 2:
            return items_and_prices

        # Find header row and column indices
        header_row = table[0]
        desc_col_idx = 0  # Default to first column
        price_col_idx = -1  # Default to last column

        # Try to find specific columns
        for i, col in enumerate(header_row):
            col_lower = str(col).lower()
            if 'beskrivning' in col_lower or 'description' in col_lower:
                desc_col_idx = i
            elif 'summa' in col_lower:
                price_col_idx = i

        # Process data rows (skip header)
        for row in table[1:]:
            if len(row) <= max(desc_col_idx, abs(price_col_idx)):
                continue

            item_name = str(row[desc_col_idx]).strip()
            price_value = str(row[price_col_idx]).strip()

            # Skip empty or invalid rows
            if not item_name or not price_value or item_name in ['None', 'nan', '']:
                continue

            # Clean price value (remove currency symbols, convert comma to dot)
            price_clean = re.sub(r'[^\d,.-]', '', price_value)
            price_clean = price_clean.replace(',', '.')

            # Validate price format
            if re.match(r'^\d+\.\d{2}$', price_clean):
                # Escape items starting with '+' or '*'
                if item_name.startswith(('+', '*')):
                    item_name = "'" + item_name

                items_and_prices.append((item_name, price_clean))

        return items_and_prices

    def extract_items_and_prices(self, pdf_path: str) -> List[Tuple[str, str]]:
        """Extract items and their prices from the PDF."""
        table = self.extract_table_from_pdf(pdf_path)
        return self.process_receipt_table(table)
    
    def create_or_update_sheet(self, spreadsheet_id: str, sheet_name: str, data: List[Tuple[str, str]], pdf_path: str) -> None:
        """Create or update a Google Sheet with the extracted data."""
        if not self.service:
            raise RuntimeError("Google Sheets service not initialized. Call authenticate_google_sheets() first.")

        # Extract PDF total
        pdf_total = self.extract_betalat_total(pdf_path)

        # Prepare data for Google Sheets (add headers with expense tracking columns)
        sheet_data = [['Item', 'Shared expenses', 'My expenses', 'Jessica expenses', '', '', '']]

        # Convert data to expanded format (starting from row 2)
        for item, price in data:
            sheet_data.append([item, price, '', '', '', '', ''])

        # Add empty row for separation
        sheet_data.append(['', '', '', '', '', '', ''])

        # Add summary rows after the data with labels in column F and formulas in column G
        start_row = len(data) + 3  # +1 for header, +1 for empty row, +1 for 1-based indexing
        sheet_data.append(['', '', '', '', '', 'Sum of shared expenses', '=SUM(B:B)'])
        sheet_data.append(['', '', '', '', '', 'Sum of my expenses', '=SUM(C:C)'])
        sheet_data.append(['', '', '', '', '', "Sum of Jessica's expenses", '=SUM(D:D)'])
        sheet_data.append(['', '', '', '', '', 'Sheet total', f'=SUM(G{start_row}:G{start_row+2})'])
        sheet_data.append(['', '', '', '', '', 'PDF total', pdf_total])
        
        try:
            # Check if sheet exists, create if it doesn't
            spreadsheet = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheet_exists = any(sheet['properties']['title'] == sheet_name for sheet in spreadsheet['sheets'])
            
            if not sheet_exists:
                # Create the sheet
                request_body = {
                    'requests': [{
                        'addSheet': {
                            'properties': {
                                'title': sheet_name
                            }
                        }
                    }]
                }
                self.service.spreadsheets().batchUpdate(
                    spreadsheetId=spreadsheet_id,
                    body=request_body
                ).execute()
                print(f"Created new sheet: {sheet_name}")
            
            # Clear the sheet
            self.service.spreadsheets().values().clear(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A:Z"
            ).execute()
            
            # Add new data
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=f"{sheet_name}!A1",
                valueInputOption='USER_ENTERED',
                body={'values': sheet_data}
            ).execute()
            
            print(f"Successfully updated sheet '{sheet_name}' with {len(data)} items, expense tracking columns, and summary calculations.")
            
        except Exception as e:
            print(f"Error updating Google Sheet: {e}")
            sys.exit(1)
    
    def create_new_spreadsheet(self, title: str, data: List[Tuple[str, str]], pdf_path: str) -> str:
        """Create a new Google Spreadsheet with the extracted data."""
        if not self.service:
            raise RuntimeError("Google Sheets service not initialized. Call authenticate_google_sheets() first.")

        # Extract PDF total
        pdf_total = self.extract_betalat_total(pdf_path)

        # Prepare data for Google Sheets (add headers with expense tracking columns)
        sheet_data = [['Item', 'Shared expenses', 'My expenses', 'Jessica expenses', '', '', '']]

        # Convert data to expanded format (starting from row 2)
        for item, price in data:
            sheet_data.append([item, price, '', '', '', '', ''])

        # Add empty row for separation
        sheet_data.append(['', '', '', '', '', '', ''])

        # Add summary rows after the data with labels in column F and formulas in column G
        start_row = len(data) + 3  # +1 for header, +1 for empty row, +1 for 1-based indexing
        sheet_data.append(['', '', '', '', '', 'Sum of shared expenses', '=SUM(B:B)'])
        sheet_data.append(['', '', '', '', '', 'Sum of my expenses', '=SUM(C:C)'])
        sheet_data.append(['', '', '', '', '', "Sum of Jessica's expenses", '=SUM(D:D)'])
        sheet_data.append(['', '', '', '', '', 'Sheet total', f'=SUM(G{start_row}:G{start_row+2})'])
        sheet_data.append(['', '', '', '', '', 'PDF total', pdf_total])
        
        try:
            # Create new spreadsheet
            spreadsheet_body = {
                'properties': {
                    'title': title
                },
                'sheets': [{
                    'properties': {
                        'title': 'Receipt Items'
                    }
                }]
            }
            
            spreadsheet = self.service.spreadsheets().create(
                body=spreadsheet_body
            ).execute()
            
            spreadsheet_id = spreadsheet['spreadsheetId']
            
            # Add data to the new spreadsheet
            self.service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range='Receipt Items!A1',
                valueInputOption='USER_ENTERED',
                body={'values': sheet_data}
            ).execute()
            
            print(f"Successfully created new spreadsheet: {title} with expense tracking columns and summary calculations")
            print(f"Spreadsheet ID: {spreadsheet_id}")
            print(f"URL: https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit")
            
            return spreadsheet_id
            
        except Exception as e:
            print(f"Error creating Google Spreadsheet: {e}")
            sys.exit(1)
    
    def save_to_csv(self, data: List[Tuple[str, str]], csv_path: str) -> None:
        """Save the extracted data to a CSV file."""
        try:
            with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(['Item', 'Price'])
                writer.writerows(data)
            print(f"Successfully saved {len(data)} items to {csv_path}")
        except Exception as e:
            print(f"Error saving to CSV: {e}")
            sys.exit(1)
    
    def process_receipt(self, pdf_path: str, spreadsheet_id: str = None,
                       sheet_name: str = "Receipt Items", create_new: bool = False, csv_path: str = None) -> None:
        """Process a receipt PDF end-to-end."""
        print(f"Processing receipt: {pdf_path}")

        # Extract table from PDF
        print("Extracting table from PDF...")
        items_and_prices = self.extract_items_and_prices(pdf_path)
        
        if not items_and_prices:
            print("No items found in the receipt.")
            return
        
        print(f"Found {len(items_and_prices)} items:")
        for item, price in items_and_prices:
            print(f"  {item}: {price}")
        
        # Save to CSV or Google Sheets
        if csv_path:
            print("Saving to CSV...")
            self.save_to_csv(items_and_prices, csv_path)
        else:
            # Authenticate with Google Sheets
            print("Authenticating with Google Sheets...")
            self.authenticate_google_sheets()

            # Upload to Google Sheets
            print("Uploading to Google Sheets...")
            if create_new:
                receipt_name = Path(pdf_path).stem
                self.create_new_spreadsheet(f"Receipt - {receipt_name}", items_and_prices, pdf_path)
            else:
                if not spreadsheet_id:
                    print("Error: Spreadsheet ID required when not creating new spreadsheet.")
                    sys.exit(1)

                # Use PDF date as sheet name if default name was used
                final_sheet_name = sheet_name
                if sheet_name == "Receipt Items":
                    pdf_date = self.extract_date_from_pdf(pdf_path)
                    if pdf_date:
                        final_sheet_name = pdf_date
                        print(f"Using PDF date as sheet name: {final_sheet_name}")

                self.create_or_update_sheet(spreadsheet_id, final_sheet_name, items_and_prices, pdf_path)


def main():
    parser = argparse.ArgumentParser(description='Process receipt PDFs with table extraction and upload to Google Sheets')
    parser.add_argument('pdf_path', help='Path to the receipt PDF file')
    parser.add_argument('--spreadsheet-id', help='Google Sheets spreadsheet ID (required unless --create-new)')
    parser.add_argument('--sheet-name', default='Receipt Items', help='Name of the sheet to update (default: Receipt Items)')
    parser.add_argument('--create-new', action='store_true', help='Create a new spreadsheet instead of updating existing one')
    parser.add_argument('--credentials', default='credentials.json', help='Path to Google API credentials file')
    parser.add_argument('--token', default='token.json', help='Path to token file for storing authentication')
    parser.add_argument('--to-csv', help='Save extracted data to CSV file instead of Google Sheets')
    
    args = parser.parse_args()
    
    if not args.to_csv and not args.create_new and not args.spreadsheet_id:
        print("Error: Either --to-csv, --spreadsheet-id, or --create-new must be provided.")
        sys.exit(1)
    
    # Initialize processor
    processor = ReceiptProcessor(args.credentials, args.token)
    
    # Process the receipt
    processor.process_receipt(
        args.pdf_path,
        args.spreadsheet_id,
        args.sheet_name,
        args.create_new,
        args.to_csv
    )


if __name__ == "__main__":
    main()