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
from abc import ABC, abstractmethod

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


class StoreParser(ABC):
    """Abstract base class for store-specific receipt parsers."""

    @abstractmethod
    def parse_items(self, pdf_path: str) -> List[Tuple[str, str]]:
        """Extract items and their prices from the PDF.

        Returns:
            List of tuples (item_name, price)
        """
        pass

    @abstractmethod
    def extract_total(self, pdf_path: str) -> str:
        """Extract the total amount paid from the PDF.

        Returns:
            Total amount as string (e.g., "123.45")
        """
        pass

    @abstractmethod
    def extract_date(self, pdf_path: str) -> str:
        """Extract the date from the PDF.

        Returns:
            Date in YYYY-MM-DD format
        """
        pass


class ICAParser(StoreParser):
    """Parser for ICA receipts."""

    def parse_items(self, pdf_path: str) -> List[Tuple[str, str]]:
        """Extract items and their prices from ICA PDF."""
        table = self._extract_table_from_pdf(pdf_path)
        return self._process_receipt_table(table)

    def extract_total(self, pdf_path: str) -> str:
        """Extract the 'Betalat' (paid) total from ICA PDF."""
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

    def extract_date(self, pdf_path: str) -> str:
        """Extract the date from ICA PDF (format: YYYY-MM-DD)."""
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

    def _extract_table_from_pdf(self, pdf_path: str) -> List[List[str]]:
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
                            parsed_table = self._parse_receipt_text(text)
                            if parsed_table:
                                all_tables.append(parsed_table)

                if not all_tables:
                    raise ValueError(f"No tables found in PDF: {pdf_path}")

                # Return the largest table (likely the main receipt table)
                return max(all_tables, key=len)

        except Exception as e:
            print(f"Error extracting table from PDF: {e}")
            sys.exit(1)

    def _parse_receipt_text(self, text: str) -> List[List[str]]:
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
                if len(parts) >= 2:
                    # Check for reduction/discount lines (negative amounts)
                    negative_price = None
                    for part in parts:
                        if re.match(r'^-\d+[,\.]\d{2}$', part):  # Negative price
                            negative_price = part.replace(',', '.')
                            break

                    if negative_price:
                        # This is a reduction line - extract everything before the price as item name
                        item_parts = []
                        for part in parts:
                            if part == parts[-1] and re.match(r'^-\d+[,\.]\d{2}$', part):
                                break
                            item_parts.append(part)

                        if item_parts:
                            item_name = ' '.join(item_parts)
                            # Format: [item_name, no_article_num, unit_price, quantity, total_price]
                            table_rows.append([item_name, '', negative_price, '1.00 st', negative_price])
                            continue

                    # Standard item processing with article numbers
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

    def _process_receipt_table(self, table: List[List[str]]) -> List[Tuple[str, str]]:
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

            # Validate price format (allow negative prices for reductions)
            if re.match(r'^-?\d+\.\d{2}$', price_clean):
                items_and_prices.append((item_name, price_clean))

        return items_and_prices


class WillysParser(StoreParser):
    """Parser for Willy's receipts."""

    def parse_items(self, pdf_path: str) -> List[Tuple[str, str]]:
        """Extract items and their prices from Willy's PDF."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        items_and_prices = []

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]

                        in_items_section = False
                        i = 0
                        pending_item_name = None

                        while i < len(lines):
                            line = lines[i]

                            # Start of items section
                            if 'Start Självscanning' in line or 'Start självscanning' in line:
                                in_items_section = True
                                i += 1
                                continue

                            # End of items section
                            if 'Slut Självscanning' in line or 'Slut självscanning' in line:
                                in_items_section = False
                                break

                            if not in_items_section:
                                i += 1
                                continue

                            # Check if this is a weight calculation line (for multi-line items)
                            # Format: "0,140kg*499,00kr/kg 69,86"
                            if pending_item_name and re.match(r'^\d+[,\.]\d+kg\*', line):
                                # Extract price from the end of the calculation line
                                parts = line.split()
                                if parts and re.match(r'^\d+[,\.]\d{2}$', parts[-1]):
                                    price = parts[-1].replace(',', '.')
                                    items_and_prices.append((pending_item_name, price))
                                    pending_item_name = None
                                    i += 1
                                    continue

                            # Parse item lines
                            item_price = self._parse_willys_line(line)
                            if item_price:
                                items_and_prices.append(item_price)
                                pending_item_name = None
                            else:
                                # This might be a multi-line item (name only, price on next line)
                                # Check if the line looks like an item name (no price at the end)
                                if not re.search(r'\d+[,\.]\d{2}$', line):
                                    # Save it as a potential item name
                                    pending_item_name = line.strip()

                            i += 1

        except Exception as e:
            print(f"Error extracting items from Willy's PDF: {e}")
            sys.exit(1)

        return items_and_prices

    def _parse_willys_line(self, line: str) -> Tuple[str, str] | None:
        """Parse a single line from Willy's receipt."""
        # Skip empty lines
        if not line:
            return None

        # Handle discount/rabatt lines (indented with "Rabatt:" or "Prisnedsättning")
        if line.strip().startswith('Rabatt:') or line.strip().startswith('Prisnedsättning'):
            # Extract item name and negative price
            # Format: "Rabatt:ITEMNAME -XX,XX" or "Prisnedsättning XX,X% -XX,XX"
            parts = line.split()

            # Find the negative price (last element that matches pattern)
            price = None
            for part in reversed(parts):
                if re.match(r'^-\d+[,\.]\d{2}$', part):
                    price = part.replace(',', '.')
                    break

            if price:
                # Extract item name (everything between "Rabatt:" and the price)
                if 'Rabatt:' in line:
                    item_start = line.find('Rabatt:') + len('Rabatt:')
                    item_end = line.rfind(price)
                    item_name = line[item_start:item_end].strip()
                elif 'Prisnedsättning' in line:
                    # For price reductions, find the item name before the percentage
                    item_start = line.find('Prisnedsättning')
                    item_end = line.rfind(price)
                    # Extract everything after "Prisnedsättning XX,X%"
                    # Actually, for Prisnedsättning, there's no item name - just the reduction
                    item_name = line[:item_start].strip()
                    if not item_name:
                        item_name = "Prisnedsättning"

                return (item_name, price)

            return None

        # Handle pant (deposit) lines
        if line.strip().startswith('+PANT'):
            # Format: "+PANT ALUMINIUMBURK 1KR 10,00" or "+PANT ENG PET >1L 2,00"
            parts = line.split()
            # Price is the last element
            if parts:
                price_str = parts[-1]
                if re.match(r'^\d+[,\.]\d{2}$', price_str):
                    price = price_str.replace(',', '.')
                    # Item name is everything except the last element
                    item_name = ' '.join(parts[:-1])
                    return (item_name, price)
            return None

        # Regular item lines
        # Try to find a price at the end of the line
        parts = line.split()
        if not parts:
            return None

        # Look for price pattern at the end
        price = None
        price_idx = -1

        # Check if last element is a price
        if re.match(r'^\d+[,\.]\d{2}$', parts[-1]):
            price = parts[-1].replace(',', '.')
            price_idx = len(parts) - 1
        else:
            # Sometimes there's calculation info before the price
            # Format: "ITEM 4st*11,90 47,60" or "ITEM 0,140kg*499,00kr/kg 69,86"
            for i in range(len(parts) - 1, -1, -1):
                if re.match(r'^\d+[,\.]\d{2}$', parts[i]):
                    price = parts[i].replace(',', '.')
                    price_idx = i
                    break

        if price and price_idx > 0:
            # Item name is everything before the price and calculation info
            # Look for the start of calculation info (contains '*', 'kg', numbers with 'kg')
            item_end_idx = price_idx
            for i in range(price_idx - 1, -1, -1):
                part = parts[i]
                # Check if this looks like calculation info
                if ('*' in part or
                    part.endswith('kg') or
                    part.endswith('kg*') or
                    'kr/' in part or
                    re.match(r'^\d+[,\.]?\d*kg', part) or  # matches "0,140kg" or "0.140kg"
                    re.match(r'^\d+st\*', part)):
                    item_end_idx = i
                else:
                    # If not calculation, this is still part of the item name
                    break

            # Extract item name from start to where calculation begins
            item_parts = parts[:item_end_idx]

            if item_parts:
                item_name = ' '.join(item_parts)
                return (item_name, price)

        return None

    def extract_total(self, pdf_path: str) -> str:
        """Extract the 'Totalt' (total) amount from Willy's PDF."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        lines = [line.strip() for line in text.split('\n') if line.strip()]

                        for line in lines:
                            # Look for "Totalt" followed by amount and SEK
                            # Format: "Totalt 1043,88 SEK"
                            if line.startswith('Totalt ') and 'SEK' in line:
                                # Extract the amount between "Totalt " and " SEK"
                                amount_str = line.replace('Totalt ', '').replace(' SEK', '').strip()
                                # Convert comma to dot for decimal
                                amount_clean = amount_str.replace(',', '.')
                                # Validate it's a proper decimal number
                                if re.match(r'^\d+\.\d{2}$', amount_clean):
                                    return amount_clean

                # If not found, return empty string
                return ""

        except Exception as e:
            print(f"Error extracting Totalt from Willy's PDF: {e}")
            return ""

    def extract_date(self, pdf_path: str) -> str:
        """Extract the date from Willy's PDF (format: YYYY-MM-DD)."""
        if not Path(pdf_path).exists():
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")

        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        # Look for date in format YYYY-MM-DD HH:MM
                        date_match = re.search(r'(\d{4}-\d{2}-\d{2})\s+\d{2}:\d{2}', text)
                        if date_match:
                            return date_match.group(1)

                # If not found, return empty string
                return ""

        except Exception as e:
            print(f"Error extracting date from Willy's PDF: {e}")
            return ""


class ReceiptProcessor:
    def __init__(self, store_parser: StoreParser, store_name: str, credentials_file: str = 'credentials.json', token_file: str = 'token.json'):
        """Initialize the receipt processor with Google Sheets API credentials."""
        self.store_parser = store_parser
        self.store_name = store_name
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
    
    def create_or_update_sheet(self, spreadsheet_id: str, sheet_name: str, data: List[Tuple[str, str]], pdf_path: str) -> None:
        """Create or update a Google Sheet with the extracted data."""
        if not self.service:
            raise RuntimeError("Google Sheets service not initialized. Call authenticate_google_sheets() first.")

        # Extract PDF total
        pdf_total = self.store_parser.extract_total(pdf_path)

        # Prepare data for Google Sheets (add headers with expense tracking columns)
        sheet_data = [['Item', 'Shared expenses', 'My expenses', 'Jessica expenses', '', '', '']]

        # Add summary rows before the data with labels in column F and formulas in column G
        sheet_data.append(['', '', '', '', '', 'Sum of shared expenses', f'=SUM(B:B)'])
        sheet_data.append(['', '', '', '', '', 'Sum of my expenses', f'=SUM(C:C)'])
        sheet_data.append(['', '', '', '', '', "Sum of Jessica's expenses", f'=SUM(D:D)'])
        sheet_data.append(['', '', '', '', '', 'Sheet total', '=SUM(G2:G4)'])
        sheet_data.append(['', '', '', '', '', 'PDF total', pdf_total])

        # Add empty row for separation
        sheet_data.append(['', '', '', '', '', '', ''])

        # Convert data to expanded format
        for item, price in data:
            # Remove leading special characters that might cause issues in Google Sheets
            # Strip quotes first (from old escaping logic), then strip problematic characters
            clean_item = item.lstrip("'").lstrip('+*=@-')
            sheet_data.append([clean_item, price, '', '', '', '', ''])
        
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
        pdf_total = self.store_parser.extract_total(pdf_path)

        # Prepare data for Google Sheets (add headers with expense tracking columns)
        sheet_data = [['Item', 'Shared expenses', 'My expenses', 'Jessica expenses', '', '', '']]

        # Add summary rows before the data with labels in column F and formulas in column G
        data_start_row = 7  # Row where actual item data will start (1-based)
        data_end_row = data_start_row + len(data) - 1
        sheet_data.append(['', '', '', '', '', 'Sum of shared expenses', f'=SUM(B{data_start_row}:B{data_end_row})'])
        sheet_data.append(['', '', '', '', '', 'Sum of my expenses', f'=SUM(C{data_start_row}:C{data_end_row})'])
        sheet_data.append(['', '', '', '', '', "Sum of Jessica's expenses", f'=SUM(D{data_start_row}:D{data_end_row})'])
        sheet_data.append(['', '', '', '', '', 'Sheet total', '=SUM(G2:G4)'])
        sheet_data.append(['', '', '', '', '', 'PDF total', pdf_total])

        # Add empty row for separation
        sheet_data.append(['', '', '', '', '', '', ''])

        # Convert data to expanded format
        for item, price in data:
            # Remove leading special characters that might cause issues in Google Sheets
            # Strip quotes first (from old escaping logic), then strip problematic characters
            clean_item = item.lstrip("'").lstrip('+*=@-')
            sheet_data.append([clean_item, price, '', '', '', '', ''])
        
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

        # Extract items from PDF
        print("Extracting items from PDF...")
        items_and_prices = self.store_parser.parse_items(pdf_path)
        
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
                    pdf_date = self.store_parser.extract_date(pdf_path)
                    if pdf_date:
                        final_sheet_name = f"{pdf_date}-{self.store_name}"
                        print(f"Using PDF date as sheet name: {final_sheet_name}")

                self.create_or_update_sheet(spreadsheet_id, final_sheet_name, items_and_prices, pdf_path)


def main():
    parser = argparse.ArgumentParser(description='Process receipt PDFs with table extraction and upload to Google Sheets')
    parser.add_argument('pdf_path', help='Path to the receipt PDF file')
    parser.add_argument('--store', required=True, choices=['ICA', 'WILLYS'], help='Store type (ICA or WILLYS)')
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

    # Create the appropriate store parser
    if args.store == 'ICA':
        store_parser = ICAParser()
    elif args.store == 'WILLYS':
        store_parser = WillysParser()
    else:
        print(f"Error: Unknown store type '{args.store}'")
        sys.exit(1)

    # Initialize processor
    processor = ReceiptProcessor(store_parser, args.store, args.credentials, args.token)

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