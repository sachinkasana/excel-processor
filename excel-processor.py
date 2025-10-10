#!/usr/bin/env python3
"""
Excel Data Processing and Filtering Script
==========================================

This script processes two Excel files (2A and 2B) containing supplier data and creates
a filtered output file based on matching GSTIN + Invoice Number combinations.

Requirements:
- pandas
- openpyxl

Usage:
    python excel-processor.py

The script expects:
- 2A.xlsx (source file)
- 2B.xlsx (reference file)

Output:
- Filtered Excel file with naming format: "Legal Name_Tax Period_Year_juric.xlsx"
- Processing summary text file

Author: Generated for Excel data processing task
Date: October 2025
"""

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import os
import sys


class ExcelProcessor:
    """Class to handle Excel file processing and filtering operations"""
    
    def __init__(self, file_2A="2A.xlsx", file_2B="2B.xlsx"):
        """
        Initialize the processor with file paths
        
        Args:
            file_2A (str): Path to 2A Excel file
            file_2B (str): Path to 2B Excel file
        """
        self.file_2A = file_2A
        self.file_2B = file_2B
        self.data_2A = {}
        self.data_2B = {}
        self.filtered_results = {}
        
        # Required sheets to process
        self.sheets_2A = ["B2B – Invoice & Rate wise", "B2BA", "CDNR", "CDNRA"]
        self.sheets_2B = ["B2B", "B2BA", "B2B-CDNR", "B2B-CDNRA"]
        
        # Extracted metadata
        self.legal_name = None
        self.tax_period = None
        self.current_year = "2025"
        self.fixed_suffix = "juric"
        
    def validate_files(self):
        """Check if input files exist"""
        if not os.path.exists(self.file_2A):
            raise FileNotFoundError(f"File {self.file_2A} not found")
        if not os.path.exists(self.file_2B):
            raise FileNotFoundError(f"File {self.file_2B} not found")
        print("✓ Input files validated")
    
    def extract_metadata(self):
        """Extract Legal Name and Tax Period from 2B Read me sheet"""
        try:
            read_me_2B = pd.read_excel(self.file_2B, sheet_name='Read me', header=None)
            
            for i in range(len(read_me_2B)):
                row_values = read_me_2B.iloc[i].tolist()
                for j, value in enumerate(row_values):
                    if pd.notna(value) and isinstance(value, str):
                        if 'Legal Name' in value:
                            if j + 2 < len(row_values) and pd.notna(row_values[j + 2]):
                                self.legal_name = str(row_values[j + 2]).strip()
                        elif 'Tax Period' in value:
                            if j + 2 < len(row_values) and pd.notna(row_values[j + 2]):
                                self.tax_period = str(row_values[j + 2]).strip()
            
            print(f"✓ Extracted metadata - Legal Name: {self.legal_name}, Tax Period: {self.tax_period}")
            
        except Exception as e:
            print(f"Warning: Could not extract metadata from 2B Read me sheet: {str(e)}")
            self.legal_name = "Unknown"
            self.tax_period = "Unknown"
    
    def read_sheet_advanced(self, file_path, sheet_name, is_2B=False):
        """
        Advanced sheet reading that handles different sheet types and multi-row headers
        
        Args:
            file_path (str): Path to Excel file
            sheet_name (str): Name of sheet to read
            is_2B (bool): Whether this is from 2B file (affects column mapping)
            
        Returns:
            pd.DataFrame: Processed dataframe with proper headers
        """
        try:
            df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
            
            # Find header rows based on sheet type
            if 'CDNR' in sheet_name or 'CDN' in sheet_name:
                header_keywords = ['gstin/uin', 'gstin of supplier']
            else:
                header_keywords = ['gstin of supplier']
            
            header_start_row = None
            for i in range(min(15, len(df_raw))):
                row_str = ' '.join([str(x) for x in df_raw.iloc[i].tolist() if pd.notna(x)]).lower()
                for keyword in header_keywords:
                    if keyword in row_str:
                        header_start_row = i
                        break
                if header_start_row is not None:
                    break
            
            if header_start_row is None:
                print(f"Warning: Could not find header row for sheet {sheet_name}")
                return pd.DataFrame()
            
            # Create proper column names from header rows
            header_row1 = df_raw.iloc[header_start_row].tolist()
            header_row2 = df_raw.iloc[header_start_row + 1].tolist() if header_start_row + 1 < len(df_raw) else [None] * len(header_row1)
            
            combined_headers = []
            for i in range(len(header_row1)):
                h1 = str(header_row1[i]).strip() if pd.notna(header_row1[i]) else ""
                h2 = str(header_row2[i]).strip() if pd.notna(header_row2[i]) else ""
                
                if h2 and h2 not in ['nan', 'NaN'] and h2 != h1:
                    combined_headers.append(h2)
                elif h1:
                    combined_headers.append(h1)
                else:
                    combined_headers.append(f"Column_{i}")
            
            # Read data starting from after headers
            data_start_row = header_start_row + 2
            df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=data_start_row, header=None)
            
            # Assign column names
            if len(combined_headers) <= len(df.columns):
                df.columns = combined_headers[:len(df.columns)]
            else:
                df.columns = combined_headers[:len(df.columns)]
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            return df
            
        except Exception as e:
            print(f"Error reading {sheet_name}: {str(e)}")
            return pd.DataFrame()
    
    def map_columns_properly(self, df, is_2B=False):
        """
        Map columns to proper names based on position and file type
        
        Args:
            df (pd.DataFrame): DataFrame to process
            is_2B (bool): Whether this is from 2B file
            
        Returns:
            pd.DataFrame: DataFrame with properly mapped columns
        """
        if is_2B:
            # Column mapping for 2B files
            column_mapping = {
                0: 'GSTIN of supplier',
                1: 'Trade/Legal name',
                2: 'Invoice number',
                3: 'Invoice type',
                4: 'Invoice Date',
                5: 'Invoice Value(₹)',
                6: 'Place of supply',
                7: 'Supply Attract Reverse Charge',
                8: 'Taxable Value (₹)',
                9: 'Integrated Tax(₹)',
                10: 'Central Tax(₹)',
                11: 'State/UT Tax(₹)',
                12: 'Cess(₹)',
                13: 'GSTR-1/IFF/GSTR-5 Period',
                14: 'GSTR-1/IFF/GSTR-5 Filing Date',
                15: 'ITC Availability',
                16: 'Reason',
                17: 'Applicable % of Tax Rate',
                18: 'Source',
                19: 'IRN',
                20: 'IRN Date'
            }
        else:
            # Column mapping for 2A files
            column_mapping = {
                0: 'GSTIN of supplier',
                1: 'Trade/Legal name of the Supplier',
                2: 'Invoice number',
                3: 'Invoice type',
                4: 'Invoice Date',
                5: 'Invoice Value (₹)',
                6: 'Place of supply',
                7: 'Supply Attract Reverse Charge',
                8: 'Rate (%)',
                9: 'Taxable Value (₹)',
                10: 'Integrated Tax  (₹)',
                11: 'Central Tax (₹)',
                12: 'State/UT tax (₹)',
                13: 'Cess  (₹)',
                14: 'Counter Party Return status',
                15: 'Filing Period',
                16: 'Notes'
            }
        
        # Apply mapping
        new_columns = []
        for i, col in enumerate(df.columns):
            if i in column_mapping:
                new_columns.append(column_mapping[i])
            else:
                new_columns.append(f"Unknown_Col_{i}")
        
        df.columns = new_columns
        return df
    
    def read_all_sheets(self):
        """Read all required sheets from both files"""
        print("Reading sheets from 2A...")
        for sheet_name in self.sheets_2A:
            df = self.read_sheet_advanced(self.file_2A, sheet_name, is_2B=False)
            if not df.empty:
                df = self.map_columns_properly(df, is_2B=False)
            self.data_2A[sheet_name] = df
            print(f"  {sheet_name}: {df.shape[0]} records")
        
        print("\nReading sheets from 2B...")
        for sheet_name in self.sheets_2B:
            df = self.read_sheet_advanced(self.file_2B, sheet_name, is_2B=True)
            if not df.empty:
                df = self.map_columns_properly(df, is_2B=True)
            self.data_2B[sheet_name] = df
            print(f"  {sheet_name}: {df.shape[0]} records")
    
    def create_unique_key(self, gstin, invoice_num):
        """Create unique key from GSTIN and Invoice Number"""
        gstin_str = str(gstin).strip() if pd.notna(gstin) else ""
        invoice_str = str(invoice_num).strip() if pd.notna(invoice_num) else ""
        return f"{gstin_str}_{invoice_str}"
    
    def process_and_filter_data(self):
        """Process all sheets and filter 2A data based on 2B data"""
        print("\nFiltering data based on matching criteria...")
        
        sheet_pairs = [
            ("B2B – Invoice & Rate wise", "B2B"),
            ("B2BA", "B2BA"),
            ("CDNR", "B2B-CDNR"),
            ("CDNRA", "B2B-CDNRA")
        ]
        
        total_filtered = 0
        
        for sheet_2A, sheet_2B in sheet_pairs:
            print(f"\nProcessing {sheet_2A} vs {sheet_2B}...")
            
            df_2A = self.data_2A.get(sheet_2A, pd.DataFrame())
            df_2B = self.data_2B.get(sheet_2B, pd.DataFrame())
            
            if df_2A.empty or df_2B.empty:
                print(f"  Skipping - one or both sheets are empty")
                self.filtered_results[sheet_2A] = pd.DataFrame()
                continue
            
            # Find the correct column names for GSTIN and Invoice number
            gstin_col_2A = self.find_column(df_2A, ['gstin', 'supplier'])
            invoice_col_2A = self.find_column(df_2A, ['invoice', 'number']) or self.find_column(df_2A, ['note number'])
            gstin_col_2B = self.find_column(df_2B, ['gstin', 'supplier'])
            invoice_col_2B = self.find_column(df_2B, ['invoice', 'number']) or self.find_column(df_2B, ['note number'])
            
            print(f"  2A columns - GSTIN: {gstin_col_2A}, Invoice: {invoice_col_2A}")
            print(f"  2B columns - GSTIN: {gstin_col_2B}, Invoice: {invoice_col_2B}")
            
            if not all([gstin_col_2A, invoice_col_2A, gstin_col_2B, invoice_col_2B]):
                print(f"  Warning: Could not find required columns")
                self.filtered_results[sheet_2A] = pd.DataFrame()
                continue
            
            # Create unique keys
            df_2B['unique_key'] = df_2B.apply(
                lambda row: self.create_unique_key(row[gstin_col_2B], row[invoice_col_2B]), axis=1
            )
            df_2A['unique_key'] = df_2A.apply(
                lambda row: self.create_unique_key(row[gstin_col_2A], row[invoice_col_2A]), axis=1
            )
            
            # Filter 2A data based on keys present in 2B
            keys_2B = set(df_2B['unique_key'].tolist())
            filtered_2A = df_2A[df_2A['unique_key'].isin(keys_2B)].copy()
            
            print(f"  Unique keys in 2B: {len(keys_2B)}")
            print(f"  Filtered records from 2A: {len(filtered_2A)}")
            
            if len(filtered_2A) > 0:
                # Add the required columns from 2B
                period_col = self.find_column(df_2B, ['gstr-1', 'period'])
                filing_date_col = self.find_column(df_2B, ['gstr-1', 'filing date'])
                
                if period_col and filing_date_col:
                    mapping_dict = df_2B.set_index('unique_key')[[period_col, filing_date_col]].to_dict('index')
                    
                    filtered_2A['GSTR-1/IFF/GSTR-5 Period'] = filtered_2A['unique_key'].map(
                        lambda x: mapping_dict.get(x, {}).get(period_col, '')
                    )
                    filtered_2A['GSTR-1/IFF/GSTR-5 Filing Date'] = filtered_2A['unique_key'].map(
                        lambda x: mapping_dict.get(x, {}).get(filing_date_col, '')
                    )
                    print(f"  Added period and filing date columns")
                
                # Sort by the order in 2B file
                key_order = df_2B['unique_key'].tolist()
                filtered_2A['sort_order'] = filtered_2A['unique_key'].map(
                    lambda x: key_order.index(x) if x in key_order else 999999
                )
                filtered_2A = filtered_2A.sort_values('sort_order')
                
                # Remove helper columns
                filtered_2A = filtered_2A.drop(['unique_key', 'sort_order'], axis=1)
                total_filtered += len(filtered_2A)
            
            self.filtered_results[sheet_2A] = filtered_2A
        
        print(f"\nTotal filtered records: {total_filtered}")
    
    def find_column(self, df, keywords):
        """Find column name containing all specified keywords"""
        for col in df.columns:
            if all(keyword.lower() in col.lower() for keyword in keywords):
                return col
        return None
    
    def create_output_excel(self):
        """Create the output Excel file with filtered data"""
        output_filename = f"{self.legal_name}_{self.tax_period}_{self.current_year}_{self.fixed_suffix}.xlsx"
        
        print(f"\nCreating output file: {output_filename}")
        
        # Create a new workbook
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        # Create sheets for each filtered result
        sheets_to_create = [
            ("B2B – Invoice & Rate wise", self.filtered_results["B2B – Invoice & Rate wise"]),
            ("B2BA", self.filtered_results["B2BA"]),
            ("CDNR", self.filtered_results["CDNR"]),
            ("CDNRA", self.filtered_results["CDNRA"])
        ]
        
        for sheet_name, df in sheets_to_create:
            ws = wb.create_sheet(title=sheet_name)
            
            if df.empty:
                ws['A1'] = "No matching data found"
                continue
            
            # Add data to worksheet
            for r in dataframe_to_rows(df, index=False, header=True):
                ws.append(r)
            
            # Format header row
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Auto-adjust column widths
            for column in ws.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column[0].column_letter].width = adjusted_width
        
        wb.save(output_filename)
        print(f"✓ Output file created: {output_filename}")
        return output_filename
    
    def create_summary_report(self, output_filename):
        """Create a detailed summary report"""
        summary_lines = []
        
        summary_lines.append("EXCEL FILE PROCESSING SUMMARY")
        summary_lines.append("=" * 50)
        summary_lines.append(f"Input File 2A: {self.file_2A}")
        summary_lines.append(f"Input File 2B: {self.file_2B}")
        summary_lines.append(f"Output File: {output_filename}")
        summary_lines.append(f"Processing Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        summary_lines.append("")
        
        summary_lines.append("EXTRACTED METADATA FROM 2B:")
        summary_lines.append(f"- Legal Name: {self.legal_name}")
        summary_lines.append(f"- Tax Period: {self.tax_period}")
        summary_lines.append(f"- Current Year: {self.current_year}")
        summary_lines.append("")
        
        summary_lines.append("SHEET PROCESSING RESULTS:")
        summary_lines.append("-" * 30)
        
        sheet_pairs = [
            ("B2B – Invoice & Rate wise", "B2B"),
            ("B2BA", "B2BA"),
            ("CDNR", "B2B-CDNR"),
            ("CDNRA", "B2B-CDNRA")
        ]
        
        total_filtered = 0
        for sheet_2A, sheet_2B in sheet_pairs:
            df_2A = self.data_2A.get(sheet_2A, pd.DataFrame())
            df_2B = self.data_2B.get(sheet_2B, pd.DataFrame())
            filtered_df = self.filtered_results.get(sheet_2A, pd.DataFrame())
            
            summary_lines.append(f"{sheet_2A}:")
            summary_lines.append(f"  2A Records: {len(df_2A)}")
            summary_lines.append(f"  2B Records: {len(df_2B)}")
            summary_lines.append(f"  Filtered Records: {len(filtered_df)}")
            
            if len(filtered_df) > 0:
                total_filtered += len(filtered_df)
            
            summary_lines.append("")
        
        summary_lines.append("COLUMNS ADDED FROM 2B:")
        summary_lines.append("- GSTR-1/IFF/GSTR-5 Period")
        summary_lines.append("- GSTR-1/IFF/GSTR-5 Filing Date")
        summary_lines.append("")
        
        summary_lines.append("OVERALL SUMMARY:")
        summary_lines.append(f"- Total records filtered and included: {total_filtered}")
        summary_lines.append(f"- Matching criteria: GSTIN of supplier + Invoice/Note number")
        summary_lines.append(f"- Records sorted by 2B file order as requested")
        
        return "\n".join(summary_lines)
    
    def process(self):
        """Main processing function"""
        print("Starting Excel Processing...")
        print("=" * 50)
        
        # Step 1: Validate files
        self.validate_files()
        
        # Step 2: Extract metadata
        self.extract_metadata()
        
        # Step 3: Read all sheets
        self.read_all_sheets()
        
        # Step 4: Process and filter data
        self.process_and_filter_data()
        
        # Step 5: Create output file
        output_filename = self.create_output_excel()
        
        # Step 6: Create summary report
        summary = self.create_summary_report(output_filename)
        summary_filename = "Processing_Summary.txt"
        
        with open(summary_filename, "w", encoding="utf-8") as f:
            f.write(summary)
        
        print(f"✓ Summary saved to: {summary_filename}")
        print("\n" + "=" * 50)
        print("PROCESSING COMPLETED SUCCESSFULLY!")
        print(f"Output file: {output_filename}")
        print(f"Summary file: {summary_filename}")
        
        return output_filename, summary_filename


def main():
    """Main function to run the processor"""
    try:
        # Initialize processor
        processor = ExcelProcessor()
        
        # Run processing
        output_file, summary_file = processor.process()
        
        print(f"\n🎉 SUCCESS!")
        print(f"📊 Output Excel File: {output_file}")
        print(f"📋 Processing Summary: {summary_file}")
        
    except FileNotFoundError as e:
        print(f"❌ ERROR: {e}")
        print("Please ensure both 2A.xlsx and 2B.xlsx files are in the same directory as this script.")
        sys.exit(1)
    
    except Exception as e:
        print(f"❌ ERROR: An unexpected error occurred: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()