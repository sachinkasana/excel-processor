import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import io
import tempfile
import os


class StreamlitExcelProcessor:
    """Juric Excel processor for GSTR data"""
    
    def __init__(self):
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
    
    def extract_metadata(self, file_2B_content):
        """Extract Legal Name and Tax Period from 2B Read me sheet"""
        try:
            read_me_2B = pd.read_excel(file_2B_content, sheet_name='Read me', header=None)
            
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
            
            return True, f"Legal Name: {self.legal_name}, Tax Period: {self.tax_period}"
            
        except Exception as e:
            self.legal_name = "Unknown"
            self.tax_period = "Unknown"
            return False, f"Could not extract metadata: {str(e)}"
    
    def read_sheet_advanced(self, file_content, sheet_name, is_2B=False):
        """Advanced sheet reading that handles different sheet types"""
        try:
            df_raw = pd.read_excel(file_content, sheet_name=sheet_name, header=None)
            
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
            df = pd.read_excel(file_content, sheet_name=sheet_name, skiprows=data_start_row, header=None)
            
            # Assign column names
            if len(combined_headers) <= len(df.columns):
                df.columns = combined_headers[:len(df.columns)]
            else:
                df.columns = combined_headers[:len(df.columns)]
            
            # Remove completely empty rows
            df = df.dropna(how='all')
            
            return df
            
        except Exception as e:
            return pd.DataFrame()
    
    def map_columns_properly(self, df, is_2B=False, sheet_name=""):
        """Map columns to proper names based on position and file type"""
        # Don't remap if columns already look correct
        if any(col for col in df.columns if 'gstin' in str(col).lower() and 'supplier' in str(col).lower()):
            return df
        
        if is_2B:
            if 'CDNR' in sheet_name:
                column_mapping = {
                    0: 'GSTIN of supplier', 1: 'Trade/Legal name', 2: 'Note number', 3: 'Note type',
                    4: 'Note Supply type', 5: 'Note date', 6: 'Note Value (₹)', 7: 'Place of supply',
                    8: 'Supply Attract Reverse Charge', 9: 'Taxable Value (₹)', 10: 'Integrated Tax(₹)',
                    11: 'Central Tax(₹)', 12: 'State/UT Tax(₹)', 13: 'Cess(₹)',
                    14: 'Whether ITC to be reduced (Taxpayer\'s Input)', 15: 'Integrated Tax(₹)',
                    16: 'Central Tax(₹)', 17: 'State/UT Tax(₹)', 18: 'Cess(₹)',
                    19: 'GSTR-1/IFF/GSTR-5 Period', 20: 'GSTR-1/IFF/GSTR-5 Filing Date',
                    21: 'ITC Availability', 22: 'Reason', 23: 'Applicable % of Tax Rate',
                    24: 'Source', 25: 'IRN', 26: 'IRN Date'
                }
            else:
                column_mapping = {
                    0: 'GSTIN of supplier', 1: 'Trade/Legal name', 2: 'Invoice number',
                    3: 'Invoice type', 4: 'Invoice Date', 5: 'Invoice Value(₹)',
                    6: 'Place of supply', 7: 'Supply Attract Reverse Charge',
                    8: 'Taxable Value (₹)', 9: 'Integrated Tax(₹)', 10: 'Central Tax(₹)',
                    11: 'State/UT Tax(₹)', 12: 'Cess(₹)', 13: 'GSTR-1/IFF/GSTR-5 Period',
                    14: 'GSTR-1/IFF/GSTR-5 Filing Date', 15: 'ITC Availability',
                    16: 'Reason', 17: 'Applicable % of Tax Rate', 18: 'Source',
                    19: 'IRN', 20: 'IRN Date'
                }
        else:
            if 'CDNR' in sheet_name:
                column_mapping = {
                    0: 'GSTIN/UIN of Supplier', 1: 'Trade/Legal name of the supplier',
                    2: 'Invoice Number', 3: 'Invoice Date', 4: 'Note type', 5: 'Note number',
                    6: 'Note  date', 7: 'Note Value (₹)', 8: 'Reason', 9: 'Rate (%)',
                    10: 'Taxable Value (₹)', 11: 'Integrated Tax (₹)', 12: 'Central Tax (₹)',
                    13: 'State Tax (₹)', 14: 'Cess Amount (₹)', 15: 'Counter Party Return status',
                    16: 'Filing Period'
                }
            else:
                column_mapping = {
                    0: 'GSTIN of supplier', 1: 'Trade/Legal name of the Supplier',
                    2: 'Invoice number', 3: 'Invoice type', 4: 'Invoice Date',
                    5: 'Invoice Value (₹)', 6: 'Place of supply', 7: 'Supply Attract Reverse Charge',
                    8: 'Rate (%)', 9: 'Taxable Value (₹)', 10: 'Integrated Tax  (₹)',
                    11: 'Central Tax (₹)', 12: 'State/UT tax (₹)', 13: 'Cess  (₹)',
                    14: 'Counter Party Return status', 15: 'Filing Period', 16: 'Notes'
                }
        
        new_columns = []
        for i, col in enumerate(df.columns):
            if i in column_mapping:
                new_columns.append(column_mapping[i])
            else:
                new_columns.append(str(col))
        
        df.columns = new_columns
        return df
    
    def read_all_sheets(self, file_2A_content, file_2B_content):
        """Read all required sheets from both files"""
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_sheets = len(self.sheets_2A) + len(self.sheets_2B)
        current_sheet = 0
        
        # Read 2A sheets
        status_text.text("Reading sheets from 2A file...")
        for sheet_name in self.sheets_2A:
            df = self.read_sheet_advanced(file_2A_content, sheet_name, is_2B=False)
            if not df.empty:
                df = self.map_columns_properly(df, is_2B=False, sheet_name=sheet_name)
            self.data_2A[sheet_name] = df
            current_sheet += 1
            progress_bar.progress(current_sheet / total_sheets)
        
        # Read 2B sheets
        status_text.text("Reading sheets from 2B file...")
        for sheet_name in self.sheets_2B:
            df = self.read_sheet_advanced(file_2B_content, sheet_name, is_2B=True)
            if not df.empty:
                df = self.map_columns_properly(df, is_2B=True, sheet_name=sheet_name)
            self.data_2B[sheet_name] = df
            current_sheet += 1
            progress_bar.progress(current_sheet / total_sheets)
        
        progress_bar.progress(1.0)
        status_text.text("✓ All sheets loaded successfully!")
    
    def create_unique_key(self, gstin, key_value):
        """Create unique key from GSTIN and key value"""
        gstin_str = str(gstin).strip() if pd.notna(gstin) else ""
        key_str = str(key_value).strip() if pd.notna(key_value) else ""
        return f"{gstin_str}_{key_str}"
    
    def find_column(self, df, keywords):
        """Find column name containing all specified keywords"""
        for col in df.columns:
            col_lower = str(col).lower()
            if all(keyword.lower() in col_lower for keyword in keywords):
                return col
        return None
    
    def process_and_filter_data(self):
        """Process all sheets and filter 2A data based on 2B data"""
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        sheet_pairs = [
            ("B2B – Invoice & Rate wise", "B2B"),
            ("B2BA", "B2BA"),
            ("CDNR", "B2B-CDNR"),
            ("CDNRA", "B2B-CDNRA")
        ]
        
        total_filtered = 0
        
        for idx, (sheet_2A, sheet_2B) in enumerate(sheet_pairs):
            status_text.text(f"Processing {sheet_2A} vs {sheet_2B}...")
            
            df_2A = self.data_2A.get(sheet_2A, pd.DataFrame())
            df_2B = self.data_2B.get(sheet_2B, pd.DataFrame())
            
            if df_2A.empty or df_2B.empty:
                self.filtered_results[sheet_2A] = pd.DataFrame()
                progress_bar.progress((idx + 1) / len(sheet_pairs))
                continue
            
            # Determine matching columns based on sheet type
            if (sheet_2A == "CDNR" and sheet_2B == "B2B-CDNR"):
                gstin_col_2A = self.find_column(df_2A, ['gstin/uin', 'supplier']) or self.find_column(df_2A, ['gstin', 'supplier'])
                key_col_2A = self.find_column(df_2A, ['note number'])
                gstin_col_2B = self.find_column(df_2B, ['gstin', 'supplier'])
                key_col_2B = self.find_column(df_2B, ['note number'])
            else:
                gstin_col_2A = self.find_column(df_2A, ['gstin', 'supplier'])
                key_col_2A = self.find_column(df_2A, ['invoice', 'number'])
                gstin_col_2B = self.find_column(df_2B, ['gstin', 'supplier'])
                key_col_2B = self.find_column(df_2B, ['invoice', 'number'])
            
            if not all([gstin_col_2A, key_col_2A, gstin_col_2B, key_col_2B]):
                self.filtered_results[sheet_2A] = pd.DataFrame()
                progress_bar.progress((idx + 1) / len(sheet_pairs))
                continue
            
            # Create unique keys
            df_2B['unique_key'] = df_2B.apply(
                lambda row: self.create_unique_key(row[gstin_col_2B], row[key_col_2B]), axis=1
            )
            df_2A['unique_key'] = df_2A.apply(
                lambda row: self.create_unique_key(row[gstin_col_2A], row[key_col_2A]), axis=1
            )
            
            # Filter 2A data based on keys present in 2B
            keys_2B = set(df_2B['unique_key'].tolist())
            filtered_2A = df_2A[df_2A['unique_key'].isin(keys_2B)].copy()
            
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
                
                # Sort by 2B order
                key_order = df_2B['unique_key'].tolist()
                filtered_2A['sort_order'] = filtered_2A['unique_key'].map(
                    lambda x: key_order.index(x) if x in key_order else 999999
                )
                filtered_2A = filtered_2A.sort_values('sort_order')
                filtered_2A = filtered_2A.drop(['unique_key', 'sort_order'], axis=1)
                total_filtered += len(filtered_2A)
            
            self.filtered_results[sheet_2A] = filtered_2A
            progress_bar.progress((idx + 1) / len(sheet_pairs))
        
        status_text.text(f"✓ Filtering complete! {total_filtered} total records matched")
        return total_filtered
    
    def create_output_excel(self):
        """Create the output Excel file and return as bytes"""
        output_filename = f"{self.legal_name}_{self.tax_period}_{self.current_year}_{self.fixed_suffix}.xlsx"
        
        # Create workbook in memory
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            sheets_created = 0
            
            for sheet_name in ["B2B – Invoice & Rate wise", "B2BA", "CDNR", "CDNRA"]:
                df = self.filtered_results.get(sheet_name, pd.DataFrame())
                
                if df.empty:
                    # Create empty sheet with message
                    empty_df = pd.DataFrame({"Message": ["No matching data found"]})
                    empty_df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    sheets_created += 1
                    
                    # Format the sheet
                    worksheet = writer.sheets[sheet_name]
                    
                    # Style header row
                    if not df.empty:
                        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
                        header_font = Font(color="FFFFFF", bold=True)
                        
                        for cell in worksheet[1]:
                            cell.fill = header_fill
                            cell.font = header_font
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        
                        # Auto-adjust column widths
                        for column in worksheet.columns:
                            max_length = 0
                            column = [cell for cell in column]
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 50)
                            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        
        output.seek(0)
        return output.getvalue(), output_filename


# Hide Streamlit header and menu
hide_streamlit_style = """
<style>
#MainMenu {visibility: hidden;}
header {visibility: hidden;}
footer {visibility: hidden;}
.stActionButton {visibility: hidden;}
/* Hide manage app button - MOST IMPORTANT */
.stAppDeployButton {display: none !important;}
.stActionButton {display: none !important;}
button[data-testid="manage-app-button"] {display: none !important;}

/* Hide everything that might contain "Manage app" */
div[data-testid="stVerticalBlock"] button:contains("Manage") {display: none !important;}

/* Super aggressive hiding of manage button */
*[class*="manage"], *[class*="Manage"], *[class*="MANAGE"] {
    display: none !important;
}

/* Remove any overlay buttons */
div[style*="z-index"] button {
    display: none !important;
}

/* Hide everything that could be a floating button */
div[style*="position: fixed"] {
    display: none !important;
}

</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="Juric",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("Juric")
    st.markdown("""
    **Process your GSTR 2A and 2B Excel files instantly!**
    
    This tool filters records from 2A based on matching data in 2B using GSTIN + Invoice/Note number combinations,
    and adds the required GSTR-1 period and filing date columns.
    """)
    
    # Sidebar with instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        ### How to use:
        1. Upload your **2A.xlsx** file
        2. Upload your **2B.xlsx** file  
        3. Click **Process Files**
        4. Download the filtered result
        
        ### What it does:
        - ✅ Matches records using **GSTIN + Invoice number**
        - ✅ For CDNR sheets: **GSTIN + Note number**
        - ✅ Adds period and filing date columns
        - ✅ Maintains proper sorting order
        - ✅ Processes 4 sheets: B2B, B2BA, CDNR, CDNRA
        
        ### File Requirements:
        - Files must be in **.xlsx** format
        - Must contain the standard GSTR sheets
        - 2B file should have a 'Read me' sheet
        """)
        
        st.markdown("---")
        st.markdown("**Made with ❤️ in India**")
    
    # Main interface
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📄 Upload 2A File")
        file_2A = st.file_uploader(
            "Choose 2A.xlsx file",
            type=['xlsx'],
            help="Upload your GSTR 2A Excel file"
        )
        
        if file_2A:
            st.success(f"✅ 2A File uploaded: {file_2A.name}")
            st.info(f"File size: {file_2A.size / 1024:.1f} KB")
    
    with col2:
        st.subheader("📄 Upload 2B File")
        file_2B = st.file_uploader(
            "Choose 2B.xlsx file",
            type=['xlsx'],
            help="Upload your GSTR 2B Excel file"
        )
        
        if file_2B:
            st.success(f"✅ 2B File uploaded: {file_2B.name}")
            st.info(f"File size: {file_2B.size / 1024:.1f} KB")
    
    # Process button
    if file_2A and file_2B:
        st.markdown("---")
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Process Files", type="primary", use_container_width=True):
                try:
                    with st.spinner("Processing your files..."):
                        processor = StreamlitExcelProcessor()
                        
                        # Extract metadata
                        metadata_success, metadata_msg = processor.extract_metadata(file_2B)
                        st.info(f"📋 Extracted: {metadata_msg}")
                        
                        # Read all sheets
                        processor.read_all_sheets(file_2A, file_2B)
                        
                        # Process and filter
                        total_records = processor.process_and_filter_data()
                        
                        # Create output
                        excel_data, filename = processor.create_output_excel()
                        
                        # Success message
                        st.success(f"🎉 Processing complete! {total_records} records filtered and processed.")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download Processed File",
                            data=excel_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                        
                        # Show summary
                        with st.expander("📊 Processing Summary"):
                            summary_data = []
                            for sheet_name in processor.filtered_results:
                                df = processor.filtered_results[sheet_name]
                                summary_data.append({
                                    "Sheet": sheet_name,
                                    "Records Found": len(df)
                                })
                            
                            summary_df = pd.DataFrame(summary_data)
                            st.dataframe(summary_df, use_container_width=True)
                
                except Exception as e:
                    st.error(f"❌ Error processing files: {str(e)}")
                    st.error("Please check your files and try again.")
    
    else:
        st.info("👆 Please upload both 2A and 2B Excel files to get started.")


if __name__ == "__main__":
    main()