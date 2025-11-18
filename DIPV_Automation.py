import fsDisppy as fs 
from datetime import datetime, date, timedelta
import os
import pandas as pd
import xlwings as xw
import numpy as np
import time
import ipywidgets as widgets
from IPython.display import display, HTML
from typing import Union, Any
import win32com.client 
import re

SAVE_FOLDER = r"C:\Users\goyaman\OneDrive - Deutsche Bank AG\Desktop\DIPV Automation"
os.makedirs(SAVE_FOLDER, exist_ok=True)


# --- Section 2 ---
def _apply_header_fix(raw_data_list) -> pd.DataFrame:  
    # creates df and handles header and column cleanup
    if isinstance(raw_data_list, str):
        if raw_data_list.strip() == "##no rows!":
            return pd.DataFrame() # Return empty DataFrame if this specific string is found
        else:
            # If actual error message or unexpected format, reurn empty dataframe
            print(f"WARNING: fs.RiskDBLoad returned an unexpected string: '{raw_data_list}'. Returning an empty DataFrame.")
            return pd.DataFrame()

    if not raw_data_list:
        return pd.DataFrame()
    
    try:
        df = pd.DataFrame(raw_data_list)
    except Exception as e:
        print(f"CRITICAL ERROR: Failed to create DataFrame from raw data (type: {type(raw_data_list)}, value: {str(raw_data_list)[:100]}...). Original error: {e}")
        return pd.DataFrame() # Return empty DataFrame on any unexpected DataFrame creation error

    # Attempt to promote first row to header if it contains meaningful data
    if not df.empty and df.iloc[0].notna().any():
        df.columns = df.iloc[0]
        df = df.iloc[1:].reset_index(drop=True)
    else:
        # If first row is empty/meaningless, just drop it
        df = df.iloc[1:].reset_index(drop=True)

    # Clean up column names: ensure strings, handle duplicates, remove leading/trailing spaces
    cleaned_cols = []
    seen_cols = {}
    for col_name in df.columns:
        str_col = str(col_name).strip()
        
        # Handle empty/NaN column names
        if not str_col or str_col == 'None' or str_col == 'nan' or str_col == '':
            str_col = 'Unnamed_Col'

        # Handle duplicate column names (e.g., if multiple unnamed columns or identical strings)
        if str_col in seen_cols:
            seen_cols[str_col] += 1
            str_col = f"{str_col}_{seen_cols[str_col]}"
        else:
            seen_cols[str_col] = 0 # Initialize count for first occurrence
        cleaned_cols.append(str_col)
    
    df.columns = cleaned_cols
    return df

# --- Function to load data using fs.RiskDBLoad ---
def load_data_for_dt(dt_obj: datetime.date):
    
    print(f"Attempting to load data for COB date: {dt_obj.strftime('%Y-%m-%d')} using fs.RiskDBLoad...")

    pfo = ['U1953','U565','U11192', 'U11119', 'U1954', 'U828', 'U3200']
    prop1 = ['Book', 'Cpty', 'Book.UBR5_Name', 'Book.UBR7_Name', 'Book.UBR8_Name', 'Book.UBR10_Name', 'Book.UBR11_Name']
    prop2 = ['Book', 'Cpty', 'Book.UBR5_Name', 'Book.UBR7_Name', 'Book.UBR8_Name']
    control = "ReportCcy=EUR, Headers"
    
    l_rple = fs.RiskDBLoad( '@lon', 'IPVVarianceRiskPLE', dt_obj, pfo, 'EOD_IP', prop1, None, None, control).asPy()
    n_rple = fs.RiskDBLoad( '@nyc', 'IPVVarianceRiskPLE', dt_obj, pfo, 'EOD_IP', prop1, None, None, control).asPy()
    l_bond = fs.RiskDBLoad( '@lon', 'IPVVarianceBond', dt_obj, pfo, 'EOD_IP', prop2, None, None, control).asPy()
    n_bond = fs.RiskDBLoad( '@nyc', 'IPVVarianceBond', dt_obj, pfo, 'EOD_IP', prop2, None, None, control).asPy()
    l_listed = fs.RiskDBLoad( '@lon', 'IPVVarianceListed', dt_obj, pfo, 'EOD_IP', prop2, None, None, control).asPy()

    print("Data loading complete.")
    return l_rple, n_rple, l_bond, n_bond, l_listed


# --- Section 3 ---
# This list will be used within the main execution function, after source_dfs are prepared
PIVOT_TABLE_CONFIGS_TEMPLATE = [
    {
        "name": "LON_RISKPLE",
        "source_df_key": "lon_riskple", 
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name"],
        "pivot_columns": ["MktType"],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["IP", "QAP", "IPV"]},
        "output_sheet_name": "L_RPLE",
        "target_excel_cell": "A1"
    },
    {
        "name": "NYC_RISKPLE",
        "source_df_key": "nyc_riskple",
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name"],
        "pivot_columns": ["MktType"],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["IP", "QAP", "IPV"]},
        "output_sheet_name": "N_RPLE",
        "target_excel_cell": "A1"
    },
    {
        "name": "LON_LISTED",
        "source_df_key": "lon_listed",
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name"],
        "pivot_columns": [],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["IPV"]},
        "output_sheet_name": "L_LISTED",
        "target_excel_cell": "A1"
    },
    {
        "name": "NYC_BONDS",
        "source_df_key": "nyc_bonds",
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name"],
        "pivot_columns": ["MktType"],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["IP", "QAP", "IPV"]},
        "output_sheet_name": "N_BOND",
        "target_excel_cell": "A1"
    },
    {
        "name": "LON_BONDS",
        "source_df_key": "lon_bonds",
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name"],
        "pivot_columns": ["MktType"],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["IP", "QAP", "IPV"]},
        "output_sheet_name": "L_BOND",
        "target_excel_cell": "A1"
    },
    {
        "name": "UNEXPLAINED",
        "source_df_key": "lon_riskple", 
        "index_columns": ["Book.UBR5_Name", "Book.UBR7_Name", "Book.UBR8_Name", "Book.UBR10_Name", "MktIPVClassification"],
        "pivot_columns": ["MktType"],
        "value_columns": ["Variance"],
        "aggregation_function": "sum",
        "fill_na_value": 0,
        "filters": {'MktIPVClassification': ["PT", "PT1"], 'Book.UBR8_Name': ["Europe Non Linear"]},
        "output_sheet_name": "UNEXPL",
        "target_excel_cell": "A1"
    }
]

# --- Core Function ---
def create_and_flatten_pivot(
    source_dataframe: pd.DataFrame,
    pivot_configuration: dict
) -> pd.DataFrame:
    """
    Creates a pivot table from the source DataFrame based on the provided configuration,
    filters it, and flattens its MultiIndex columns into single-level headers.
    """
    
    # check for missing columns
    all_specified_cols = list(pivot_configuration['index_columns']) +                          list(pivot_configuration['pivot_columns']) +                          list(pivot_configuration['value_columns'])
    if 'filters' in pivot_configuration and pivot_configuration['filters']:
        all_specified_cols.extend(list(pivot_configuration['filters'].keys()))
    
    all_specified_cols = list(set(all_specified_cols))

    missing_cols = [col for col in all_specified_cols if col not in source_dataframe.columns]
    if missing_cols:
        error_df_cols = list(pivot_configuration['index_columns']) + ['ERROR_MISSING_COLUMNS'] + missing_cols
        return pd.DataFrame(columns=error_df_cols)

    # Filter DataFrame
    df_for_pivot = source_dataframe.copy()
    if 'filters' in pivot_configuration and pivot_configuration['filters']:
        for col, values in pivot_configuration['filters'].items():
            if col in df_for_pivot.columns: # Redundant check but good for robustness
                df_for_pivot = df_for_pivot[df_for_pivot[col].isin(values)]
            else:
                print(f"  Warning: Filter column '{col}' not found in DataFrame for pivot '{pivot_configuration['name']}'. Filter skipped.")

    # Handle Empty DataFrame after Filtering
    if df_for_pivot.empty:
        output_cols_base = list(pivot_configuration['index_columns'])
        predicted_data_cols = []

        if not pivot_configuration['pivot_columns']:
            predicted_data_cols = list(pivot_configuration['value_columns'])
        else:
            dummy_multi_index_tuples = []
            if len(pivot_configuration['value_columns']) > 1: # Values on top level, then pivot_columns
                for val_col in pivot_configuration['value_columns']:
                    for p_col in pivot_configuration['pivot_columns']:
                        dummy_multi_index_tuples.append((val_col, f"Predicted_{p_col}_Value"))
            else: # If only one value column, pivot_columns are the main headers
                for p_col in pivot_configuration['pivot_columns']:
                    dummy_multi_index_tuples.append((pivot_configuration['value_columns'][0], f"Predicted_{p_col}_Value"))

            if dummy_multi_index_tuples:
                dummy_multi_index_obj = pd.MultiIndex.from_tuples(dummy_multi_index_tuples)
                for col_tuple in dummy_multi_index_obj:
                    final_name = None
                    for part in reversed(col_tuple):
                        if part is not None and str(part).strip() != '':
                            final_name = str(part)
                            break
                    if final_name is None:
                        final_name = '_'.join(str(p) for p in col_tuple if p is not None and str(p).strip() != '')
                        if not final_name:
                            final_name = 'Unnamed_Col'
                    predicted_data_cols.append(final_name)
            else:
                predicted_data_cols.append("NO_DATA_COLUMNS_PREDICTED")

        final_empty_df_cols = output_cols_base + predicted_data_cols
        return pd.DataFrame(columns=final_empty_df_cols)

    # Create Pivot Table
    pivot_df_result = pd.pivot_table(
        df_for_pivot,
        index=pivot_configuration['index_columns'],
        columns=pivot_configuration['pivot_columns'],
        values=pivot_configuration['value_columns'],
        aggfunc=pivot_configuration['aggregation_function'],
        fill_value=pivot_configuration['fill_na_value']
    )

    # Reset Index
    pivot_df_result = pivot_df_result.reset_index()

    # Flatten MultiIndex Columns
    if isinstance(pivot_df_result.columns, pd.MultiIndex):
        new_column_names = []
        for col_tuple in pivot_df_result.columns:
            final_name = None
            for part in reversed(col_tuple):
                if part is not None and str(part).strip() != '':
                    final_name = str(part)
                    break
            
            if final_name is None:
                final_name = '_'.join(str(p) for p in col_tuple if p is not None and str(p).strip() != '')
                if not final_name:
                     final_name = 'Unnamed_Col'

            new_column_names.append(final_name)
        
        pivot_df_result.columns = new_column_names
    
    return pivot_df_result


# In[4]:


# --- Main Function to Export Pivots to Excel ---
def automate_pivot_export_to_excel(
    target_cob_date: datetime.date, 
    base_save_folder: str, # This parameter is still passed but the template_filepath is hardcoded
    pivot_table_configs: list
):
    """
    Automates the creation of multiple pivot tables based on configurations.
    It opens a specific 'DIPV Template.xlsx' template, populates it with pivots, and then
    saves it as 'DIPV_YYYY_MM_DD.xlsx' in the specified save folder.
    """
    
    template_filepath = r"C:\Users\goyaman\OneDrive - Deutsche Bank AG\Desktop\DIPV Automation\DIPV Template.xlsx"
    
    output_filename = f"DIPV_{target_cob_date.strftime('%Y_%m_%d')}.xlsx"
    output_filepath = os.path.join(base_save_folder, output_filename) # Still uses base_save_folder for output

    print(f"\nStarting automation for COB: {target_cob_date.strftime('%Y-%m-%d')}")
    print(f"Template file to open: {template_filepath}")
    print(f"Output file to save as: {output_filepath}")

    excel_app = None
    excel_workbook = None

    try:
        # Check if the DIPV Template.xlsx template exists
        if not os.path.exists(template_filepath):
            raise FileNotFoundError(f"Template file not found: {template_filepath}. Please ensure 'DIPV Template.xlsx' exists at the specified path.")

        # Initialize xlwings application
        excel_app = xw.App(visible=False, add_book=False)

        # Open the specific 'DIPV Template.xlsx' template
        print(f"Please wait, this process takes time!")
        excel_workbook = excel_app.books.open(template_filepath)
                
        # Load raw data for the determined COB date using fs.RiskDBLoad
        l_rple, n_rple, l_bond, n_bond, l_listed = load_data_for_dt(target_cob_date)

        # Apply header fix to all loaded raw data to get usable DataFrames
        source_dfs_dict = {
            "lon_riskple": _apply_header_fix(l_rple),
            "nyc_riskple": _apply_header_fix(n_rple),
            "lon_listed": _apply_header_fix(l_listed),
            "nyc_bonds": _apply_header_fix(n_bond),
            "lon_bonds": _apply_header_fix(l_bond)
        }

        # Iterate through pivot configurations, create pivots, and export to Excel
        for config_template in pivot_table_configs:
            config = config_template.copy()
            config["source_df"] = source_dfs_dict.get(config["source_df_key"], pd.DataFrame())
            sheet_name = config["output_sheet_name"]
            target_cell = config["target_excel_cell"]

            try:
                final_pivot_dataframe = create_and_flatten_pivot(
                    config["source_df"],
                    config
                )
                
                if sheet_name not in [s.name for s in excel_workbook.sheets]:
                    print(f"Warning: Sheet '{sheet_name}' not found in 'DIPV Template.xlsx'. Adding it.")
                    excel_sheet = excel_workbook.sheets.add(name=sheet_name)
                else:
                    excel_sheet = excel_workbook.sheets[sheet_name]

                excel_sheet.clear_contents()
                excel_sheet.range(target_cell).options(index=False, header=True).value = final_pivot_dataframe

            except Exception as e_pivot:
                print(f"❌ Error processing pivot '{config['name']}': {e_pivot}")     
        
        
        # Refresh all calculations and external data connections in the workbook before saving
        print("Refreshing all Excel calculations and external data connections...")
        excel_workbook.api.RefreshAll()
        excel_app.calculate()
        
        dipv_sheet_name = "DIPV"
        if dipv_sheet_name in [s.name for s in excel_workbook.sheets]:
            dipv_sheet = excel_workbook.sheets[dipv_sheet_name]
        
            formatted_cob_date = target_cob_date.strftime('%Y-%m-%d')
            dipv_sheet.range('AE3').value = formatted_cob_date
                      
            check_value = dipv_sheet.range('AE6').value
            print(f"Process {check_value}")       
            
        else:
            print(f"Warning: Sheet '{dipv_sheet_name}' not found in '{os.path.basename(template_filepath)}'.")
        
        excel_workbook.save(output_filepath)
        print(f"✅ Successfully saved Excel file as: {output_filepath}")
          
                      
    except FileNotFoundError as e:
        print(f"❌ ERROR: {e}")
    except Exception as e_main:
        print(f"❌ An unexpected error occurred during Excel automation: {e_main}")

    finally:
        if excel_workbook:
            try:
                excel_workbook.close()
            except Exception as e:
                print(f"Warning: Error closing workbook: {e}")
        if excel_app:
            try:
                excel_app.quit()
            except Exception as e:
                print(f"Warning: Error quitting Excel app: {e}")
    return output_filepath


# --- Section 5 ---

HTML_EMAIL_STYLES = """
    <style>
        body { font-family: Aptos, Arial, Segoe UI, sans-serif; font-size: 14.5px; }
        p { font-family: Aptos, Arial, Segoe UI, sans-serif; font-size: 14.5px; }
        table {
            border-collapse: collapse;
            width: 100%; /* Make tables responsive to email client width */
            font-family: Aptos, Arial, Segoe UI, sans-serif;
            font-size: 12px;
            table-layout: fixed; /* CRUCIAL for explicit column widths */
            margin-bottom: 20px; /* Space between tables */
        }
        th, td {
            border: 1px solid #cccccc; /* Light gray border for a cleaner look */
            padding: 8px; /* More padding for readability */
            vertical-align: top;
            word-wrap: break-word; /* Allow long words to break */
            overflow-wrap: break-word; /* Allow long words to break */
        }
        th {
            background-color: #004080; /* Dark blue header */
            color: white;
            font-weight: bold;
            text-align: center;
            white-space: normal; /* Ensure headers wrap if needed */
        }
        /* Zebra striping will be applied inline for better compatibility */
        /* Negative numbers will have inline style="color: red;" */
    </style>
"""

# NEW FUNCTION: _dataframe_to_html_manual to manually generate HTML table
def _dataframe_to_html_manual(df: pd.DataFrame, table_title: str = "Table") -> str:
    """
    Manually converts a pandas DataFrame to an HTML table string, applying numeric formatting
    and conditional coloring for negative numbers, with explicit column width control.
    """
    if not isinstance(df, pd.DataFrame) or df is None or df.empty or df.shape[1] == 0:
        print(f"  WARNING: Input DataFrame for '{table_title}' was empty or invalid. Returning placeholder HTML.")
        return f"<p style='font-family: Aptos, Arial, Segoe UI, sans-serif;'>No data available for {table_title}.</p>"

    # Define desired column widths in pixels. TUNE THESE VALUES!
    # These are example widths. Adjust them based on your actual data content.
    # If a column isn't listed here, it will get a default '100px'.
    # Ensure the sum of widths doesn't exceed a reasonable email client width (e.g., 800-1000px)
    # or make the table wider than necessary, forcing horizontal scroll.
    specific_col_widths = {
        'Book': '80px',
        'Cpty': '100px',
        'Book.UBR5_Name': '120px',
        'Book.UBR7_Name': '120px',
        'Book.UBR8_Name': '150px',
        'Book.UBR10_Name': '150px',
        'Book.UBR11_Name': '150px',
        'MktType': '70px',
        'Variance': '100px',
        'MktIPVClassification': '100px',
        # Add any other column names that require specific widths.
        # Example for a hypothetical 'Comment' column:
        'Comment': '200px' 
    }
    DEFAULT_COL_WIDTH = '100px' # Default width for columns not in specific_col_widths

    html_string = "<table style='table-layout:fixed; width:100%;'>\n" # Inline style for critical properties

    # --- Generate <colgroup> and <col> tags for robust width control ---
    html_string += "  <colgroup>\n"
    total_specified_width_px = 0
    specified_cols_count = 0
    for col in df.columns:
        width_str = specific_col_widths.get(col, DEFAULT_COL_WIDTH)
        html_string += f"    <col style='width: {width_str};'>\n"
        # Sum up for percentage calculation if needed, but pixel is often more reliable
        try:
            total_specified_width_px += int(width_str.replace('px', ''))
            specified_cols_count += 1
        except ValueError:
            pass # Handle non-pixel widths if they appear
    html_string += "  </colgroup>\n"


    # --- Table Headers ---
    html_string += "  <thead>\n    <tr>\n"
    for col in df.columns:
        header_width_style = f"width: {specific_col_widths.get(col, DEFAULT_COL_WIDTH)};"
        # Headers should wrap, so white-space:normal
        html_string += f"      <th style='{header_width_style} white-space: normal; overflow: hidden; text-overflow: ellipsis;'>{col}</th>\n"
    html_string += "    </tr>\n  </thead>\n"

    # --- Table Body ---
    html_string += "  <tbody>\n"
    for idx, row in df.iterrows():
        row_style = "background-color: #f2f2f2;" if idx % 2 == 0 else "" # Inline zebra striping
        html_string += f"    <tr style='{row_style}'>\n"
        for col_name in df.columns:
            cell_value = row[col_name]
            cell_style = ""
            
            # Use specific column width for the cell
            cell_width_style = f"width: {specific_col_widths.get(col_name, DEFAULT_COL_WIDTH)};"

            # Replace None/NaN with empty string
            if pd.isna(cell_value) or str(cell_value).strip().lower() == 'none':
                formatted_value = ""
            else:
                # Try to format as number
                try:
                    num_value = pd.to_numeric(cell_value, errors='coerce')
                    if not pd.isna(num_value): # Successfully converted to number
                        if num_value < 0:
                            cell_style += "color: red;"
                        
                        # Apply numeric formatting: commas, 2 decimal places or integer for large numbers
                        # This logic ensures the value is consistently numeric before formatting
                        if isinstance(num_value, (int, float)) and num_value == int(num_value) and abs(num_value) > 999:
                            formatted_value = f"{int(num_value):,}"
                        elif isinstance(num_value, (int, float)):
                            formatted_value = f"{num_value:,.2f}"
                        else:
                            formatted_value = str(cell_value) # Fallback, should not happen if not NaN
                    else: # Converted to NaN due to coerce, means it wasn't numeric
                        formatted_value = str(cell_value)
                except (ValueError, TypeError):
                    formatted_value = str(cell_value) # Not a number, keep as string

            # MODIFIED: Apply ZWSP for comment columns to aid wrapping
            if col_name in specific_col_widths and 'Comment' in col_name and isinstance(formatted_value, str):
                formatted_value = formatted_value.replace(' ', ' \u200b') # Inject ZWSP

            # Final cell style for td: combine width, text alignment and conditional color
            # Use `text-align: right;` for numeric columns to improve readability
            final_td_style = f"{cell_width_style} {cell_style}"
            if pd.api.types.is_numeric_dtype(df[col_name].apply(pd.to_numeric, errors='coerce')):
                final_td_style += " text-align: right;"
            else:
                final_td_style += " text-align: left;"
            
            # Ensure cell content wraps by default, unless overridden by a specific column style
            final_td_style += " white-space: normal; word-wrap: break-word; overflow-wrap: break-word;"

            html_string += f"      <td style='{final_td_style}'>{formatted_value}</td>\n"
        html_string += "    </tr>\n"
    html_string += "  </tbody>\n"
    html_string += "</table>"
    
    return html_string


# MODIFIED: send_dipv_emails function to use manual HTML generation
def send_dipv_emails(excel_filepath: str, cob_date: datetime.date, dipv_check_value: Union[str, float, None]):

    print("\n--- Starting Email Sending Process via Outlook ---")

    RECIPIENT_EMAIL = 'manish.goyal@db.com' #Replace with actual recipients
    #CC_EMAIL = 'cc_user@db.com'
    #BCC_EMAIL = 'bcc_user@db.com'

    if not os.path.exists(excel_filepath):
        print(f"❌ Error: Excel file not found at {excel_filepath}. Cannot send email.")
        return

    outlook_app = None
    excel_app = None
    excel_workbook = None

    try:
        try:
            outlook_app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception:
            outlook_app = win32com.client.Dispatch("Outlook.Application")

        if outlook_app is None:
            raise RuntimeError("Could not initialize Outlook application.")

        # Initialize xlwings to open the generated Excel file
        excel_app = xw.App(visible=False, add_book=False)
        excel_workbook = excel_app.books.open(excel_filepath, read_only=True)
        dipv_sheet = excel_workbook.sheets["DIPV"]

        cob_date_str = cob_date.strftime('%d-%m-%Y') #Format the date for the email
        
        process_value = ""
        if dipv_check_value is not None:
            process_value = f" ({dipv_check_value})"

        # --- Extract Excel ranges into Pandas DataFrames and convert to HTML manually ---
        
        df_table1 = dipv_sheet.range('A1:AB25').options(pd.DataFrame, header=1, index=False).value
        
        # MODIFIED: Call new manual HTML generation function
        html_table1 = _dataframe_to_html_manual(df_table1, "CORE Rates")

        df_table2 = dipv_sheet.range('A28:AB42').options(pd.DataFrame, header=1, index=False).value
        
        # MODIFIED: Call new manual HTML generation function
        html_table2 = _dataframe_to_html_manual(df_table2, "CRU Details (Range A28:AB42)")
        
        # --- Common HTML Body Header ---
        common_html_header = (
            f"<!DOCTYPE html>"
            f"<html><head>{HTML_EMAIL_STYLES}</head><body>"
            f"<p>"
            f"Hi All,<br><br>"
            f"PFA DIPV summary for COB date: <b>{cob_date_str}</b>.<br><br>"
            f"Process: {process_value}<br><br>"
            f"</p>"
        )
        
        # --- Send First Email (CORE) ---
        mail_item_core = outlook_app.CreateItem(0)
        mail_item_core.To = RECIPIENT_EMAIL
        # mail_item_core.CC = CC_EMAIL
        # mail_item_core.BCC = BCC_EMAIL
        mail_item_core.Subject = f"DIPV Summary - CORE Rates - COB {cob_date_str}" 
        mail_item_core.HTMLBody = (
            f"{common_html_header}"
            f"<h3>CORE Rates:</h3>"
            f"{html_table1}" # This will now be the raw HTML string
            f"<p>"
            f"<b>Note:</b> This view <b>does not</b> incorporate any adjustment numbers and"
            f" this report has been generated utilizing data derived from pfo.(U1953,U1954,U11192,U828,U3200,U11119,U565), pfr.(IPVVarianceRiskPLE, IPVVarianceBond, IPVVarianceListed) under 'EOD_IP' MktTag for DB(LON,NYC).<br><br>"
            f"Thanks,<br>"
            f"Manish Goyal"
            f"</p>"
            f"</body></html>"
        )
        mail_item_core.Attachments.Add(excel_filepath) 
        mail_item_core.Send()
        print(f"✅ CORE Rates Email sent successfully.")
        
        # --- Send Second Email (CRU) ---
        mail_item_cru = outlook_app.CreateItem(0)
        mail_item_cru.To = RECIPIENT_EMAIL
        # mail_item_cru.CC = CC_EMAIL
        # mail_item_cru.BCC = BCC_EMAIL
        mail_item_cru.Subject = f"DIPV Summary - CRU Rates - COB {cob_date_str}" 
        mail_item_cru.HTMLBody = (
            f"{common_html_header}"
            f"<h3>CRU Rates:</h3>"
            f"{html_table2}" # This will now be the raw HTML string
            f"<p>"
            f"<b>Note:</b> This view <b>does not</b> incorporate any adjustment numbers and"
            f" this report has been generated utilizing data derived from pfo.(U1953,U1954,U11192,U828,U3200,U11119,U565), pfr.(IPVVarianceRiskPLE, IPVVarianceBond, IPVVarianceListed) under 'EOD_IP' MktTag for DB(LON,NYC).<br><br>"
            f"Thanks,<br>"
            f"Manish Goyal"
            f"</p>"
            f"</body></html>"
        )
        mail_item_cru.Attachments.Add(excel_filepath) 
        mail_item_cru.Send()
        print(f"✅ CRU Rates Email sent successfully.")

    except Exception as e:
        print(f"❌ An error occurred during Outlook email sending: {e}")
        import traceback
        print("\n--- Full Traceback ---")
        traceback.print_exc()
        print("--- End Traceback ---")
    finally:
        # Clean up Excel objects
        if excel_workbook:
            try:
                excel_workbook.close()
            except Exception as e: print(f"Warning: Error closing workbook for read: {e}")
        if excel_app:
            try:
                excel_app.quit()
            except Exception as e:
                print(f"Warning: Error quitting Excel app: {e}")

        # Explicitly delete mail_item objects if they were created
        if 'mail_item_core' in locals() and mail_item_core is not None:
            try: del mail_item_core
            except Exception as e: print(f"Warning: Error cleaning up mail_item_core COM object: {e}")
        if 'mail_item_cru' in locals() and mail_item_cru is not None:
            try: del mail_item_cru
            except Exception as e: print(f"Warning: Error cleaning up mail_item_cru COM object: {e}")
        if outlook_app:
            try:
                del outlook_app
            except Exception as e:
                print(f"Warning: Error releasing Outlook COM object: {e}")

# --- Helper function to find the Nth previous business day ---
def get_n_previous_business_day(start_date: datetime.date, n: int) -> datetime.date:
#Calculates the date of the Nth previous business day from start_date
    
    current_date = start_date
    business_days_found = 0
    while business_days_found < n:
        current_date -= timedelta(days=1) # Monday=0, Sunday=6
        if current_date.weekday() < 5:  # Check if it's a weekday
            business_days_found += 1
    return current_date

# --- Function to validate and determine COB date from user ---
def validate_and_get_cob_date_for_dipv(input_value: str) -> Union[datetime.date, None]:
#Validates user input for COB date. Returns a datetime.date object if valid, else None.

    input_value = str(input_value).strip()

    if input_value == '0':
        try:
            today_date = date.today()
            determined_dt = get_n_previous_business_day(today_date, 2)
            print(f"Using LATEST COB date (T-2): {determined_dt.strftime('%Y-%m-%d')}")
            return determined_dt
        except Exception as e:
            print(f"❌ ERROR: Cannot determine latest COB date - {e}. Aborting.")
            return None
    else:
        try:
            cob_dt_obj = datetime.datetime.strptime(input_value, "%d%m%Y").date()
            print(f"Proceeding with COB date: {cob_dt_obj.strftime('%Y-%m-%d')}")
            return cob_dt_obj
        except ValueError:
            print(f"❌ Invalid COB date input: '{input_value}'. Expected '0' for latest or 'DDMMYYYY'. Aborting operation.")
            return None

last_generated_excel_filepath = None
last_determined_cob_date = None

def on_run_process_btn_clicked(b):
    global last_generated_excel_filepath, last_determined_cob_date
    print("\n--- Starting DIPV Process ---")
    
    start_time = time.time()
    now = datetime.now()
    start_script_time = now.strftime('%H:%M:%S')
    print(f'\nScript started at {start_script_time}')
                  
    determined_cob_date = validate_and_get_cob_date_for_dipv(cob_date_input.value)

    if determined_cob_date:
        last_determined_cob_date = determined_cob_date        
        print(f"Final COB Date for processing: {determined_cob_date.strftime('%Y-%m-%d')}")
   
        generated_filepath = automate_pivot_export_to_excel(determined_cob_date, SAVE_FOLDER, PIVOT_TABLE_CONFIGS_TEMPLATE)

        # Check if automate_pivot_export_to_excel returned a valid path
        if generated_filepath:
            last_generated_excel_filepath = generated_filepath
        else:
            last_generated_excel_filepath = None # Ensure it's None
    else:
        print("DIPV Process aborted due to invalid COB date input.")
        last_determined_cob_date = None # Ensure it's None
        last_generated_excel_filepath = None # Ensure it's None
              
    now = datetime.now()
    end_script_time = now.strftime('%H:%M:%S')
    print(f'\nScript ended at {end_script_time}')

    end_time = time.time()
    print(f"Execution time: {end_time - start_time:.4f} seconds")
              
    print("\n--- DIPV Process Finished ---")

def on_send_emails_btn_clicked(b):
    print("\n--- Initiating Email Sending ---")

    # Determine the COB date using the same logic as the main process
    determined_cob_date = validate_and_get_cob_date_for_dipv(cob_date_input.value)

    if not determined_cob_date:
        print("❌ Email sending aborted: Could not determine a valid COB date from input.")
        print("--- Email Sending Attempt Finished ---")
        return

    # Construct the expected file path based on the determined COB date
    output_filename = f"DIPV_{determined_cob_date.strftime('%Y_%m_%d')}.xlsx"
    expected_filepath = os.path.join(SAVE_FOLDER, output_filename)

    print(f"Attempting to send email for COB: {determined_cob_date.strftime('%Y-%m-%d')}")

    # Check if the file exists before attempting to send
    if os.path.exists(expected_filepath): 
        email_check_value = None
        excel_app_read = None 
        excel_workbook_read = None
        try:
            excel_app_read = xw.App(visible=False, add_book=False)
            excel_workbook_read = excel_app_read.books.open(expected_filepath, read_only=True)
            dipv_sheet_read = excel_workbook_read.sheets["DIPV"] 
            email_check_value = dipv_sheet_read.range('AE6').value
            print(f"Process was: *{email_check_value}*")
          
        except Exception as e:
            if isinstance(e, KeyError):
                print(f"❌ Error: Sheet 'DIPV' not found in '{os.path.basename(expected_filepath)}'. Cannot read AE6(process) value.")
            else:
                print(f"❌ Error reading value from DIPV!AE6 in {expected_filepath}: {e}")
            email_check_value = None 
        finally:
            if excel_workbook_read:
                try:
                    excel_workbook_read.close()
                except Exception as e: print(f"Warning: Error closing workbook for read: {e}")
            if excel_app_read:
                try:
                    excel_app_read.quit()
                except Exception as e: print(f"Warning: Error quitting Excel app for read: {e}")
        
        # Now call the modified send_dipv_emails
        send_dipv_emails(expected_filepath, determined_cob_date, email_check_value)
    else:
        print(f"❌ Error: Excel file not found at the expected path: {expected_filepath}")
        print("Please ensure the 'DIPV Process' has been successfully run for this date, file name is also case sensitive")

    print("--- Email Sending Attempt Finished ---")

# --- Jupyter Widgets for User Interface ---

# Text input widget for COB date
cob_date_input = widgets.Text(
    value='',
    placeholder='DDMMYYYY (e.g., 26102025) or 0 for LATEST',
    description='COB Date:',
    disabled=False,
    layout=widgets.Layout(width='400px')
)

# Button to trigger the entire DIPV process
run_process_btn = widgets.Button(
    description="Run DIPV Process",
    button_style='success',
    tooltip="Generates DIPV view for the COB date entered above.",
    layout=widgets.Layout(width='400px', height='60px')
)
run_process_btn.style.font_weight = 'bold'

# Assign the event handler to the button
run_process_btn.on_click(on_run_process_btn_clicked)

# Button to send emails
send_email_btn = widgets.Button(
    description="Send DIPV Email",
    button_style='info',
    tooltip="Sends email with DIPV Excel report as attachment.",
    layout=widgets.Layout(width='400px', height='60px')
)
send_email_btn.style.font_weight = 'bold'

# Assign the event handler to the button
send_email_btn.on_click(on_send_emails_btn_clicked)

# Display the input widget and buttons in the Jupyter Notebook interface
print("Please enter a COB Date and click the button to run the process.")
display(cob_date_input, run_process_btn)
display(send_email_btn)


