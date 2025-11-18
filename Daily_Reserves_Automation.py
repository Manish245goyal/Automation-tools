
# =============================================================================
# SECTION 1: Imports & Configuration
# =============================================================================
import pandas as pd
import xlwings as xw
import time
from datetime import datetime
import os
import re
import win32com.client as win32
from IPython.display import display, HTML
import ipywidgets as widgets

# --- User Configuration ---
# Folder where processed Evidence Templates are saved.
SAVE_FOLDER = r"C:\Users\goyaman\OneDrive - Deutsche Bank AG\Desktop\Reserves Automatation" 

# Local folder where Booked and Calculated reserves files are saved from Outlook.
folder = r"C:\Users\goyaman\OneDrive - Deutsche Bank AG\Outlook Files" 

# Contents in the Evidence Template (Case sensitive).
BOOKED_SHEET = "Booked Reserves"
CALC_SHEET = "Calculated Reserves"
CHECK_SHEET = "Check"
CHECK_CELL = "F3" # Cell containing the completion status in the "Check" sheet.

# Email Recipients
RECIPIENT_EMAIL = "manish.goyal@db.com"   # Email of performer for process incomplete notifications.
EMAIL_CC = "manish.goyal@db.com"          # CC for summary emails.
EMAIL_TO_EMEA = "manish.goyal@db.com"     # To for EMEA summary emails.
EMAIL_TO_AMER = "manish.goyal@db.com"     # To for AMER summary emails.

# Main sheets present in EMEA/AMER Evidence Template for conditional checks.
SHEET_MAP_EMEA = ["FVReserves Template 1953", "FVReserves Template 11192", "FVReserves Template 828"]
SHEET_MAP_AMER = ["FVReserves Template 1954", "FVReserves Template 565"]

# Subject prefix for summary emails.
SUBJECT_PREFIX_PATTERN = "Booked and Calculated Reserves with Proposed Adjustments"

# =============================================================================
# SECTION 2: Core File Handling Functions
# =============================================================================

def get_reserve_file_path(search_folder, type_name="booked", target_cob_date_str=None):
    """
    Retrieves the file path for the latest reserve file (booked/calculated) or, 
    for a specific target_cob_date_str (YYYY-MM-DD format).
    """
    files = os.listdir(search_folder)
    pattern = f"{type_name}_reserves_"
    
    candidate_files = [f for f in files if pattern in f.lower()]
    
    if not candidate_files:
        raise FileNotFoundError(f"No {type_name} reserves files found in {search_folder}")
        
    file_info = []
    for f in candidate_files:
        match = re.search(r'_(\d{4})_(\d{2})_(\d{2})_', f)
        if match:
            dt_obj = datetime.strptime(f"{match.group(1)}-{match.group(2)}-{match.group(3)}", "%Y-%m-%d")
            file_info.append((dt_obj, os.path.join(search_folder, f)))
    
    if not file_info:
        raise FileNotFoundError(f"No {type_name} reserves files with valid date format found in {search_folder}")

    file_info.sort(key=lambda x: x[0], reverse=True)

    if target_cob_date_str:
        target_dt_obj = datetime.strptime(target_cob_date_str, "%Y-%m-%d")
        for dt, f_path in file_info:
            if dt == target_dt_obj:
                return f_path, target_cob_date_str
        raise FileNotFoundError(f"No {type_name} reserves file found for COB date {target_cob_date_str} in {search_folder}")
    else:
        latest_dt_obj, latest_f_path = file_info[0]
        return latest_f_path, latest_dt_obj.strftime("%Y-%m-%d")


def read_local_reserve_files(target_cob_date_str=None):
    """
    Reads the latest or specific booked and calculated reserve files into DataFrames.
    """
    booked_file_path, cob_date = get_reserve_file_path(folder, type_name="booked", target_cob_date_str=target_cob_date_str)
    calc_file_path, _ = get_reserve_file_path(folder, type_name="calculated", target_cob_date_str=target_cob_date_str)
    
    booked_df = pd.read_excel(booked_file_path)
    calc_df = pd.read_excel(calc_file_path)
    
    print(f"Loaded reserves files for COB date: {cob_date}")
    return booked_df, calc_df, cob_date


def get_latest_template_file(search_folder, region="EMEA", cob_date_dt_obj=None):
    """
    Returns the path to the latest Evidence Template file (<= COB date) for a given region.
    """
    files = os.listdir(search_folder)
    region_upper = region.upper()
    pattern = f"{region_upper} - Evidence template.xlsx"
    
    region_template_files = [f for f in files if re.match(r'\d{8} - ' + re.escape(pattern), f)]
    
    if not region_template_files:
        raise FileNotFoundError(f"No {region_upper} Evidence template files found in {search_folder} matching 'DDMMYYYY - {region_upper} - Evidence template.xlsx' pattern.")
    
    file_info = []
    for f in region_template_files:
        match = re.match(r'(\d{8})', f)
        if match:
            dt_obj = datetime.strptime(match.group(1), "%d%m%Y")
            if cob_date_dt_obj and dt_obj > cob_date_dt_obj:
                continue
            file_info.append((dt_obj, os.path.join(search_folder, f)))
    
    if not file_info:
        if cob_date_dt_obj:
            raise FileNotFoundError(f"No {region_upper} Evidence template <= COB date {cob_date_dt_obj.strftime('%Y-%m-%d')} found in {search_folder}.")
        else:
            raise FileNotFoundError(f"No {region_upper} Evidence template files with valid date format found in {search_folder}.")
    
    file_info.sort(key=lambda x: x[0], reverse=True)
    
    return file_info[0][1]

# =============================================================================
# SECTION 3: Excel Automation & Update
# =============================================================================

def update_main_excel(booked_df, calc_df, cob_date_str, region="EMEA"):
    """
    Pastes datasets into the Evidence Template, refreshes workbook, inserts COB date,
    and saves with COB date. Includes error handling and Excel process management.
    """
    booked_df = booked_df.loc[:, booked_df.columns.notna()]
    calc_df = calc_df.loc[:, calc_df.columns.notna()]

    booked_df['Start'] = pd.to_datetime(booked_df['Start'], errors='coerce')

    cob_dt_obj = datetime.strptime(cob_date_str, "%Y-%m-%d")
    booked_df_potential_dates = booked_df[booked_df['Start'] <= cob_dt_obj].copy()
    booked_df_filtered = booked_df[booked_df['Start'] <= cob_dt_obj]

    if not booked_df_filtered.empty:
        date_counts = booked_df_potential_dates.groupby('Start').size().reset_index(name='count')        
        # Identify dates with more than 100 rows
        dates_with_sufficient_rows = date_counts[date_counts['count'] > 100]
        selected_start_date = None

        if not dates_with_sufficient_rows.empty:
            # If there are dates with sufficient rows, pick the latest among them
            selected_start_date = dates_with_sufficient_rows['Start'].max()
        else:
            # If no date has > 100 rows, then fall back to the absolute latest 'Start' date
            selected_start_date = booked_df_potential_dates['Start'].max()
            print(f"No 'Start' dates found with > 500 rows. Falling back to the absolute latest date: {selected_start_date.strftime('%Y-%m-%d')}.")
        
        # Filter booked_df_potential_dates to keep only rows corresponding to the selected_start_date
        booked_df_filtered = booked_df_potential_dates[booked_df_potential_dates['Start'] == selected_start_date].copy()
        print(f"Booked Reserves filtered to Start/Reporting date = {selected_start_date.strftime('%Y-%m-%d')} ({len(booked_df_filtered)} rows).")
    
    else:
        booked_df_filtered = pd.DataFrame(columns=booked_df.columns)
        print("⚠️ No Booked Reserves rows with Start <= COB. Sheet will be empty.")

    formatted_date_display = cob_dt_obj.strftime("%d-%m-%Y")
    formatted_date_file = cob_dt_obj.strftime("%d%m%Y")

    try:
        template_file_path = get_latest_template_file(SAVE_FOLDER, region, cob_dt_obj)
        # print(f"Using {region} template: {os.path.basename(template_file_path)}") # Debug/Optional
    except FileNotFoundError as e:
        print(f"❌ ERROR: {e}. Aborting {region} processing.")
        return False, formatted_date_display, None
    
    app = None
    wb = None
    file_to_open = template_file_path
    
    try:
        # Check and close if the file is already open in *any* associated xlwings instance
        for running_app in xw.apps:
            for book in running_app.books:
                if book.fullname == os.path.abspath(file_to_open):
                    print(f"⚠️ Warning: Closing already open workbook '{book.name}' before opening.") 
                    book.close() 
                    break 
            if book.fullname == os.path.abspath(file_to_open):
                break 
        
        # Ensure the file is writable *before* opening
        if os.path.exists(file_to_open):
            try:
                if not os.access(file_to_open, os.W_OK):
                    os.chmod(file_to_open, 0o666)
                    print(f"Removed read-only attribute from template: {os.path.basename(file_to_open)}.") # Debug/Optional
            except Exception as e:
                print(f"Warning: Could not change permissions for template {os.path.basename(file_to_open)}: {e}") # Debug/Optional

        app = xw.App(visible=False, add_book=False) 
        wb = app.books.open(file_to_open)
        
        sht_booked = wb.sheets[BOOKED_SHEET]
        sht_booked.clear_contents()
        sht_booked.range("A1").options(index=False, header=True).value = booked_df_filtered

        sht_calc = wb.sheets[CALC_SHEET]
        sht_calc.clear_contents()
        sht_calc.range("A1").options(index=False, header=True).value = calc_df

        sht_check = wb.sheets[CHECK_SHEET]
        sht_check.range("B6").value = cob_dt_obj

        try:
            wb.api.RefreshAll()
            app.calculate()
            print("Refreshing Excel files.. Wait for 3 seconds...") 
            time.sleep(3)
            # print("Excel refreshed.") # Debug/Optional
        except Exception as e:
            print(f"⚠️ Warning: Refresh/Calculate failed: {e}") 

        check_value = sht_check.range(CHECK_CELL).value
        is_complete = bool(check_value) if check_value is not None else False
        print(f"Check Status: *{is_complete}*") 

        save_path = os.path.join(SAVE_FOLDER, f"{formatted_date_file} - {region} - Evidence template.xlsx")
        
        # Ensure the target save location is writable if a file already exists there
        if os.path.exists(save_path):
            try:
                if not os.access(save_path, os.W_OK): 
                    os.chmod(save_path, 0o666)
                    # print(f"Removed read-only attribute from existing file at {os.path.basename(save_path)}.") # Debug/Optional
            except Exception as e:
                print(f"Warning: Could not change permissions for existing file {os.path.basename(save_path)}: {e}") # Debug/Optional

        wb.save(save_path)
        print(f"Template saved: {save_path}") 

        return is_complete, formatted_date_display, save_path
    except Exception as e:
        print(f"❌ Error during Excel processing for region {region} on file {os.path.basename(file_to_open)}: {e}") 
        return False, formatted_date_display, None
    finally:
        if wb: 
            try: wb.close()
            except Exception as e: print(f"Warning: Error closing workbook: {e}")
        if app: 
            try: app.quit()
            except Exception as e: print(f"Warning: Error quitting Excel app: {e}") 
        # print("Excel application closure attempt finished for update_main_excel.") # Debug/Optional


# =============================================================================
# SECTION 4: Email Notification Functions
# =============================================================================

def send_email_notification(cob_date_display):
    """Sends an email notification if the daily reserves process did not complete successfully."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = RECIPIENT_EMAIL
        mail.Subject = f"Process Incomplete - daily reserves for COB: ({cob_date_display})"
        mail.Body = (
            f"Hello,\n\n"
            f"The daily reserves automation process for COB {cob_date_display} did not complete successfully.\n"
            f"Please make sure that it wasn't a public holiday, Refer to 'Check' sheet of Evidence template for more details.\n\n"
            f"Regards,\n.py script"
        )
        mail.Send()
        print("*Process incomplete*. Email sent to performer.") 
    except Exception as e:
        print(f"❌ ERROR: Could not send email notification to {RECIPIENT_EMAIL}: {e}") 


# =============================================================================
# SECTION 5: Main Daily Reserves Execution Logic
# =============================================================================

def run_daily_reserves_automation(target_cob_date_str_to_process):
    """
    Main function to run the daily reserves process (read, update Excel, save) for a given COB date.
    Handles EMEA and AMER regions.
    """
    try:
        booked_df, calc_df, actual_cob_date_from_files_str = read_local_reserve_files(target_cob_date_str=target_cob_date_str_to_process)
        
        if actual_cob_date_from_files_str != target_cob_date_str_to_process:
             print(f"⚠️ Warning: Expected COB {target_cob_date_str_to_process} but found files for {actual_cob_date_from_files_str}. Using date from file.") 
             target_cob_date_str_to_process = actual_cob_date_from_files_str

        print("\n--- Processing EMEA ---") 
        is_complete_emea, cob_date_display_emea, save_path_emea = update_main_excel(
            booked_df, calc_df, target_cob_date_str_to_process, region="EMEA"
        )

        if not is_complete_emea:
            send_email_notification(cob_date_display_emea)
            return
        else:
            print(f"EMEA daily reserves process completed successfully for COB: {cob_date_display_emea}.") 

        print("\n--- Processing AMER ---") 
        is_complete_amer, cob_date_display_amer, save_path_amer = update_main_excel(
            booked_df, calc_df, target_cob_date_str_to_process, region="AMER"
        )

        if not is_complete_amer:
            send_email_notification(cob_date_display_amer)
            return
        else:
            print(f"AMER daily reserves process completed successfully for COB: {cob_date_display_amer}.\n") 

        #print("Daily Reserves Automation process finished.")
        msg1 = widgets.HTML(f"<b style='color:green;'> Daily Reserves Automation process finished for COB: {target_cob_date_str_to_process}.</b>")
        display(msg1)

    except FileNotFoundError as e:
        print(f"❌ ERROR: File not found - {e}. Please ensure the necessary files exist and paths are correct.") 
    except Exception as e:
        print(f"❌ An unexpected ERROR occurred during process: {e}") 


# =============================================================================
# SECTION 6: Conditional Move Check & Summary Email
# =============================================================================

def read_and_filter_sheets(file_path, sheet_list):
    """
    Reads specified sheets from an Evidence Template and filters rows where
    absolute 'Move by Reserve Net' exceeds 1,000,000.
    """
    result = {}
    try:
        for sheet in sheet_list:
            df = pd.read_excel(file_path, sheet_name=sheet)            
            
            if df.shape[0] > 7:
                df.columns = df.iloc[7]
                df = df.iloc[8:].reset_index(drop=True)
            else:
                # print(f"Warning: Sheet '{sheet}' in '{os.path.basename(file_path)}' has fewer than 8 rows. Skipping.") # Debug/Optional
                continue

            move_col = [c for c in df.columns if "Move by Reserve Net" in str(c)]
            if move_col:
                move_col = move_col[0]
                df[move_col] = (
                    df[move_col].astype(str).str.replace(",", "", regex=False).str.strip()
                    .replace(["", "-"], "0").astype(float)
                )
                filtered = df[df[move_col].abs() > 1000000]
                
                keep_cols_names = []
                original_indices = [0, 5, 6, 7, 8, 9, 11]
                for idx in original_indices:
                    if idx < len(df.columns):
                        keep_cols_names.append(df.columns[idx])
                
                filtered = filtered[keep_cols_names] 
                filtered = filtered.fillna("")
                
                if not filtered.empty:
                    result[sheet.replace("FVReserves Template ", "")] = filtered
        return result
    except Exception as e:
        print(f"Error reading and filtering sheets from {file_path}: {e}") 
        return {}


def build_html_table(region, sheet_dfs, cob_date_str):
    """
    Builds an HTML table string for the email body from filtered DataFrames.
    """
    # Define CSS styles for the HTML email
    styles = """
        <style>
            table {
                border-collapse: collapse;
                width: 100%;
                font-family: Aptos, Arial, Segoe UI, sans-serif;
                font-size: 13px;
                table-layout: fixed; /* CRUCIAL: Forces column widths to be respected */
            }
            th {
                background-color: #004080;
                color: white;
                text-align: center;
                padding: 6px;
                border: 1px solid #808080;
                word-wrap: break-word; 
                overflow-wrap: break-word;
                white-space: normal;
            }
            td {
                border: 1px solid #808080;
                padding: 6px;
                text-align: left;
                vertical-align: top;
                word-wrap: break-word; 
                overflow-wrap: break-word; 
                white-space: normal; 
                line-height: 1.2; /* Adjust line spacing if needed */
            }
            .negative { color: red; }
            p { font-family: Aptos, Arial, Segoe UI, sans-serif; font-size: 14.5px; }
        </style>
    """

    cob_date_display = datetime.strptime(cob_date_str, '%Y-%m-%d').strftime('%d-%m-%Y')
    html_content = (
        f"<p>"
        f"Hi team,<br>"
        f"Could you please review the below moves and validate them. Also, review the adjustments in the attached file, "
        f"for COB <b>{cob_date_display}</b>.<br><br>"
    )

    for sheet_name, df in sheet_dfs.items():
        if df.empty:
            continue

        df_formatted = df.copy()

        comment_col_name = None
        for col in df_formatted.columns:
            if "Comment" in str(col): 
                comment_col_name = col
                break
        
        for col in df_formatted.columns:
            if pd.api.types.is_numeric_dtype(df_formatted[col]):
                df_formatted[col] = df_formatted[col].apply(
                    lambda x: f"<span class='negative'>{x:,.0f}</span>" if x < 0 else f"{x:,.0f}"
                    if pd.notnull(x) else ""
                )
            else:
                df_formatted[col] = df_formatted[col].fillna("")
                # Inject Zero-Width Space (ZWSP) after every space in the Comment column for robust wrapping.
                if col == comment_col_name:
                    df_formatted[col] = df_formatted[col].astype(str).apply(
                        lambda x: x.replace(' ', ' \u200b') if isinstance(x, str) else x
                    )

        html_table_str = df_formatted.to_html(index=False, escape=False, classes='dataframe')

        # Customize table tag and set table-layout fixed directly in style attribute
        html_table_str = html_table_str.replace('<table border="1" class="dataframe">', '<table style="table-layout:fixed; width:100%;">')

        # --- Inject inline style with fixed width for the Comment column's TH and <col> tag ---
        if comment_col_name in df_formatted.columns:
            escaped_comment_col_name = re.escape(comment_col_name)
            comment_col_width = "250px" # <--- ADJUST THIS PIXEL VALUE (e.g., "250px") FOR COMMENT COLUMN WIDTH

            # Add <col> tags for robust column width control, respecting table-layout:fixed
            col_tags = ""
            for i, col_header in enumerate(df_formatted.columns):
                col_width_html_attr = ""
                if col_header == comment_col_name:
                    col_width_html_attr = f' width="{comment_col_width}"'
                col_tags += f'<col{col_width_html_attr}>'
            
            # Inject colgroup after the main <table> tag
            html_table_str = re.sub(r'(<table[^>]*>)', r'\1<colgroup>' + col_tags + r'</colgroup>', html_table_str, 1)

            # Target the <th> tag for the comment column and apply inline styles for header consistency
            html_table_str = re.sub(
                fr'(<th[^>]*>)\s*(?:<span[^>]*>)?({escaped_comment_col_name})(?:<\/span>)?\s*(<\/th>)',
                fr'\1<span style="width:{comment_col_width}; max-width:{comment_col_width}; white-space:normal; word-wrap:break-word; overflow-wrap:break-word; display:inline-block;">\2</span>\3',
                html_table_str, flags=re.IGNORECASE | re.DOTALL
            )

        html_content += f"<b>U{sheet_name}:</b>{html_table_str}<br>"

    html_content += (
        "<p>Thanks,<br>"
        "Manish Goyal"
        "</p>"
    )

    return styles + html_content


def send_conditional_email(region, file_path, sheet_dfs, cob_date_str):
    """Sends a summary email with conditional tables for large moves."""
    try:
        outlook = win32.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        cob_date_display = datetime.strptime(cob_date_str, '%Y-%m-%d').strftime('%d-%m-%Y')
        mail.Subject = f"{SUBJECT_PREFIX_PATTERN} for Date {cob_date_display}"

        mail.To = EMAIL_TO_EMEA if region == "EMEA" else EMAIL_TO_AMER
        mail.CC = EMAIL_CC

        has_large_moves = any(not df.empty for df in sheet_dfs.values())

        if not has_large_moves:
            mail.HTMLBody = (
                f"<p style='font-family:Aptos,Arial,Segoe UI;font-size:14.5px;'>"
                f"Hi All,<br>"
                f"Please review the adjustments in the attached file for COB <b>{cob_date_display}</b>.<br><br>"
                f"Thanks,<br>"
                f"Manish Goyal"
                f"</p>"
            )
        else:
            mail.HTMLBody = build_html_table(region, sheet_dfs, cob_date_str)

        mail.Attachments.Add(file_path)
        mail.Send()
        print(f"Summary Email sent successfully for {region}") 

    except Exception as e:
        print(f"❌ ERROR: Could not send {region} email: {e}") 


def run_conditional_move_check(target_cob_date_str_to_process):
    """
    Triggers the move analysis and Sends email for a specified COB date.
    """
    try:
        cob_date_dt_obj = datetime.strptime(target_cob_date_str_to_process, "%Y-%m-%d")
        cob_date_file_format = cob_date_dt_obj.strftime("%d%m%Y")

        emea_file = os.path.join(SAVE_FOLDER, f"{cob_date_file_format} - EMEA - Evidence template.xlsx")
        amer_file = os.path.join(SAVE_FOLDER, f"{cob_date_file_format} - AMER - Evidence template.xlsx")

        if not os.path.exists(emea_file) or not os.path.exists(amer_file):
            print(f"❌ Processed files not found for COB {target_cob_date_str_to_process}. Please ensure the 'Run Daily Reserves' step was completed successfully for this date. No Email sent!") 
            return

        print(f"\nChecking for move > ±1M for COB Date: {target_cob_date_str_to_process}") 
        
        emea_dfs = read_and_filter_sheets(emea_file, SHEET_MAP_EMEA)
        amer_dfs = read_and_filter_sheets(amer_file, SHEET_MAP_AMER)
        
        send_conditional_email("EMEA", emea_file, emea_dfs, target_cob_date_str_to_process) 
        send_conditional_email("AMER", amer_file, amer_dfs, target_cob_date_str_to_process) 

        msg = widgets.HTML(
            f"<b style='color:green;'> Daily Reserves email sent successfully for COB {target_cob_date_str_to_process}.</b>"
        )
        display(msg) 
    except FileNotFoundError as e:
        print(f"❌ ERROR in email check: {e}") 
    except Exception as e:
        print(f"❌ An unexpected ERROR occurred during conditional check: {e}") 
    finally:
        for app_instance in xw.apps:
            try:
                app_instance.quit()
            except Exception as e:
                # print(f"Warning: Error quitting Excel app during cleanup: {e}") # Debug/Optional
                pass # Suppress warning for clean exit
        # print("xlwings cleanup finished for run_conditional_move_check.") # Debug/Optional


# In[11]:


# =============================================================================
# SECTION 7: Jupyter Widgets for User Interface
# =============================================================================

# Single Input widget for COB date
cob_date_master_input = widgets.Text(
    value='',
    placeholder='DDMMYYYY (e.g., 26102024) or 0 for LATEST',
    description='COB Date:',
    disabled=False,
    layout=widgets.Layout(width='400px')
)

# Button to trigger the daily reserves process
run_daily_reserves_btn = widgets.Button( 
    description="Run Daily Reserves",
    button_style='success',
    tooltip="Processes files and saves new templates for the COB date entered above.",
    layout=widgets.Layout(width='400px', height='60px')
)

# Button for 'Send Summary Email'
send_summary_email_btn = widgets.Button(
    description="Send Summary Email", 
    button_style='warning',
    tooltip="Checks for move > ±1m and sends email for the COB date entered above.",
    layout = widgets.Layout(width='400px', height='60px'),
)
send_summary_email_btn.style.font_weight = 'bold'


def validate_and_get_cob_date(input_value):
    """
    Validates user input. Returns YYYY-MM-DD string if valid, else None.
    Prints error messages for invalid inputs.
    """
    input_value = input_value.strip()

    if input_value == '0':
        try:
            _, determined_cob_date_str = get_reserve_file_path(folder, type_name="booked", target_cob_date_str=None)
            print(f"LATEST COB date available: {determined_cob_date_str}") 
            return determined_cob_date_str
        except FileNotFoundError as e:
            print(f"❌ ERROR: Cannot determine latest COB date - {e}. Aborting.") 
            return None
    else:
        try:
            cob_dt_obj = datetime.strptime(input_value, "%d%m%Y")
            determined_cob_date_str = cob_dt_obj.strftime("%Y-%m-%d")
            print(f"Proceeding with COB date: {determined_cob_date_str}") 
            return determined_cob_date_str
        except ValueError:
            print(f"❌ Invalid COB date input: '{input_value}'. Expected '0' for latest or 'DDMMYYYY'. Aborting operation.") 
            return None

def on_run_daily_reserves_btn_clicked(b):
    """Callback for 'Run Daily Reserves' button click."""
    print("\n**Starting Daily Reserves Process**") 
    for app_instance in xw.apps:
        try:
            app_instance.quit()
        except Exception as e:
            # print(f"Warning: Error quitting Excel app during pre-run cleanup: {e}") # Debug/Optional
            pass # Suppress warning for clean exit
    
    determined_cob_date_str = validate_and_get_cob_date(cob_date_master_input.value)
    if determined_cob_date_str:
        run_daily_reserves_automation(determined_cob_date_str)
    else:
        print("Daily Reserves Process aborted due to invalid COB date input.") 


def on_send_summary_email_btn_clicked(b):
    """Callback for 'Send Summary Email' button click."""
    print("\n***Started Framing Summary Email***") 
    for app_instance in xw.apps:
        try:
            app_instance.quit()
        except Exception as e:
            # print(f"Warning: Error quitting Excel app during pre-email cleanup: {e}") # Debug/Optional
            pass # Suppress warning for clean exit
            
    determined_cob_date_str = validate_and_get_cob_date(cob_date_master_input.value)
    if determined_cob_date_str:
        run_conditional_move_check(determined_cob_date_str)
    else:
        print("Summary Email operation aborted due to invalid COB date input.") 


# Assign the event handlers
run_daily_reserves_btn.on_click(on_run_daily_reserves_btn_clicked)
send_summary_email_btn.on_click(on_send_summary_email_btn_clicked)

# Display the single input and both buttons in the Jupyter Notebook interface
display(cob_date_master_input, run_daily_reserves_btn, send_summary_email_btn)


# PFB the 'VBA script' for downloading Booked/Calculated reserves files for desired COB date... 
# press ALT+F11 and paste below code after commenting out select all and (Ctrl + /)... you can also create a macro button to ease the process execution..


# Code starts from below

# Option Explicit
# '================== USER CONFIGURATION =================='
# ' Path where downloaded attachments will be saved. Ensure this folder exists.
# Const SAVE_PATH As String = "C:\Users\goyaman\OneDrive - Deutsche Bank AG\Outlook Files\"

# Const SENDER_FILTER As String = "Kannon-noreply@discard.mail.db.com"
# Const SUBJECT_PREFIX_PATTERN As String = "Booked and Calculated Reserves with Proposed Adjustments"

# Private g_SearchForSpecificDate As Boolean
# ' Stores the user-specified COB Date (as a Date object). Only valid if g_SearchForSpecificDate is True.
# Private g_TargetCOBDate As Date
# Private g_TargetCOBDate_SubjectFormat As String
# ' -----------------------------------------------------------------


# Public Sub Download_Reserves_Files_byCOB()
#     On Error GoTo ErrHandler
    
#     Dim ns As Outlook.NameSpace
#     Dim rootFolder As Outlook.folder
#     Dim targetMail As Outlook.MailItem
#     Dim inputResponse As String ' Captures the user's input string from InputBox
#     Dim dtLatestMatch As Date   ' Used only if g_SearchForSpecificDate is False

#     ' --- Reset global flags/variables for each run ---
#     g_SearchForSpecificDate = False
#     g_TargetCOBDate = CDate("1/1/1900") ' Initialize to a very old date
#     g_TargetCOBDate_SubjectFormat = ""
#     dtLatestMatch = CDate("1/1/1900")   ' Initialize to a very old date for latest search
#     ' -------------------------------------------------

#     inputResponse = InputBox("Enter the COB date (DDMMYYYY)" & vbCrLf & _
#                              "Enter '0' (zero) to search for the LATEST available files.", "Enter COB Date or 0 for LATEST")
    
#     If inputResponse = "" Then
#         MsgBox "Operation cancelled by user.", vbInformation, "UBR Automation"
#         Exit Sub
#     End If
    
#     inputResponse = Trim(inputResponse)
    
#     If inputResponse = "0" Then
#         g_SearchForSpecificDate = False
#     Else
#         g_SearchForSpecificDate = True
        
#         If Len(inputResponse) <> 8 Or Not IsNumeric(inputResponse) Then
#             MsgBox "Invalid input: Expected '0' for latest or 8 digits for DDMMYYYY." & vbCrLf & _
#                    "Operation aborted.", vbCritical, "UBR Automation"
#             Exit Sub
#         End If
        
#         Dim sDay As String: sDay = Left(inputResponse, 2)
#         Dim sMonth As String: sMonth = Mid(inputResponse, 3, 2)
#         Dim sYear As String: sYear = Right(inputResponse, 4)
        
#         Dim parsedDay As Integer
#         Dim parsedMonth As Integer
#         Dim parsedYear As Integer
        
#         parsedDay = CInt(sDay)
#         parsedMonth = CInt(sMonth)
#         parsedYear = CInt(sYear)
        
#         If Not IsDate(parsedMonth & "/" & parsedDay & "/" & parsedYear) Then
#             MsgBox "The entered date '" & inputResponse & "' is not a valid calendar date." & vbCrLf & _
#                    "Operation aborted.", vbCritical, "UBR Automation"
#             Exit Sub
#         End If
        
#         ' Valid date confirmed. Store it globally.
#         g_TargetCOBDate = DateSerial(parsedYear, parsedMonth, parsedDay)
#         g_TargetCOBDate_SubjectFormat = Format(g_TargetCOBDate, "YYYY-MM-DD")
#     End If
    
#     ' --- Initialize Outlook NameSpace and Root Folder ---
#     Set ns = Application.GetNamespace("MAPI")
#     Set rootFolder = ns.GetDefaultFolder(olFolderInbox).Parent  ' Access the entire mailbox structure
    
#     Set targetMail = FindMatchingMail(rootFolder, dtLatestMatch)
    
#     ' --- Check if a matching email was found ---
#     If targetMail Is Nothing Then
#         If g_SearchForSpecificDate Then
#             MsgBox "? No mail found for COB date " & Format(g_TargetCOBDate, "DD-MM-YYYY") & " anywhere in your mailbox that matches the criteria (sender, subject prefix, and date).", vbCritical, "UBR Automation"
#         Else
#             MsgBox "? No matching mail found anywhere in your mailbox that fits the criteria (sender and subject prefix).", vbCritical, "UBR Automation"
#         End If
#         Exit Sub
#     End If
    
#     ' --- Process and Save Attachments ---
#     Dim att As Attachment
#     Dim savedCount As Long
#     savedCount = 0
    
#     For Each att In targetMail.Attachments
#         ' Check if attachment filename contains "booked_reserves" or "calculated_reserves" (case-insensitive)
#         If InStr(1, att.FileName, "booked_reserves", vbTextCompare) > 0 _
#            Or InStr(1, att.FileName, "calculated_reserves", vbTextCompare) > 0 Then
#             att.SaveAsFile SAVE_PATH & att.FileName ' Save the attachment
#             savedCount = savedCount + 1 ' Increment counter for saved attachments
#         End If
#     Next att
    
#     ' --- Provide user feedback on attachment download status ---
#     If savedCount > 0 Then
#         MsgBox "Download complete!" & vbCrLf & _
#                "From: " & targetMail.SenderEmailAddress & vbCrLf & _
#                "Subject: " & targetMail.Subject & vbCrLf & _
#                "Attachments saved: " & savedCount & vbCrLf & _
#                "Saved to: " & SAVE_PATH, vbInformation, "Reserve Automation"
#     Else
#         Dim searchDescription As String
#         If g_SearchForSpecificDate Then
#             searchDescription = "for COB date " & Format(g_TargetCOBDate, "DD-MM-YYYY")
#         Else
#             searchDescription = "in the latest matching mail found"
#         End If
#         MsgBox "?? No matching attachments (booked_reserves or calculated_reserves) found " & searchDescription & ".", vbExclamation
#     End If
    
#     Exit Sub

# ' --- Error Handler ---
# ErrHandler:
#     MsgBox "? An unexpected Error occurred: " & Err.Description & vbCrLf & _
#            "Please ensure Outlook is running and configured correctly.", vbCritical, "UBR Automation"
# End Sub


# '================== UNIFIED RECURSIVE SEARCH FUNCTION =================='
# Private Function FindMatchingMail(ByVal folder As Outlook.folder, ByRef dtLatestMatch As Date) As Outlook.MailItem
#     Dim itm As Object
#     Dim subFolder As Outlook.folder
#     Dim tempMatchMail As Outlook.MailItem ' Used to hold a mail found in a recursive call
    
#     Dim baseSubjectFilter As String
#     baseSubjectFilter = SUBJECT_PREFIX_PATTERN ' Subject must at least contain this pattern

#     For Each itm In folder.Items
#         If itm.Class = olMail Then
#             If InStr(1, itm.SenderEmailAddress, SENDER_FILTER, vbTextCompare) > 0 _
#                 And InStr(1, itm.Subject, baseSubjectFilter, vbTextCompare) > 0 Then
                
#                 If g_SearchForSpecificDate Then
#                     Dim subjectToMatch_DatePart As String
#                     subjectToMatch_DatePart = " Date " & g_TargetCOBDate_SubjectFormat
                    
#                     If InStr(1, itm.Subject, subjectToMatch_DatePart, vbTextCompare) > 0 Then
#                         Set FindMatchingMail = itm ' Found the specific mail!
#                         Exit Function ' Exit immediately once found - no need to search further
#                     End If
#                 Else
#                     If itm.ReceivedTime > dtLatestMatch Then
#                         Set FindMatchingMail = itm
#                         dtLatestMatch = itm.ReceivedTime ' Update the latest time found so far (passed by reference)
#                     End If
#                 End If
#             End If
#         End If
#     Next itm ' Continue to the next item in the current folder

#     For Each subFolder In folder.Folders
#         Set tempMatchMail = FindMatchingMail(subFolder, dtLatestMatch)
        
#         If Not tempMatchMail Is Nothing Then
#             If g_SearchForSpecificDate Then
#                 ' If searching for a specific date and it was found in a subfolder,
#                 ' propagate that mail object up the call stack immediately.
#                 Set FindMatchingMail = tempMatchMail
#                 Exit Function ' Exit immediately once found
#             Else
#                 If tempMatchMail.ReceivedTime = dtLatestMatch Then ' This means tempMatchMail is the true latest
#                     Set FindMatchingMail = tempMatchMail
#                 End If
#             End If
#         End If
#     Next subFolder ' Continue to the next subfolder
    
# End Function
