import pandas as pd
import os
import shutil
import tempfile
import uuid
import sys
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta

# ==============================================================================
# --- CONFIGURATION & SETUP ---
# ==============================================================================

def find_desktop():
    """
    Tries to find the actual Desktop path, accounting for OneDrive syncs.
    """
    home = os.path.expanduser("~")
    
    # Priority list of paths to check
    candidates = [
        os.path.join(home, "OneDrive - Sony Pictures Entertainment", "Desktop"),
        os.path.join(home, "OneDrive", "Desktop"),
        os.path.join(home, "Desktop")
    ]
    
    for path in candidates:
        if os.path.exists(path):
            return path
            
    # Fallback: If no desktop found, save to the user's home folder
    print("[WARNING] Could not find a Desktop folder. Saving to Home directory.")
    return home

def select_master_file():
    """Opens a file dialog for the user to select the Master List."""
    print("Please select the Master List Excel file from the popup window...")
    
    # Initialize Tkinter and hide the main window
    root = tk.Tk()
    root.withdraw() 
    root.attributes('-topmost', True) # Force the popup to the front
    
    file_path = filedialog.askopenfilename(
        title="Select the Master List Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xls")]
    )
    
    root.destroy() # Close Tkinter
    
    if not file_path:
        print("No file selected. Exiting.")
        sys.exit()
        
    print(f"File selected: {os.path.basename(file_path)}")
    return file_path

# 1. FILE PATHS
# Dynamic Selection (Pop-up window)
PATH_MASTER = select_master_file()

# Output Path (Smart Desktop Detection)
DESKTOP_PATH = find_desktop()
PATH_OUTPUT = os.path.join(DESKTOP_PATH, "Genre_Active_Window_Report.xlsx")

# 2. MASTER LIST SETTINGS
SHEET_MASTER = "FY 26 ACTIVE"
HEADER_ROW_MASTER = 3  # Row 4 in Excel
MASTER_IDX_TITLE = 2   # Col C (Title)
MASTER_IDX_GENRE = 11  # Col L (Main Genre) - 0-based index (A=0 ... L=11)
MASTER_IDX_GENRE_2 = 12 # Col M (Second Genre) - Index 12

# Windows Config
MASTER_WINDOWS = [ (15, 16), (18, 19), (21, 22), (24, 25) ]

# ==============================================================================
# --- UTILITY FUNCTIONS ---
# ==============================================================================

def clean_text(val):
    if pd.isna(val): return ""
    return str(val).strip() 

def parse_date(val):
    if pd.isna(val) or val == "" or str(val).strip().upper() == "EMPTY":
        return None
    try:
        return pd.to_datetime(val).normalize()
    except:
        return None

def format_date_str(date_val):
    if pd.isna(date_val): return ""
    try:
        return date_val.strftime('%m-%d-%Y') 
    except:
        return str(date_val)

def load_excel_safe(filepath, sheet_name_or_index, header_row):
    """Loads Excel safely using a temp copy to avoid file locking issues."""
    if not os.path.exists(filepath):
        print(f"[ERROR] File not found: {filepath}")
        return None

    temp_dir = tempfile.gettempdir()
    unique_name = f"audit_{uuid.uuid4().hex}.xlsx"
    temp_path = os.path.join(temp_dir, unique_name)
    
    try:
        shutil.copy2(filepath, temp_path)
        return pd.read_excel(temp_path, sheet_name=sheet_name_or_index, header=header_row)
    except Exception as e:
        print(f"[ERROR] Could not read file: {e}")
        return None
    finally:
        try: os.remove(temp_path)
        except: pass

def get_all_valid_blocks(row):
    """Parses row for all start/end dates and merges overlapping ranges."""
    raw_ranges = []
    for start_col, end_col in MASTER_WINDOWS:
        if start_col >= len(row) or end_col >= len(row): continue
        s = parse_date(row.iloc[start_col])
        e = parse_date(row.iloc[end_col])
        if s and e and e >= s:
            raw_ranges.append((s, e))

    if not raw_ranges: return []
    raw_ranges.sort(key=lambda x: x[0])

    merged = []
    curr_s, curr_e = raw_ranges[0]
    for next_s, next_e in raw_ranges[1:]:
        if next_s <= (curr_e + timedelta(days=1)):
            curr_e = max(curr_e, next_e)
        else:
            merged.append((curr_s, curr_e))
            curr_s, curr_e = next_s, next_e
    merged.append((curr_s, curr_e))
    return merged

def autofit_columns(writer):
    """Autofits columns in the final Excel output."""
    for sheet_name in writer.sheets:
        worksheet = writer.sheets[sheet_name]
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter 
            for cell in column:
                try:
                    if cell.value:
                        cell_len = len(str(cell.value))
                        if cell_len > max_length: max_length = cell_len
                except: pass
            
            adjusted_width = (max_length + 2)
            if adjusted_width > 100: adjusted_width = 100
            worksheet.column_dimensions[column_letter].width = adjusted_width

# ==============================================================================
# --- USER INPUT FUNCTIONS ---
# ==============================================================================

def get_user_dates():
    print("\n--- DEFINE ACTIVE WINDOW ---")
    while True:
        try:
            s_str = input("Enter Start Date (mm-dd-yyyy): ").strip()
            e_str = input("Enter End Date   (mm-dd-yyyy): ").strip()
            
            start_date = datetime.strptime(s_str, "%m-%d-%Y")
            end_date = datetime.strptime(e_str, "%m-%d-%Y")
            
            if end_date < start_date:
                print("Error: End date cannot be before start date. Please try again.")
                continue
                
            return start_date, end_date
        except ValueError:
            print("Invalid format. Please use mm-dd-yyyy (e.g., 05-20-2026).")

def get_user_genre(df_master):
    print("\n--- SELECT MAIN GENRE ---")
    print("Scanning Master List for unique genres...")
    
    unique_genres_set = set()
    
    # 1. Iterate through column L to build a clean set of individual genres
    raw_column = df_master.iloc[:, MASTER_IDX_GENRE].dropna().astype(str)
    
    for val in raw_column:
        # Normalize separators: Replace comma with semicolon
        normalized_val = val.replace(",", ";")
        # Split by semicolon
        parts = normalized_val.split(";")
        
        for p in parts:
            clean_p = p.strip()
            if clean_p:
                unique_genres_set.add(clean_p) 
    
    # 2. Sort list
    genres = sorted(list(unique_genres_set))
    
    if not genres:
        print("No genres found in the Master List column L.")
        return None

    # 3. Display Options
    for i, genre in enumerate(genres):
        print(f"{i + 1}) {genre}")
        
    while True:
        try:
            selection = input(f"\nSelect a number (1-{len(genres)}): ").strip()
            if not selection.isdigit():
                raise ValueError
            
            idx = int(selection) - 1
            if 0 <= idx < len(genres):
                selected_genre = genres[idx]
                print(f"Selected: {selected_genre}")
                return selected_genre
            else:
                print("Number out of range.")
        except:
            print("Invalid input. Please enter the number corresponding to the genre.")

# ==============================================================================
# --- MAIN LOGIC ---
# ==============================================================================

def run_genre_checker():
    print("Loading Master List... Please wait.")
    df_master = load_excel_safe(PATH_MASTER, SHEET_MASTER, HEADER_ROW_MASTER)
    
    if df_master is None:
        print("Could not load Master List.")
        return

    # 1. Get Inputs
    user_start, user_end = get_user_dates()
    selected_genre = get_user_genre(df_master)
    
    if not selected_genre:
        print("Genre selection failed. Exiting.")
        return

    print(f"\nSearching for titles containing '{selected_genre}' active from {format_date_str(user_start)} to {format_date_str(user_end)}...")

    matches_active = []
    matches_inactive = []

    # 2. Iterate and Filter
    for idx, row in df_master.iterrows():
        # A. Check Genre First
        if MASTER_IDX_GENRE >= len(row): continue
        
        raw_genre_str = clean_text(row.iloc[MASTER_IDX_GENRE])
        if not raw_genre_str: continue

        # Parse and clean row genres
        row_genre_list = [x.strip().upper() for x in raw_genre_str.replace(",", ";").split(";")]
        search_upper = selected_genre.upper()
        
        if search_upper not in row_genre_list:
            continue

        # --- GET SECOND GENRE (Col M) ---
        sec_genre_val = ""
        if MASTER_IDX_GENRE_2 < len(row):
             sec_genre_val = clean_text(row.iloc[MASTER_IDX_GENRE_2])
            
        # --- CALCULATE SORT SCORE ---
        match_index = row_genre_list.index(search_upper)
        is_sole_genre = (len(row_genre_list) == 1)
        
        if is_sole_genre:
            sort_score = 0
        else:
            sort_score = match_index + 1
            
        # B. Check Title
        if MASTER_IDX_TITLE >= len(row): continue
        title = row.iloc[MASTER_IDX_TITLE]
        if pd.isna(title): continue
        
        # C. Check Dates
        valid_blocks = get_all_valid_blocks(row)
        
        is_active_during_window = False
        supporting_block = None
        
        for v_start, v_end in valid_blocks:
            # Check if active window covers the FULL requested window
            if v_start <= user_start and v_end >= user_end:
                is_active_during_window = True
                supporting_block = (v_start, v_end)
                break
        
        # --- DATA PREP FOR ROWS ---
        row_data = {
            "Title": title,
            "Main Genre": selected_genre,           # Col B
            "Full Main Genre Values": raw_genre_str, # Col C
            "Second Genre": sec_genre_val,          # Col D
            "Requested Window Start": format_date_str(user_start),
            "Requested Window End": format_date_str(user_end),
            "Sort_Score": sort_score # Hidden helper column
        }

        if is_active_during_window:
            # --- ACTIVE LOGIC ---
            valid_str = f"[{format_date_str(supporting_block[0])}] to [{format_date_str(supporting_block[1])}]"
            row_data["Valid Master Window"] = valid_str
            row_data["Status"] = "ACTIVE"
            matches_active.append(row_data)
        else:
            # --- INACTIVE LOGIC ---
            if not valid_blocks:
                valid_str = "NO VALID DATES"
            else:
                block_strs = [f"[{format_date_str(s)}] to [{format_date_str(e)}]" for s, e in valid_blocks]
                valid_str = " OR ".join(block_strs)
            
            row_data["Valid Master Window"] = valid_str
            row_data["Status"] = "INACTIVE"
            matches_inactive.append(row_data)

    # 3. Output
    print(f"\nFound {len(matches_active)} ACTIVE titles.")
    print(f"Found {len(matches_inactive)} INACTIVE titles.")
    print(f"Attempting to save to: {PATH_OUTPUT}")
    
    try:
        with pd.ExcelWriter(PATH_OUTPUT, engine='openpyxl') as writer:
            
            # --- TAB 1: ACTIVE TITLES ---
            if matches_active:
                matches_active.sort(key=lambda x: x["Sort_Score"])
                df_active = pd.DataFrame(matches_active)
                df_active.drop(columns=["Sort_Score"], inplace=True)
                df_active.to_excel(writer, index=False, sheet_name="Active Titles")
            else:
                pd.DataFrame(["No active titles found"]).to_excel(writer, sheet_name="Active Titles", header=False)

            # --- TAB 2: INACTIVE TITLES ---
            if matches_inactive:
                matches_inactive.sort(key=lambda x: x["Sort_Score"])
                df_inactive = pd.DataFrame(matches_inactive)
                df_inactive.drop(columns=["Sort_Score"], inplace=True)
                df_inactive.to_excel(writer, index=False, sheet_name="Inactive Titles")
            else:
                pd.DataFrame(["No inactive titles found"]).to_excel(writer, sheet_name="Inactive Titles", header=False)
            
            print("Autofitting columns...")
            autofit_columns(writer)
            
        print(f"Report saved successfully: {PATH_OUTPUT}")
        os.startfile(PATH_OUTPUT)
    except Exception as e:
        print(f"Error saving report: {e}")

if __name__ == "__main__":
    run_genre_checker()