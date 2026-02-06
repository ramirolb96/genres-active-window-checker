import pandas as pd
import os
import shutil
import tempfile
import uuid
import sys
from datetime import datetime, timedelta

# ==============================================================================
# --- CONFIGURATION ---
# ==============================================================================

# 1. FILE PATHS
# Master List (Static)
PATH_MASTER = r"C:\Users\RLozadaBillot\Sony Pictures Entertainment\Sony One Planning Strategy - Master List\Master List Amazon_Movies 2.0_fy26.xlsx"

# Output Path (Desktop)
PATH_OUTPUT = r"C:\Users\RLozadaBillot\OneDrive - Sony Pictures Entertainment\Desktop\Genre_Active_Window_Report.xlsx"

# 2. MASTER LIST SETTINGS
SHEET_MASTER = "FY 26 ACTIVE"
HEADER_ROW_MASTER = 3  # Row 4 in Excel
MASTER_IDX_TITLE = 2   # Col C (Title)
MASTER_IDX_GENRE = 11  # Col L (Main Genre) - 0-based index (A=0 ... L=11)

# Windows Config (Same as previous script)
MASTER_WINDOWS = [ (15, 16), (18, 19), (21, 22), (24, 25) ]

# ==============================================================================
# --- UTILITY FUNCTIONS ---
# ==============================================================================

def clean_text(val):
    if pd.isna(val): return ""
    return str(val).strip() # keeping case for Genre display, will upper for comparison if needed

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
    
    # Extract unique genres from Col L (Index 11)
    # Filter out NaNs/Empties
    raw_genres = df_master.iloc[:, MASTER_IDX_GENRE].dropna().astype(str).unique()
    
    # Clean and sort
    genres = sorted([g.strip() for g in raw_genres if g.strip() != ""])
    
    if not genres:
        print("No genres found in the Master List column L.")
        return None

    # Display Options
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

    print(f"\nSearching for '{selected_genre}' titles active from {format_date_str(user_start)} to {format_date_str(user_end)}...")

    matches = []

    # 2. Iterate and Filter
    for idx, row in df_master.iterrows():
        # Check Genre First (Fastest check)
        if MASTER_IDX_GENRE >= len(row): continue
        
        row_genre = clean_text(row.iloc[MASTER_IDX_GENRE])
        if row_genre.upper() != selected_genre.upper():
            continue
            
        # Check Title
        if MASTER_IDX_TITLE >= len(row): continue
        title = row.iloc[MASTER_IDX_TITLE]
        if pd.isna(title): continue
        
        # Check Dates
        valid_blocks = get_all_valid_blocks(row)
        
        is_active_during_window = False
        supporting_block = None
        
        for v_start, v_end in valid_blocks:
            # Logic: The USER window must be fully inside the VALID block
            # i.e. Block Start <= User Start AND Block End >= User End
            if v_start <= user_start and v_end >= user_end:
                is_active_during_window = True
                supporting_block = (v_start, v_end)
                break
        
        if is_active_during_window:
            matches.append({
                "Title": title,
                "Main Genre": selected_genre,
                "Requested Window Start": format_date_str(user_start),
                "Requested Window End": format_date_str(user_end),
                "Valid Master Window": f"{format_date_str(supporting_block[0])} to {format_date_str(supporting_block[1])}",
                "Status": "ACTIVE"
            })

    # 3. Output
    if matches:
        print(f"\nFound {len(matches)} matching titles.")
        try:
            with pd.ExcelWriter(PATH_OUTPUT, engine='openpyxl') as writer:
                df_out = pd.DataFrame(matches)
                df_out.to_excel(writer, index=False, sheet_name="Genre Report")
                
                print("Autofitting columns...")
                autofit_columns(writer)
                
            print(f"Report saved successfully: {PATH_OUTPUT}")
            os.startfile(PATH_OUTPUT)
        except Exception as e:
            print(f"Error saving report: {e}")
    else:
        print("\nNo titles found matching criteria.")

if __name__ == "__main__":
    run_genre_checker()
