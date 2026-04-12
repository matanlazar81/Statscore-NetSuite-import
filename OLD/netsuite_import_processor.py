"""
NetSuite Import Processor
- Opens a GUI to select CSV file and Excel mapping file
- Identifies Intercompany (IC) transactions using the mapping file
- Uses REF column to identify batches (journal entries)
- If ANY transaction in a REF batch is IC, the ENTIRE batch is classified as IC
- Creates two separate CSV files:
  1. NetSuite_Non_IC.csv - All non-IC transactions
  2. NetSuite_IC.csv - All IC transactions (complete batches)
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def select_files():
    """Open a single GUI window to select both CSV and mapping files"""
    result = {'csv': None, 'mapping': None}
    
    def browse_csv():
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            csv_entry.delete(0, tk.END)
            csv_entry.insert(0, file_path)
            result['csv'] = file_path
    
    def browse_mapping():
        file_path = filedialog.askopenfilename(
            title="Select Mapping Excel File",
            filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
        )
        if file_path:
            mapping_entry.delete(0, tk.END)
            mapping_entry.insert(0, file_path)
            result['mapping'] = file_path
    
    def process():
        result['csv'] = csv_entry.get().strip()
        result['mapping'] = mapping_entry.get().strip()
        
        if not result['csv']:
            messagebox.showerror("Error", "Please select a CSV file")
            return
        if not result['mapping']:
            messagebox.showerror("Error", "Please select a mapping file")
            return
        
        root.destroy()
    
    def cancel():
        result['csv'] = None
        result['mapping'] = None
        root.destroy()
    
    # Create main window
    root = tk.Tk()
    root.title("NetSuite Import Processor")
    root.geometry("650x250")
    root.resizable(False, False)
    
    # Center the window
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 650) // 2
    y = (root.winfo_screenheight() - 250) // 2
    root.geometry(f"650x250+{x}+{y}")
    
    # Main frame with padding
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # Title
    title_label = ttk.Label(main_frame, text="NetSuite Import Processor", font=('Helvetica', 14, 'bold'))
    title_label.pack(pady=(0, 15))
    
    # CSV file row
    csv_frame = ttk.Frame(main_frame)
    csv_frame.pack(fill=tk.X, pady=5)
    ttk.Label(csv_frame, text="CSV File:", width=15).pack(side=tk.LEFT)
    csv_entry = ttk.Entry(csv_frame, width=50)
    csv_entry.pack(side=tk.LEFT, padx=5)
    ttk.Button(csv_frame, text="Browse...", command=browse_csv).pack(side=tk.LEFT)
    
    # Mapping file row
    mapping_frame = ttk.Frame(main_frame)
    mapping_frame.pack(fill=tk.X, pady=5)
    ttk.Label(mapping_frame, text="Mapping File:", width=15).pack(side=tk.LEFT)
    mapping_entry = ttk.Entry(mapping_frame, width=50)
    mapping_entry.pack(side=tk.LEFT, padx=5)
    ttk.Button(mapping_frame, text="Browse...", command=browse_mapping).pack(side=tk.LEFT)
    
    # Buttons row
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=20)
    ttk.Button(button_frame, text="Process", command=process, width=15).pack(side=tk.LEFT, padx=10)
    ttk.Button(button_frame, text="Cancel", command=cancel, width=15).pack(side=tk.LEFT, padx=10)
    
    root.mainloop()
    
    return result['csv'], result['mapping']


def load_mapping(mapping_file):
    """
    Load the mapping file and create dictionaries for:
    - Internal -> Intercompany status
    - LS Account -> Opposite Account (contra account)
    - LS Account -> Internal (for lookup)
    - Internal -> Opposite Internal (for duplicated rows)
    - Internal -> Account Name (for name lookup)
    - Internal -> LS Account (for account number lookup)
    """
    df_mapping = pd.read_excel(mapping_file)
    
    # Find the correct column names
    internal_col = None
    intercompany_col = None
    ls_account_col = None
    opposite_account_col = None
    opposite_internal_col = None
    account_name_col = None
    
    for col in df_mapping.columns:
        col_clean = str(col).strip().lower()
        # Match 'internal', 'ls internal', or any column containing 'internal' (but not 'opposite internal')
        if 'internal' in col_clean and 'opposite' not in col_clean and 'oposite' not in col_clean and internal_col is None:
            internal_col = col
        elif col_clean == 'intercompany' or 'intercompany' in col_clean:
            intercompany_col = col
        elif col_clean == 'ls account' or col_clean == 'lsaccount':
            ls_account_col = col
        elif ('oposite' in col_clean or 'opposite' in col_clean) and 'internal' in col_clean:
            # "Opposite Internal" column
            opposite_internal_col = col
        elif ('oposite' in col_clean or 'opposite' in col_clean) and 'account' in col_clean:
            # "Opposite Account" column
            opposite_account_col = col
        elif 'account name' in col_clean or col_clean == 'ls account name':
            account_name_col = col
    
    if internal_col is None:
        raise ValueError(f"Could not find 'Internal' column in mapping file. Available columns: {list(df_mapping.columns)}")
    if intercompany_col is None:
        raise ValueError(f"Could not find 'Intercompany' column in mapping file. Available columns: {list(df_mapping.columns)}")
    
    print(f"Using mapping columns:")
    print(f"  Internal: '{internal_col}'")
    print(f"  Intercompany: '{intercompany_col}'")
    if ls_account_col:
        print(f"  LS Account: '{ls_account_col}'")
    if opposite_account_col:
        print(f"  Opposite Account: '{opposite_account_col}'")
    if opposite_internal_col:
        print(f"  Opposite Internal: '{opposite_internal_col}'")
    if account_name_col:
        print(f"  Account Name: '{account_name_col}'")
    
    # Create mapping dictionaries
    ic_mapping = {}  # Internal -> is_intercompany (True/False)
    account_to_opposite = {}  # LS Account -> Opposite Account
    account_to_internal = {}  # LS Account -> Internal
    internal_to_opposite_account = {}  # Internal -> Opposite Account (direct mapping)
    internal_to_opposite_internal = {}  # Internal -> Opposite Internal (via opposite account)
    internal_to_account_name = {}  # Internal -> Account Name
    internal_to_ls_account = {}  # Internal -> LS Account
    
    for _, row in df_mapping.iterrows():
        internal_val = row[internal_col]
        intercompany_val = str(row[intercompany_col]).strip().lower()
        ic_mapping[internal_val] = intercompany_val == 'yes'
        
        # Build Internal -> Account Name mapping
        if account_name_col:
            account_name = row[account_name_col]
            if pd.notna(account_name):
                internal_to_account_name[internal_val] = account_name
                try:
                    internal_to_account_name[int(internal_val)] = account_name
                    internal_to_account_name[str(int(internal_val))] = account_name
                except (ValueError, TypeError):
                    pass
        
        # Build Internal -> LS Account mapping
        if ls_account_col:
            ls_account = row[ls_account_col]
            if pd.notna(ls_account):
                internal_to_ls_account[internal_val] = ls_account
                try:
                    internal_to_ls_account[int(internal_val)] = ls_account
                    internal_to_ls_account[str(int(internal_val))] = ls_account
                except (ValueError, TypeError):
                    pass
        
        # Build LS Account -> Internal mapping for ALL accounts
        if ls_account_col:
            ls_account = row[ls_account_col]
            if pd.notna(ls_account):
                account_to_internal[ls_account] = internal_val
                try:
                    int_acc = int(float(ls_account))
                    account_to_internal[int_acc] = internal_val
                    account_to_internal[str(int_acc)] = internal_val
                except (ValueError, TypeError):
                    pass
        
        # Build opposite account mapping if columns exist
        if ls_account_col and opposite_account_col:
            ls_account = row[ls_account_col]
            opposite_account = row[opposite_account_col]
            if pd.notna(ls_account) and pd.notna(opposite_account):
                # Store LS account -> Opposite account with multiple key types
                account_to_opposite[ls_account] = opposite_account
                try:
                    # Store as int and string for better matching
                    int_acc = int(float(ls_account))
                    account_to_opposite[int_acc] = opposite_account
                    account_to_opposite[str(int_acc)] = opposite_account
                except (ValueError, TypeError):
                    pass
                
                # Also map Internal -> Opposite Account directly
                internal_to_opposite_account[internal_val] = opposite_account
                try:
                    internal_to_opposite_account[int(internal_val)] = opposite_account
                    internal_to_opposite_account[str(int(internal_val))] = opposite_account
                except (ValueError, TypeError):
                    pass
    
    # Build Internal -> Opposite Internal mapping
    # Use the "Opposite Internal" column if it exists, otherwise compute from opposite account
    if opposite_internal_col:
        # Direct mapping from the "Opposite Internal" column
        print(f"  Using 'Opposite Internal' column for direct mapping")
        for _, row in df_mapping.iterrows():
            internal_val = row[internal_col]
            opposite_internal_val = row[opposite_internal_col]
            if pd.notna(opposite_internal_val):
                internal_to_opposite_internal[internal_val] = opposite_internal_val
                # Store with multiple key types for reliable lookup
                try:
                    int_internal = int(internal_val)
                    int_opp_internal = int(float(opposite_internal_val))
                    internal_to_opposite_internal[int_internal] = int_opp_internal
                    internal_to_opposite_internal[str(int_internal)] = int_opp_internal
                except (ValueError, TypeError):
                    pass
    elif ls_account_col and opposite_account_col:
        # Fallback: compute from opposite account lookup
        ls_to_internal = {}
        opposite_to_internal = {}
        for _, row in df_mapping.iterrows():
            ls_account = row[ls_account_col]
            internal_val = row[internal_col]
            opposite_account = row[opposite_account_col]
            if pd.notna(ls_account):
                ls_to_internal[ls_account] = internal_val
            if pd.notna(opposite_account):
                opposite_to_internal[opposite_account] = internal_val
        
        # Now build internal -> opposite internal
        for _, row in df_mapping.iterrows():
            internal_val = row[internal_col]
            opposite_account = row[opposite_account_col]
            if pd.notna(opposite_account):
                # Find the internal for the opposite account
                if opposite_account in ls_to_internal:
                    opp_internal = ls_to_internal[opposite_account]
                    internal_to_opposite_internal[internal_val] = opp_internal
                    # Store with multiple key types
                    try:
                        internal_to_opposite_internal[int(internal_val)] = opp_internal
                        internal_to_opposite_internal[str(int(internal_val))] = opp_internal
                    except (ValueError, TypeError):
                        pass
    
    print(f"\nLoaded {len(ic_mapping)} mappings from Excel file")
    print(f"I/C accounts: {sum(1 for v in ic_mapping.values() if v)}")
    print(f"Non-I/C accounts: {sum(1 for v in ic_mapping.values() if not v)}")
    if account_to_opposite:
        print(f"Opposite account mappings: {len(account_to_opposite)}")
    if internal_to_opposite_internal:
        print(f"Internal to Opposite Internal mappings: {len(internal_to_opposite_internal)}")
    if internal_to_account_name:
        print(f"Internal to Account Name mappings: {len(internal_to_account_name)}")
    if internal_to_ls_account:
        print(f"Internal to LS Account mappings: {len(internal_to_ls_account)}")
    
    # Debug: print some mappings
    print(f"\nSample account_to_opposite mappings (LS account -> Opposite account):")
    for i, (k, v) in enumerate(account_to_opposite.items()):
        if i < 10:
            print(f"  {k} -> {v}")
    
    # Debug: Check if specific accounts exist in mappings
    print(f"\nVerifying key account mappings:")
    for test_account in [11011, '11011', 13401, '13401']:
        if test_account in account_to_internal:
            print(f"  account_to_internal[{test_account}] = {account_to_internal[test_account]}")
    
    # Debug: Show sample internal_to_opposite_internal mappings
    print(f"\nSample internal_to_opposite_internal mappings (Internal -> Opposite Internal):")
    shown = 0
    for internal_key, opp_internal in internal_to_opposite_internal.items():
        if shown < 10 and isinstance(internal_key, int):  # Only show int keys to avoid duplicates
            print(f"  {internal_key} -> {opp_internal}")
            shown += 1
    
    return {
        'ic_mapping': ic_mapping,
        'account_to_opposite': account_to_opposite,
        'account_to_internal': account_to_internal,
        'internal_to_opposite_account': internal_to_opposite_account,
        'internal_to_opposite_internal': internal_to_opposite_internal,
        'internal_to_account_name': internal_to_account_name,
        'internal_to_ls_account': internal_to_ls_account
    }


def read_csv_with_encoding(csv_file):
    """
    Attempt to read CSV file with multiple encodings.
    Tries common encodings: utf-8, windows-1252, iso-8859-1, cp1252
    """
    encodings = ['utf-8', 'windows-1252', 'iso-8859-1', 'cp1252', 'latin1']
    
    for encoding in encodings:
        try:
            df = pd.read_csv(csv_file, encoding=encoding)
            print(f"Successfully read CSV file with encoding: {encoding}")
            return df
        except UnicodeDecodeError:
            continue
        except Exception as e:
            # For other errors, try next encoding
            continue
    
    # If all encodings fail, try with error handling
    try:
        df = pd.read_csv(csv_file, encoding='utf-8', errors='replace')
        print("Warning: Read CSV with UTF-8 encoding using error replacement (some characters may be corrupted)")
        return df
    except Exception as e:
        raise Exception(f"Failed to read CSV file with any encoding. Last error: {str(e)}")


def process_csv(csv_file, mapping_data):
    """
    Process the CSV file:
    1. Identify IC transactions based on mapping
    2. Use REF column to identify batches - if ANY line in a batch is IC, entire batch is IC
    3. Create two separate DataFrames: Non-IC and IC transactions
    4. For IC file: add/modify columns and duplicate with opposite accounts
    """
    # Extract mapping dictionaries
    ic_mapping = mapping_data['ic_mapping']
    account_to_opposite = mapping_data['account_to_opposite']
    account_to_internal = mapping_data['account_to_internal']
    internal_to_opposite_account = mapping_data['internal_to_opposite_account']
    internal_to_opposite_internal = mapping_data['internal_to_opposite_internal']
    internal_to_account_name = mapping_data['internal_to_account_name']
    internal_to_ls_account = mapping_data['internal_to_ls_account']
    
    # Read CSV with encoding detection
    df = read_csv_with_encoding(csv_file)
    
    # Find the required columns
    internal_col = None
    ref_col = None
    class_col = None
    location_col = None
    date_col = None
    subsidiary_col = None
    account_col = None
    account_name_col = None
    
    for col in df.columns:
        col_clean = str(col).strip().lower()
        # Match 'internal', 'ls internal', or any column containing 'internal'
        if 'internal' in col_clean and internal_col is None:
            internal_col = col
        elif col_clean == 'ref':
            ref_col = col
        elif col_clean == 'class':
            class_col = col
        elif 'location' in col_clean and 'expense' in col_clean:
            location_col = col
        elif col_clean == 'date':
            date_col = col
        elif col_clean == 'subsidiary':
            subsidiary_col = col
        elif col_clean == 'account' or col_clean == 'ls account':
            account_col = col
        elif 'account name' in col_clean or col_clean == 'ls account name':
            account_name_col = col
    
    if internal_col is None:
        raise ValueError(f"Could not find 'Internal' column in CSV file. Available columns: {list(df.columns)}")
    if ref_col is None:
        raise ValueError(f"Could not find 'REF' column in CSV file. Available columns: {list(df.columns)}")
    
    print(f"Processing {len(df)} rows from CSV file")
    print(f"Internal column: '{internal_col}'")
    print(f"REF column: '{ref_col}'")
    
    # === Validate internal IDs against mapping ===
    # Get all unique internal IDs from CSV
    csv_internals = set()
    for val in df[internal_col].dropna().unique():
        try:
            csv_internals.add(int(val))
        except (ValueError, TypeError):
            csv_internals.add(val)
    
    # Get all internal IDs from mapping
    mapping_internals = set()
    for key in ic_mapping.keys():
        try:
            mapping_internals.add(int(key))
        except (ValueError, TypeError):
            mapping_internals.add(key)
    
    # Find unmapped internals
    unmapped_internals = csv_internals - mapping_internals
    
    if unmapped_internals:
        sorted_unmapped = sorted([x for x in unmapped_internals if isinstance(x, int)])
        print(f"\n  ⚠️  WARNING: Found {len(unmapped_internals)} internal IDs in CSV that are NOT in the mapping file:")
        for internal_id in sorted_unmapped[:20]:  # Show first 20
            print(f"      - Internal {internal_id}")
        if len(sorted_unmapped) > 20:
            print(f"      ... and {len(sorted_unmapped) - 20} more")
        
        # Show dialog asking user what to do
        unmapped_list = "\n".join([f"  • {x}" for x in sorted_unmapped[:15]])
        if len(sorted_unmapped) > 15:
            unmapped_list += f"\n  ... and {len(sorted_unmapped) - 15} more"
        
        root = tk.Tk()
        root.withdraw()
        user_choice = messagebox.askyesnocancel(
            "Unmapped Internal IDs Found",
            f"Found {len(unmapped_internals)} internal IDs in the CSV that are NOT in the mapping file:\n\n"
            f"{unmapped_list}\n\n"
            f"Do you want to FILTER OUT these rows?\n\n"
            f"• Yes = Remove rows with unmapped internals\n"
            f"• No = Keep all rows (may have incorrect mappings)\n"
            f"• Cancel = Stop processing"
        )
        root.destroy()
        
        if user_choice is None:
            # User clicked Cancel
            raise ValueError("Processing cancelled by user due to unmapped internal IDs")
        elif user_choice:
            # User clicked Yes - filter out unmapped rows
            rows_before = len(df)
            
            def is_mapped(internal_val):
                try:
                    return int(internal_val) in mapping_internals
                except (ValueError, TypeError):
                    return internal_val in mapping_internals
            
            df = df[df[internal_col].apply(is_mapped)]
            rows_removed = rows_before - len(df)
            print(f"  Filtered out {rows_removed} rows with unmapped internal IDs")
        else:
            # User clicked No - keep all rows
            print(f"  Keeping all rows (including unmapped internals)")
    else:
        print(f"  ✓ All internal IDs in CSV are mapped in the mapping file")
    
    if class_col:
        print(f"Class column: '{class_col}'")
    if location_col:
        print(f"Location column: '{location_col}'")
    if date_col:
        print(f"Date column: '{date_col}'")
    if account_col:
        print(f"Account column: '{account_col}'")
    if account_name_col:
        print(f"Account Name column: '{account_name_col}'")
    
    # Determine I/C status for each row
    def is_intercompany(internal_val):
        # Try to match with mapping dictionary
        if internal_val in ic_mapping:
            return ic_mapping[internal_val]
        # Try converting to int/string variations
        try:
            int_val = int(internal_val)
            if int_val in ic_mapping:
                return ic_mapping[int_val]
            if str(int_val) in ic_mapping:
                return ic_mapping[str(int_val)]
        except (ValueError, TypeError):
            pass
        return False  # Default to non-I/C if not found in mapping
    
    # Mark each row as IC or not based on individual account mapping
    df['_is_ic_row'] = df[internal_col].apply(is_intercompany)
    
    # Find all REF values that contain at least one IC transaction
    # These are the "IC batches" - entire batch should be treated as IC
    ic_refs = df[df['_is_ic_row']][ref_col].unique()
    print(f"Found {len(ic_refs)} unique REF batches containing IC transactions")
    
    # Mark all rows belonging to IC batches
    df['_is_ic_batch'] = df[ref_col].isin(ic_refs)
    
    # Split into IC and Non-IC DataFrames based on batch
    df_ic = df[df['_is_ic_batch']].copy()
    df_non_ic = df[~df['_is_ic_batch']].copy()
    
    # Store original internal values before any modifications
    df_ic['_original_internal'] = df_ic[internal_col]
    if account_col:
        df_ic['_original_account'] = df_ic[account_col]
    
    # === Update Account and Account Name based on Internal ID ===
    # This ensures the account and account name are correctly mapped from the internal ID
    if account_col and internal_to_ls_account:
        def get_account_from_internal(internal_val):
            if internal_val in internal_to_ls_account:
                return internal_to_ls_account[internal_val]
            try:
                int_val = int(internal_val)
                if int_val in internal_to_ls_account:
                    return internal_to_ls_account[int_val]
                if str(int_val) in internal_to_ls_account:
                    return internal_to_ls_account[str(int_val)]
            except (ValueError, TypeError):
                pass
            return None  # Return None if not found
        
        # Update account column based on internal ID
        def update_account(row):
            mapped_account = get_account_from_internal(row[internal_col])
            if mapped_account is not None:
                return mapped_account
            return row[account_col]  # Keep original if not found
        
        df_ic[account_col] = df_ic.apply(update_account, axis=1)
        print(f"  Updated Account column based on Internal ID mapping")
    
    if account_name_col and internal_to_account_name:
        def get_account_name_from_internal(internal_val):
            if internal_val in internal_to_account_name:
                return internal_to_account_name[internal_val]
            try:
                int_val = int(internal_val)
                if int_val in internal_to_account_name:
                    return internal_to_account_name[int_val]
                if str(int_val) in internal_to_account_name:
                    return internal_to_account_name[str(int_val)]
            except (ValueError, TypeError):
                pass
            return None  # Return None if not found
        
        # Update account name column based on internal ID
        def update_account_name(row):
            mapped_name = get_account_name_from_internal(row[internal_col])
            if mapped_name is not None:
                return mapped_name
            return row[account_name_col]  # Keep original if not found
        
        df_ic[account_name_col] = df_ic.apply(update_account_name, axis=1)
        print(f"  Updated Account Name column based on Internal ID mapping")
    
    # === Modify IC DataFrame with additional columns ===
    
    # 1. Rename "Subsidiary" to "Subsidiary Header" and add subsidiary-related columns
    if subsidiary_col and subsidiary_col in df_ic.columns:
        # Rename the column
        df_ic = df_ic.rename(columns={subsidiary_col: 'Subsidiary Header'})
        
        # Set Subsidiary Header to 6 (STATSCORE) since the JE is recorded in ST books
        df_ic['Subsidiary Header'] = 6
        
        # Insert "Subsidiary Line" column right after "Subsidiary Header" with value 6 (opposite of 3)
        cols = list(df_ic.columns)
        sub_idx = cols.index('Subsidiary Header')
        df_ic.insert(sub_idx + 1, 'Subsidiary Line', 6)
        
        # Insert "Subsidiary Line/ DUE TO/ FROM SUBSIDIARY" column right after "Subsidiary Line" with value 3
        cols = list(df_ic.columns)
        sub_line_idx = cols.index('Subsidiary Line')
        df_ic.insert(sub_line_idx + 1, 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY', 3)
    
    # 2. Clear CLASS column (leave empty)
    if class_col and class_col in df_ic.columns:
        df_ic[class_col] = ''
    
    # 3. Set Location (Expenses) (Req) to "Other"
    if location_col and location_col in df_ic.columns:
        df_ic[location_col] = 'Other'
    
    # 4. Add "Name" column with 11064 for LSPORTS
    df_ic['Name'] = 11064
    
    # 5. Add "Posting Period" column based on Date (format: Mmm YYYY)
    if date_col and date_col in df_ic.columns:
        # Use explicit English month names to avoid locale issues (Hebrew system)
        MONTH_NAMES = {
            1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
            7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
        }
        
        def format_posting_period(date_val):
            try:
                # Parse the date
                parsed_date = pd.to_datetime(date_val, dayfirst=True)
                # Get month and year as integers
                month_num = int(parsed_date.month)
                year_num = int(parsed_date.year)
                # Build string manually with explicit English month
                month_str = MONTH_NAMES.get(month_num, 'Jan')
                return str(month_str) + ' ' + str(year_num)  # e.g., "Jan 2025"
            except:
                return ''
        df_ic['Posting Period'] = df_ic[date_col].apply(format_posting_period)
    else:
        df_ic['Posting Period'] = ''
    
    # 6. Add "Eliminate" column based on IC status of each row
    df_ic['Eliminate'] = df_ic['_is_ic_row'].apply(lambda x: 'Yes' if x else 'No')
    
    # 7. Add "Department" column with 24 for all
    df_ic['Department'] = 24
    
    # 8. Add "EXTERNAL ID" column with "JE_ICMMYY" where MMYY is month and year from newest date
    # E.g., October 2025 = "JE_IC1025", January 2025 = "JE_IC0125"
    # This column should be at position 0 (column A)
    if date_col and date_col in df_ic.columns:
        # Find the newest date in the Date column
        try:
            dates = pd.to_datetime(df_ic[date_col], dayfirst=True, errors='coerce')
            newest_date = dates.max()
            if pd.notna(newest_date):
                month_num = int(newest_date.month)
                year_short = int(newest_date.year) % 100  # Get last 2 digits of year
                external_id = f'JE_IC{month_num:02d}{year_short:02d}'
                df_ic.insert(0, 'EXTERNAL ID', external_id)
                print(f"  EXTERNAL ID set to: {external_id} (from newest date: {newest_date.strftime('%Y-%m-%d')})")
            else:
                df_ic.insert(0, 'EXTERNAL ID', '')
        except Exception as e:
            print(f"  Warning: Could not determine month for EXTERNAL ID: {e}")
            df_ic.insert(0, 'EXTERNAL ID', '')
    else:
        df_ic.insert(0, 'EXTERNAL ID', '')
    
    # === Special rule: Apply 400001 -> 400020 switch BEFORE creating duplicates ===
    # If a REF batch contains account 120002 (Accounts Receivable),
    # then switch account 400001 to 400020 in that batch
    if account_col and ref_col:
        def has_account_120002(acc_val):
            try:
                return int(float(acc_val)) == 120002
            except:
                return str(acc_val).strip() == '120002'
        
        refs_with_120002 = df_ic[df_ic[account_col].apply(has_account_120002)][ref_col].unique()
        
        if len(refs_with_120002) > 0:
            print(f"  Found {len(refs_with_120002)} REF batches with account 120002")
            
            def switch_account_orig(row):
                ref_val = row[ref_col]
                acc_val = row[account_col]
                if ref_val in refs_with_120002:
                    try:
                        if int(float(acc_val)) == 400001:
                            return 400020
                    except:
                        if str(acc_val).strip() == '400001':
                            return 400020
                return acc_val
            
            def switch_internal_orig(row):
                ref_val = row[ref_col]
                acc_val = row[account_col]
                internal_val = row[internal_col]
                if ref_val in refs_with_120002:
                    try:
                        if int(float(acc_val)) == 400001:
                            if 400020 in account_to_internal:
                                return account_to_internal[400020]
                            elif '400020' in account_to_internal:
                                return account_to_internal['400020']
                    except:
                        if str(acc_val).strip() == '400001':
                            if 400020 in account_to_internal:
                                return account_to_internal[400020]
                            elif '400020' in account_to_internal:
                                return account_to_internal['400020']
                return internal_val
            
            # Apply switches on original IC rows
            rows_before = len(df_ic[df_ic[account_col].apply(lambda x: str(x).strip() == '400001' or (isinstance(x, (int, float)) and int(x) == 400001))])
            df_ic[internal_col] = df_ic.apply(switch_internal_orig, axis=1)
            df_ic[account_col] = df_ic.apply(switch_account_orig, axis=1)
            rows_after = len(df_ic[df_ic[account_col].apply(lambda x: str(x).strip() == '400001' or (isinstance(x, (int, float)) and int(x) == 400001))])
            switched = rows_before - rows_after
            if switched > 0:
                print(f"  Switched {switched} rows from account 400001 to 400020 in original IC rows")
                
                # Update Account Name for switched rows based on new internal ID
                if account_name_col and internal_to_account_name:
                    def update_switched_account_name(row):
                        internal_val = row[internal_col]
                        if internal_val in internal_to_account_name:
                            return internal_to_account_name[internal_val]
                        try:
                            int_val = int(internal_val)
                            if int_val in internal_to_account_name:
                                return internal_to_account_name[int_val]
                            if str(int_val) in internal_to_account_name:
                                return internal_to_account_name[str(int_val)]
                        except (ValueError, TypeError):
                            pass
                        return row[account_name_col]
                    
                    df_ic[account_name_col] = df_ic.apply(update_switched_account_name, axis=1)
    
    # === CONFIGURATION: Show Name=11065 or leave empty ===
    # Set to True to show 11065 in Name column, False to leave Name empty for opposite rows
    SHOW_NAME_11065 = False  # Change to True to show 11065 in Name column
    
    # === Create duplicated rows with opposite accounts ===
    # Original table: uses LS internal (col 1) and LS account (col 2) - already applied above
    # Opposite table: uses Opposite internal (col 9) and Opposite account (col 8) from mapping
    df_ic_duplicated = df_ic.copy()
    
    # In duplicated rows:
    # - "Subsidiary Line" = 5 (xElimination - Parent)
    # - "Subsidiary Line/ DUE TO/ FROM SUBSIDIARY" = 6
    # - "Name" = 11065 for STATSCORE (or empty if SHOW_NAME_11065 is False)
    # - Account = Opposite account from mapping column 8
    # - Internal = Opposite internal from mapping column 9
    
    if 'Subsidiary Line' in df_ic_duplicated.columns:
        df_ic_duplicated['Subsidiary Line'] = 5  # xElimination - Parent
    
    if 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY' in df_ic_duplicated.columns:
        df_ic_duplicated['Subsidiary Line/ DUE TO/ FROM SUBSIDIARY'] = 6
    
    # Set Name column - 11065 or empty based on configuration
    if SHOW_NAME_11065:
        df_ic_duplicated['Name'] = 11065
    else:
        df_ic_duplicated['Name'] = ''  # Empty cell instead of 11065
    
    # Replace account with opposite account from "Opposite account" column (column 8)
    if account_col:
        def get_opposite_account(row):
            # Use the ORIGINAL internal value to look up the opposite account
            original_internal = row['_original_internal']
            opposite_account = None
            
            # Look up using internal_to_opposite_account (maps LS internal -> Opposite account)
            if original_internal in internal_to_opposite_account:
                opposite_account = internal_to_opposite_account[original_internal]
            else:
                try:
                    int_internal = int(original_internal)
                    if int_internal in internal_to_opposite_account:
                        opposite_account = internal_to_opposite_account[int_internal]
                    elif str(int_internal) in internal_to_opposite_account:
                        opposite_account = internal_to_opposite_account[str(int_internal)]
                except (ValueError, TypeError):
                    pass
            
            # Normalize the opposite account to int if possible
            if opposite_account is not None:
                try:
                    return int(float(opposite_account))
                except (ValueError, TypeError):
                    return opposite_account
            
            # Return original account if no mapping found
            return row[account_col]
        
        df_ic_duplicated[account_col] = df_ic_duplicated.apply(get_opposite_account, axis=1)
    
    # Replace internal with the opposite internal from "Opposite Internal" column (column 9)
    def get_opposite_internal(row):
        # Use the ORIGINAL internal value to look up the opposite internal
        original_internal = row['_original_internal']
        
        # Try direct lookup in internal_to_opposite_internal
        if original_internal in internal_to_opposite_internal:
            return internal_to_opposite_internal[original_internal]
        
        # Try with type conversions
        try:
            int_internal = int(original_internal)
            if int_internal in internal_to_opposite_internal:
                return internal_to_opposite_internal[int_internal]
            if str(int_internal) in internal_to_opposite_internal:
                return internal_to_opposite_internal[str(int_internal)]
        except (ValueError, TypeError):
            pass
        
        # Return original internal if no mapping found
        return original_internal
    
    df_ic_duplicated[internal_col] = df_ic_duplicated.apply(get_opposite_internal, axis=1)
    
    # Update Account Name for duplicated rows based on the new internal ID
    if account_name_col and internal_to_account_name:
        def get_account_name_for_dup(row):
            internal_val = row[internal_col]
            if internal_val in internal_to_account_name:
                return internal_to_account_name[internal_val]
            try:
                int_val = int(internal_val)
                if int_val in internal_to_account_name:
                    return internal_to_account_name[int_val]
                if str(int_val) in internal_to_account_name:
                    return internal_to_account_name[str(int_val)]
            except (ValueError, TypeError):
                pass
            return row[account_name_col]  # Keep original if not found
        
        df_ic_duplicated[account_name_col] = df_ic_duplicated.apply(get_account_name_for_dup, axis=1)
    
    # Concatenate original IC rows with duplicated rows (opposite rows always included, Name may be empty)
    df_ic_final = pd.concat([df_ic, df_ic_duplicated], ignore_index=True)
    if SHOW_NAME_11065:
        print(f"  Opposite rows: Name=11065 ({len(df_ic_duplicated)} rows)")
    else:
        print(f"  Opposite rows: Name=empty ({len(df_ic_duplicated)} rows) - set SHOW_NAME_11065=True to show 11065")
    
    # Remove helper columns
    helper_cols = ['_is_ic_row', '_is_ic_batch', '_original_internal']
    if '_original_account' in df_ic_final.columns:
        helper_cols.append('_original_account')
    df_ic_final = df_ic_final.drop(columns=helper_cols)
    df_non_ic = df_non_ic.drop(columns=['_is_ic_row', '_is_ic_batch'])
    
    # Find debit and credit columns
    debit_col = None
    credit_col = None
    for col in df_ic_final.columns:
        col_clean = str(col).strip().lower()
        if 'debit' in col_clean:
            debit_col = col
        elif 'credit' in col_clean:
            credit_col = col
    
    # Filter out matching debit/credit pairs within same REF batch and same account
    # If same REF has rows with same account where debit equals credit, remove both rows
    if debit_col and credit_col and ref_col and account_col:
        rows_before = len(df_ic_final)
        
        # Convert to numeric
        df_ic_final['_debit_num'] = pd.to_numeric(df_ic_final[debit_col], errors='coerce').fillna(0)
        df_ic_final['_credit_num'] = pd.to_numeric(df_ic_final[credit_col], errors='coerce').fillna(0)
        
        # Create list of indices to remove
        indices_to_remove = set()
        
        # Group by REF and Account
        for (ref_val, acc_val), group in df_ic_final.groupby([ref_col, account_col]):
            if len(group) < 2:
                continue
            
            # Find debit rows (debit > 0, credit = 0)
            debit_rows = group[(group['_debit_num'] > 0) & (group['_credit_num'] == 0)]
            # Find credit rows (credit > 0, debit = 0)
            credit_rows = group[(group['_credit_num'] > 0) & (group['_debit_num'] == 0)]
            
            # For each debit row, check if there's a matching credit row
            for d_idx, d_row in debit_rows.iterrows():
                debit_amount = d_row['_debit_num']
                # Find matching credit rows with same amount
                matching_credits = credit_rows[credit_rows['_credit_num'] == debit_amount]
                if len(matching_credits) > 0:
                    # Mark both for removal
                    indices_to_remove.add(d_idx)
                    # Remove first matching credit
                    c_idx = matching_credits.index[0]
                    indices_to_remove.add(c_idx)
                    # Update credit_rows to exclude this matched one
                    credit_rows = credit_rows.drop(c_idx)
        
        # Remove marked rows
        if indices_to_remove:
            df_ic_final = df_ic_final.drop(index=list(indices_to_remove))
            print(f"  Filtered out {len(indices_to_remove)} rows with matching debit/credit in same REF and account")
        
        # Remove helper columns
        df_ic_final = df_ic_final.drop(columns=['_debit_num', '_credit_num'])
    
    # Filter out rows where both debit and credit are zero
    if debit_col and credit_col:
        rows_before = len(df_ic_final)
        # Convert to numeric, treating empty/NaN as 0
        debit_vals = pd.to_numeric(df_ic_final[debit_col], errors='coerce').fillna(0)
        credit_vals = pd.to_numeric(df_ic_final[credit_col], errors='coerce').fillna(0)
        # Keep rows where at least one of debit or credit is non-zero
        df_ic_final = df_ic_final[(debit_vals != 0) | (credit_vals != 0)]
        rows_filtered = rows_before - len(df_ic_final)
        if rows_filtered > 0:
            print(f"  Filtered out {rows_filtered} rows with zero debit and credit")
    
    # Sort by REF column
    if ref_col and ref_col in df_ic_final.columns:
        df_ic_final = df_ic_final.sort_values(by=ref_col).reset_index(drop=True)
        print(f"  Sorted by {ref_col} column")
    
    print(f"\nResults:")
    print(f"  Non-IC transactions: {len(df_non_ic)} rows")
    print(f"  IC transactions (with duplicates): {len(df_ic_final)} rows")
    print(f"  IC batches (unique REFs): {len(ic_refs)}")
    
    return df_non_ic, df_ic_final


def save_csv(df, output_path):
    """Save DataFrame to CSV file with UTF-8 encoding"""
    # Save with UTF-8 encoding with BOM for better compatibility
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved {len(df)} rows to {output_path}")


def get_month_filename(posting_period, prefix="CSV_Statscore_intercompany"):
    """
    Convert posting period to filename format.
    E.g., "Jan 2025" -> "CSV_Statscore_intercompany_1.25", "Feb 2025" -> "CSV_Statscore_intercompany_2.25"
    """
    MONTH_TO_NUM = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    try:
        # Parse the posting period (e.g., "Jan 2025")
        parts = str(posting_period).strip().split()
        if len(parts) >= 2:
            month_str = parts[0].lower()[:3]
            year_str = parts[-1]
            
            month_num = MONTH_TO_NUM.get(month_str)
            if month_num:
                # Get last 2 digits of year
                year_short = year_str[-2:]
                return f"{prefix}_{month_num}.{year_short}"
    except:
        pass
    
    return None


def get_statscore_import_folder():
    """
    Create and return the path to "Statscore Import" folder on the desktop.
    """
    # Get the desktop path
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    
    # Create "Statscore Import" folder
    statscore_folder = os.path.join(desktop_path, "Statscore Import")
    
    if not os.path.exists(statscore_folder):
        os.makedirs(statscore_folder)
        print(f"  Created folder: {statscore_folder}")
    
    return statscore_folder


def add_external_id_column(df, date_col, is_ic=True):
    """
    Add EXTERNAL ID column to DataFrame based on date column.
    Format for IC: JE_ICMMYY (e.g., JE_IC0125 for January 2025)
    Format for Non-IC: JE_nonICMMYY (e.g., JE_nonIC0125 for January 2025)
    """
    prefix = 'JE_IC' if is_ic else 'JE_nonIC'
    
    if date_col and date_col in df.columns:
        try:
            dates = pd.to_datetime(df[date_col], dayfirst=True, errors='coerce')
            newest_date = dates.max()
            if pd.notna(newest_date):
                month_num = int(newest_date.month)
                year_short = int(newest_date.year) % 100
                external_id = f'{prefix}{month_num:02d}{year_short:02d}'
                df.insert(0, 'EXTERNAL ID', external_id)
                return external_id
        except:
            pass
    
    df.insert(0, 'EXTERNAL ID', '')
    return ''


def save_by_month(df, output_dir, file_prefix, is_ic=True):
    """
    Split DataFrame by Posting Period and save each month to a separate file.
    Files are saved to "Statscore Import" folder on the desktop.
    Returns list of saved file paths.
    
    Args:
        df: DataFrame to split and save
        output_dir: Original output directory (not used, kept for compatibility)
        file_prefix: Prefix for filenames (e.g., "CSV_Statscore_intercompany" or "NetSuite_Non_IC")
        is_ic: Whether this is IC data (IC files already have EXTERNAL ID)
    """
    saved_files = []
    
    # Get the Statscore Import folder on desktop
    statscore_folder = get_statscore_import_folder()
    
    # Find the Posting Period column
    posting_period_col = None
    for col in df.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    # Find the Date column for Non-IC files (to add EXTERNAL ID)
    date_col = None
    if not is_ic:
        for col in df.columns:
            if str(col).strip().lower() == 'date':
                date_col = col
                break
    
    if posting_period_col is None:
        print(f"  Warning: Could not find 'Posting Period' column. Saving all to single file.")
        output_path = os.path.join(statscore_folder, f"{file_prefix}.csv")
        
        # Add EXTERNAL ID for Non-IC files
        if not is_ic and 'EXTERNAL ID' not in df.columns:
            add_external_id_column(df, date_col, is_ic=False)
        
        save_csv(df, output_path)
        return [output_path], statscore_folder
    
    # Group by Posting Period
    grouped = df.groupby(posting_period_col)
    
    print(f"\n  Splitting {file_prefix} by Posting Period...")
    print(f"  Saving to: {statscore_folder}")
    
    for posting_period, group_df in grouped:
        group_df = group_df.copy().reset_index(drop=True)
        
        # Add EXTERNAL ID for Non-IC files
        if not is_ic and 'EXTERNAL ID' not in group_df.columns:
            add_external_id_column(group_df, date_col, is_ic=False)
        
        # Get the filename for this month
        filename = get_month_filename(posting_period, prefix=file_prefix)
        
        if filename is None:
            # Fallback: use the posting period as filename (sanitized)
            safe_period = str(posting_period).replace('/', '-').replace('\\', '-').replace(' ', '_')
            filename = f"{file_prefix}_{safe_period}"
        
        output_path = os.path.join(statscore_folder, f"{filename}.csv")
        save_csv(group_df, output_path)
        saved_files.append((posting_period, output_path, len(group_df)))
    
    return saved_files, statscore_folder


def main():
    print("NetSuite Import Processor")
    print("=" * 50)
    
    # Select files via GUI
    csv_file, mapping_file = select_files()
    
    if not csv_file or not mapping_file:
        print("File selection cancelled. Exiting.")
        return
    
    print(f"CSV file: {csv_file}")
    print(f"Mapping file: {mapping_file}")
    
    try:
        # Load mapping
        print("\nLoading mapping file...")
        mapping_dict = load_mapping(mapping_file)
        
        # Process CSV - returns two DataFrames (Non-IC and IC)
        print("\nProcessing CSV file...")
        df_non_ic, df_ic = process_csv(csv_file, mapping_dict)
        
        # Generate output directory
        output_dir = os.path.dirname(csv_file)
        
        # Add Posting Period column to Non-IC data for splitting
        date_col = None
        for col in df_non_ic.columns:
            if str(col).strip().lower() == 'date':
                date_col = col
                break
        
        if date_col and 'Posting Period' not in df_non_ic.columns:
            MONTH_NAMES = {
                1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
                7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
            }
            
            def format_posting_period(date_val):
                try:
                    parsed_date = pd.to_datetime(date_val, dayfirst=True)
                    month_num = int(parsed_date.month)
                    year_num = int(parsed_date.year)
                    month_str = MONTH_NAMES.get(month_num, 'Jan')
                    return str(month_str) + ' ' + str(year_num)
                except:
                    return ''
            
            df_non_ic['Posting Period'] = df_non_ic[date_col].apply(format_posting_period)
        
        # Save results
        print("\nSaving results...")
        
        # Save Non-IC transactions split by month (to Statscore Import folder on desktop)
        non_ic_files, statscore_folder = save_by_month(df_non_ic, output_dir, "NetSuite_Non_IC", is_ic=False)
        
        # Save IC transactions split by month (to Statscore Import folder on desktop)
        ic_files, _ = save_by_month(df_ic, output_dir, "CSV_Statscore_intercompany", is_ic=True)
        
        # Build success message for Non-IC files
        non_ic_files_msg = ""
        if isinstance(non_ic_files[0], tuple):
            for posting_period, path, count in non_ic_files:
                filename = os.path.basename(path)
                non_ic_files_msg += f"  {filename} ({count} rows)\n"
        else:
            non_ic_files_msg = f"  {os.path.basename(non_ic_files[0])} ({len(df_non_ic)} rows)\n"
        
        # Build success message for IC files
        ic_files_msg = ""
        if isinstance(ic_files[0], tuple):
            for posting_period, path, count in ic_files:
                filename = os.path.basename(path)
                ic_files_msg += f"  {filename} ({count} rows)\n"
        else:
            ic_files_msg = f"  {os.path.basename(ic_files[0])} ({len(df_ic)} rows)\n"
        
        # Show success message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", 
            f"Processing complete!\n\n"
            f"All files saved to:\n{statscore_folder}\n\n"
            f"Non-IC transactions ({len(df_non_ic)} rows):\n{non_ic_files_msg}\n"
            f"IC transactions ({len(df_ic)} rows):\n{ic_files_msg}")
        root.destroy()
        
        print("\nDone!")
        
    except Exception as e:
        print(f"\nError: {e}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        root.destroy()


if __name__ == "__main__":
    main()

