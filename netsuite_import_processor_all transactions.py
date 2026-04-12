"""
NetSuite Import Processor
- Opens a GUI to select CSV file and Excel mapping file
- Identifies Intercompany (IC) transactions using the mapping file
- Uses REF column to identify batches (journal entries)
- If ANY transaction in a REF batch is IC, the ENTIRE batch is classified as IC
- Creates combined CSV files per month with same IC structure for all transactions
- Creates IC pivot table summary for elimination journal entry
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re

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
    default_mapping = r"\\mainsrv\d\Data\ביקורת\Lsports Data LTD\NETSUITE\יבוא מ-STATSCORE\ST mapping.xlsx"
    if os.path.exists(default_mapping):
        mapping_entry.insert(0, default_mapping)
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
    account_type_col = None
    
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
        elif 'account type' in col_clean or col_clean == 'ls account type':
            account_type_col = col
    
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
    if account_type_col:
        print(f"  Account Type: '{account_type_col}'")
    
    # Create mapping dictionaries
    ic_mapping = {}  # Internal -> is_intercompany (True/False)
    account_to_opposite = {}  # LS Account -> Opposite Account
    account_to_internal = {}  # LS Account -> Internal
    internal_to_opposite_account = {}  # Internal -> Opposite Account (direct mapping)
    internal_to_opposite_internal = {}  # Internal -> Opposite Internal (via opposite account)
    internal_to_account_name = {}  # Internal -> Account Name
    internal_to_ls_account = {}  # Internal -> LS Account
    internal_to_account_type = {}  # Internal -> Account Type (for P&L filtering)
    
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
        
        # Build Internal -> Account Type mapping (for P&L filtering)
        if account_type_col:
            account_type = row[account_type_col]
            if pd.notna(account_type):
                internal_to_account_type[internal_val] = str(account_type).strip()
                try:
                    internal_to_account_type[int(internal_val)] = str(account_type).strip()
                    internal_to_account_type[str(int(internal_val))] = str(account_type).strip()
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
    if internal_to_account_type:
        print(f"Internal to Account Type mappings: {len(internal_to_account_type)}")
    
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
    
    # === Load "LS COA" sheet for elimination account names ===
    # Column D = Internal ID, Column F = Account Name
    ls_coa_internal_to_name = {}
    try:
        df_ls_coa = pd.read_excel(mapping_file, sheet_name='LS COA')
        
        # Get column D (index 3) and column F (index 5)
        if len(df_ls_coa.columns) >= 6:
            internal_col_coa = df_ls_coa.columns[3]  # Column D (0-indexed = 3)
            name_col_coa = df_ls_coa.columns[5]      # Column F (0-indexed = 5)
            
            print(f"\nLoading LS COA sheet:")
            print(f"  Internal column (D): '{internal_col_coa}'")
            print(f"  Account Name column (F): '{name_col_coa}'")
            
            for _, row in df_ls_coa.iterrows():
                internal_val = row[internal_col_coa]
                account_name = row[name_col_coa]
                
                if pd.notna(internal_val) and pd.notna(account_name):
                    ls_coa_internal_to_name[internal_val] = account_name
                    # Also store with type conversions for reliable lookup
                    try:
                        int_internal = int(internal_val)
                        ls_coa_internal_to_name[int_internal] = account_name
                        ls_coa_internal_to_name[str(int_internal)] = account_name
                    except (ValueError, TypeError):
                        pass
            
            print(f"  Loaded {len(ls_coa_internal_to_name) // 3} account names from LS COA sheet")
        else:
            print(f"  Warning: LS COA sheet has fewer than 6 columns, skipping")
    except Exception as e:
        print(f"  Warning: Could not load 'LS COA' sheet: {e}")
    
    return {
        'ic_mapping': ic_mapping,
        'account_to_opposite': account_to_opposite,
        'account_to_internal': account_to_internal,
        'internal_to_opposite_account': internal_to_opposite_account,
        'internal_to_opposite_internal': internal_to_opposite_internal,
        'internal_to_account_name': internal_to_account_name,
        'internal_to_ls_account': internal_to_ls_account,
        'internal_to_account_type': internal_to_account_type,
        'ls_coa_internal_to_name': ls_coa_internal_to_name
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
    4. Apply same column structure to both IC and Non-IC (like IC file)
    5. For IC: duplicate with opposite accounts
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
    
    # Find unmapped internals (for informational purposes only - NO FILTERING)
    unmapped_internals = csv_internals - mapping_internals
    
    if unmapped_internals:
        sorted_unmapped = sorted([x for x in unmapped_internals if isinstance(x, int)])
        print(f"\n  ⚠️  INFO: Found {len(unmapped_internals)} internal IDs in CSV that are NOT in the mapping file:")
        for internal_id in sorted_unmapped[:20]:  # Show first 20
            print(f"      - Internal {internal_id}")
        if len(sorted_unmapped) > 20:
            print(f"      ... and {len(sorted_unmapped) - 20} more")
        print(f"  Keeping ALL rows (no filtering applied)")
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
    
    # === INCLUDE ALL TRANSACTIONS - Identify IC but do NOT duplicate ===
    print(f"\n  Including ALL transactions (no IC duplication)")
    print(f"  Total rows to process: {len(df)}")
    
    # Determine I/C status for each row (for Eliminate column)
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
    ic_count = df['_is_ic_row'].sum()
    non_ic_count = len(df) - ic_count
    print(f"  IC rows: {ic_count}, Non-IC rows: {non_ic_count}")
    
    # All transactions go into df_all - NO duplication
    df_all = df.copy()
    df_ic = pd.DataFrame()  # Empty - IC duplication is disabled
    
    # Store original internal values
    df_all['_original_internal'] = df_all[internal_col]
    if account_col:
        df_all['_original_account'] = df_all[account_col]
    
    # === Update Account and Account Name for ALL transactions based on Internal ID ===
    # Use the FIRST 2 columns of mapping: LS Internal -> LS Account
    if len(df_all) > 0 and internal_to_ls_account:
        print(f"  Applying first 2 columns mapping (LS Internal -> LS Account) for ALL transactions...")
        
        def get_account_from_internal(internal_val):
            """Map Internal ID to LS Account using columns 1-2 of mapping file"""
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
            return None
        
        # Update Account column based on Internal ID -> LS Account mapping
        if account_col:
            mapped_count = 0
            unmapped_count = 0
            
            def update_account_all(row):
                nonlocal mapped_count, unmapped_count
                mapped_account = get_account_from_internal(row[internal_col])
                if mapped_account is not None:
                    mapped_count += 1
                    # Normalize to int if possible for consistency
                    try:
                        return int(float(mapped_account))
                    except (ValueError, TypeError):
                        return mapped_account
                unmapped_count += 1
                return row[account_col]  # Keep original if not found in mapping
            
            df_all[account_col] = df_all.apply(update_account_all, axis=1)
            print(f"    Account column: {mapped_count} rows mapped from Internal->LS Account, {unmapped_count} kept original")
        
        # Update Account Name column based on Internal ID -> Account Name mapping
        if account_name_col and internal_to_account_name:
            name_mapped_count = 0
            
            def get_account_name_from_internal(internal_val):
                """Map Internal ID to Account Name using mapping file"""
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
                return None
            
            def update_account_name_all(row):
                nonlocal name_mapped_count
                mapped_name = get_account_name_from_internal(row[internal_col])
                if mapped_name is not None:
                    name_mapped_count += 1
                    return mapped_name
                return row[account_name_col]  # Keep original if not found
            
            df_all[account_name_col] = df_all.apply(update_account_name_all, axis=1)
            print(f"    Account Name column: {name_mapped_count} rows mapped from Internal->Account Name")
    elif len(df_all) > 0:
        print(f"  Warning: No Internal->LS Account mapping available")
    
    # === Modify DataFrame with additional columns for NetSuite import ===
    # Use explicit English month names to avoid locale issues (Hebrew system)
    MONTH_NAMES = {
        1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'May', 6: 'Jun',
        7: 'Jul', 8: 'Aug', 9: 'Sep', 10: 'Oct', 11: 'Nov', 12: 'Dec'
    }
    
    if len(df_all) > 0:
        # 1. Rename "Subsidiary" to "Subsidiary Header" and add subsidiary-related columns
        if subsidiary_col and subsidiary_col in df_all.columns:
            df_all = df_all.rename(columns={subsidiary_col: 'Subsidiary Header'})
            df_all['Subsidiary Header'] = 6
            
            cols = list(df_all.columns)
            sub_idx = cols.index('Subsidiary Header')
            df_all.insert(sub_idx + 1, 'Subsidiary Line', 6)
            
            cols = list(df_all.columns)
            sub_line_idx = cols.index('Subsidiary Line')
            df_all.insert(sub_line_idx + 1, 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY', '')
        
        # 2. Clear CLASS column
        if class_col and class_col in df_all.columns:
            df_all[class_col] = ''
        
        # 3. Set Location (Expenses) (Req) to "Other"
        if location_col and location_col in df_all.columns:
            df_all[location_col] = 'Other'
        
        # 4. Add "Name" column (empty)
        df_all['Name'] = ''
        
        # 5. Add "Posting Period" column based on Date
        if date_col and date_col in df_all.columns:
            def format_posting_period_all(date_val):
                try:
                    parsed_date = pd.to_datetime(date_val, dayfirst=True)
                    month_num = int(parsed_date.month)
                    year_num = int(parsed_date.year)
                    month_str = MONTH_NAMES.get(month_num, 'Jan')
                    return str(month_str) + ' ' + str(year_num)
                except:
                    return ''
            df_all['Posting Period'] = df_all[date_col].apply(format_posting_period_all)
        else:
            df_all['Posting Period'] = ''
        
        # 6. Add "Eliminate" column - 'Yes' for IC rows, 'No' for non-IC rows
        df_all['Eliminate'] = df_all['_is_ic_row'].apply(lambda x: 'Yes' if x else 'No')
        
        # 7. Add "Department" column with 24 for all
        df_all['Department'] = 24
        
        # 8. Add "EXTERNAL ID" column with "JE_AllMMYY"
        if date_col and date_col in df_all.columns:
            try:
                dates_all = pd.to_datetime(df_all[date_col], dayfirst=True, errors='coerce')
                newest_date_all = dates_all.max()
                if pd.notna(newest_date_all):
                    month_num_all = int(newest_date_all.month)
                    year_short_all = int(newest_date_all.year) % 100
                    external_id_all = f'JE_All{month_num_all:02d}{year_short_all:02d}'
                    df_all.insert(0, 'EXTERNAL ID', external_id_all)
                    print(f"  EXTERNAL ID set to: {external_id_all}")
                else:
                    df_all.insert(0, 'EXTERNAL ID', '')
            except Exception as e:
                print(f"  Warning: Could not determine month for EXTERNAL ID: {e}")
                df_all.insert(0, 'EXTERNAL ID', '')
        else:
            df_all.insert(0, 'EXTERNAL ID', '')
        
        print(f"  Applied column structure to ALL transactions")
    
    # === Remove helper columns ===
    helper_cols = ['_original_internal', '_is_ic_row']
    if '_original_account' in df_all.columns:
        helper_cols.append('_original_account')
    df_all = df_all.drop(columns=[c for c in helper_cols if c in df_all.columns])
    
    # Find debit and credit columns (for reference only - NO FILTERING)
    debit_col = None
    credit_col = None
    for col in df_all.columns:
        col_clean = str(col).strip().lower()
        if 'debit' in col_clean:
            debit_col = col
        elif 'credit' in col_clean:
            credit_col = col
    
    # NO FILTERING - Include ALL rows regardless of debit/credit values
    # (Previously filtered out rows where both debit and credit are zero)
    print(f"  Including ALL rows (no debit/credit filtering)")
    
    # Sort by REF column
    if ref_col and ref_col in df_all.columns:
        df_all = df_all.sort_values(by=ref_col).reset_index(drop=True)
        print(f"  Sorted by {ref_col} column")
    
    # Get unique REF count
    unique_refs = df_all[ref_col].nunique() if ref_col else 0
    
    print(f"\nResults:")
    print(f"  Total transactions: {len(df_all)} rows")
    print(f"  Unique REF batches: {unique_refs}")
    
    # Return df_all as the main result, and empty DataFrame for IC (IC processing disabled)
    return df_all, df_ic


def save_csv(df, output_path):
    """Save DataFrame to CSV file with UTF-8 encoding.
    Column 'internal' (column C) is forced to TEXT so Excel does not treat it as numeric.
    """
    df = df.copy()
    # Find internal column (name contains 'internal', not 'opposite') and force as string (TEXT not numeric)
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if 'internal' in col_clean and 'opposite' not in col_clean and 'oposite' not in col_clean:
            df[col] = df[col].astype(str)
            break
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


def create_ic_pivot_summary(df_ic, output_dir, mapping_data=None, internal_col=None):
    """
    Create an Excel file with IC transactions summary for elimination journal entry.
    Contains 3 sheets:
    - Sheet 1: Original Transactions - the original IC rows (Name=11064, Subsidiary Line=6)
    - Sheet 2: Elimination Transactions - the duplicated rows with opposite accounts (Name=11065, Subsidiary Line=3)
    - Sheet 3: Pivot Summary - breakdown by account with totals
    
    Account names are looked up from the mapping file based on internal ID.
    
    Args:
        df_ic: DataFrame with IC transactions (includes both original and duplicated)
        output_dir: Directory to save the summary file
        mapping_data: Dictionary containing mapping data including internal_to_account_name
        internal_col: Name of the internal ID column
    
    Returns:
        Path to the summary file
    """
    if len(df_ic) == 0:
        return None
    
    # Get the Statscore Import folder on desktop
    statscore_folder = get_statscore_import_folder()
    
    # Extract mappings from mapping_data
    internal_to_account_name = {}
    internal_to_account_type = {}
    ic_mapping = {}
    internal_to_opposite_internal = {}
    internal_to_opposite_account = {}
    internal_to_ls_account = {}
    ls_coa_internal_to_name = {}  # Account names from LS COA sheet
    if mapping_data:
        if 'internal_to_account_name' in mapping_data:
            internal_to_account_name = mapping_data['internal_to_account_name']
        if 'internal_to_account_type' in mapping_data:
            internal_to_account_type = mapping_data['internal_to_account_type']
        if 'ic_mapping' in mapping_data:
            ic_mapping = mapping_data['ic_mapping']
        if 'internal_to_opposite_internal' in mapping_data:
            internal_to_opposite_internal = mapping_data['internal_to_opposite_internal']
        if 'internal_to_opposite_account' in mapping_data:
            internal_to_opposite_account = mapping_data['internal_to_opposite_account']
        if 'internal_to_ls_account' in mapping_data:
            internal_to_ls_account = mapping_data['internal_to_ls_account']
        if 'ls_coa_internal_to_name' in mapping_data:
            ls_coa_internal_to_name = mapping_data['ls_coa_internal_to_name']
    
    # Note: Previously filtered for P&L account types only, now including all I/C accounts for balanced entry
    
    # Find debit and credit columns
    debit_col = None
    credit_col = None
    account_col = None
    name_col = None
    subsidiary_line_col = None
    
    for col in df_ic.columns:
        col_clean = str(col).strip().lower()
        if 'debit' in col_clean and debit_col is None:
            debit_col = col
        elif 'credit' in col_clean and credit_col is None:
            credit_col = col
        elif (col_clean == 'account' or col_clean == 'ls account') and account_col is None:
            account_col = col
        elif col_clean == 'name' and name_col is None:
            name_col = col
        elif col_clean == 'subsidiary line' and subsidiary_line_col is None:
            subsidiary_line_col = col
    
    if debit_col is None or credit_col is None:
        print("  Warning: Could not find debit/credit columns for IC summary")
        return None
    
    # Convert to numeric
    df_ic = df_ic.copy()
    df_ic['_debit_num'] = pd.to_numeric(df_ic[debit_col], errors='coerce').fillna(0)
    df_ic['_credit_num'] = pd.to_numeric(df_ic[credit_col], errors='coerce').fillna(0)
    
    # Find the internal column in IC data
    if internal_col is None:
        for col in df_ic.columns:
            if 'internal' in str(col).lower():
                internal_col = col
                break
    
    # Add Account Name column based on internal ID from LS COA sheet (primary) or mapping file (fallback)
    def get_account_name_from_mapping(internal_val):
        # First try LS COA sheet (column D -> column F)
        if ls_coa_internal_to_name:
            if internal_val in ls_coa_internal_to_name:
                return ls_coa_internal_to_name[internal_val]
            try:
                int_val = int(internal_val)
                if int_val in ls_coa_internal_to_name:
                    return ls_coa_internal_to_name[int_val]
                if str(int_val) in ls_coa_internal_to_name:
                    return ls_coa_internal_to_name[str(int_val)]
            except (ValueError, TypeError):
                pass
        
        # Fallback to original mapping file account names
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
        return ''  # Return empty if not found
    
    # Create Account Name column from LS COA sheet or mapping file
    if internal_col and internal_col in df_ic.columns and (ls_coa_internal_to_name or internal_to_account_name):
        df_ic['Account Name (from Mapping)'] = df_ic[internal_col].apply(get_account_name_from_mapping)
        account_name_col = 'Account Name (from Mapping)'
        if ls_coa_internal_to_name:
            print(f"  Added Account Name column from LS COA sheet")
        else:
            print(f"  Added Account Name column from mapping file")
    else:
        account_name_col = None
    
    # === Split into Original and Elimination transactions ===
    # Original transactions: Name=11064 (LSPORTS) or Subsidiary Line=6
    # Elimination transactions: Name=11065 (STATSCORE) or Subsidiary Line=3
    
    if name_col and name_col in df_ic.columns:
        # Use Name column to split
        df_original = df_ic[df_ic[name_col] == 11064].copy()
        df_elimination = df_ic[df_ic[name_col] == 11065].copy()
        print(f"  Split by Name column: {len(df_original)} original, {len(df_elimination)} elimination")
    elif subsidiary_line_col and subsidiary_line_col in df_ic.columns:
        # Use Subsidiary Line column to split
        df_original = df_ic[df_ic[subsidiary_line_col] == 6].copy()
        df_elimination = df_ic[df_ic[subsidiary_line_col] == 3].copy()
        print(f"  Split by Subsidiary Line column: {len(df_original)} original, {len(df_elimination)} elimination")
    else:
        # Fallback: split in half (first half is original, second half is elimination)
        half_len = len(df_ic) // 2
        df_original = df_ic.iloc[:half_len].copy()
        df_elimination = df_ic.iloc[half_len:].copy()
        print(f"  Split by position: {len(df_original)} original, {len(df_elimination)} elimination")
    
    # === Swap debit and credit in elimination sheet ===
    # Elimination entries should have opposite debit/credit compared to original
    if debit_col and credit_col and debit_col in df_elimination.columns and credit_col in df_elimination.columns:
        # Swap the values
        df_elimination[debit_col], df_elimination[credit_col] = df_elimination[credit_col].copy(), df_elimination[debit_col].copy()
        print(f"  Swapped debit/credit in Elimination Transactions sheet")
    
    # === Apply special mappings for non-eliminated transactions based on memo/description patterns ===
    # Find memo or description column
    memo_col = None
    for col in df_elimination.columns:
        col_clean = str(col).strip().lower()
        if 'memo' in col_clean or 'description' in col_clean or 'line memo' in col_clean:
            memo_col = col
            break
    
    # Special internal ID mappings (internal -> opposite_internal, opposite_account)
    SPECIAL_INTERNAL_MAPPINGS = {
        244: (601, 120002),
    }
    
    # Special mappings for non-eliminated transactions based on memo patterns
    # These are exact/partial matches - ONLY used as fallback when mapping file doesn't have opposite values
    SPECIAL_MEMO_MAPPINGS = [
        # (pattern list, opposite_internal, opposite_account)
        (['lsportdataltd20200065', 'lsportdataltdift'], 913, 650004),  # Specific LSportDataLtd patterns
        (['lsportdev'], 914, 650005),
        (['lsportsupport'], 917, 650008),
        (['lsportdatafeed', 'lsportsfeed'], 915, 650006),
        (['lsportsup'], 377, 660001),
        (['lsportdataltddistribution', 'lsportdistribution'], 912, 650003),  # Distribution patterns -> 912
        (['marketingcloudaccountengagement'], 1073, 640002),
    ]
    
    # Wildcard suffix patterns (*DEV, *feed, *SUP, etc.)
    # These match if the memo ENDS with or CONTAINS the pattern
    WILDCARD_SUFFIX_MAPPINGS = [
        # (suffix_pattern, opposite_internal, opposite_account)
        ('dev', 914, 650005),           # *DEV
        ('feed', 915, 650006),          # *feed
        ('sup', 917, 650008),           # *SUP (but not lsportsup which is 377)
        ('distribution', 912, 650003),  # *distribution
        ('disribution', 912, 650003),   # *disribution (typo variant)
        ('ftp', 1073, 640002),          # *FTP
        ('anulowanefakturys', 912, 650003),  # *Anulowanefakturys*
        ('ystrybucja', 912, 650003),    # *ystrybucja
        ('oddservice', 915, 650006),    # *OddService* -> same as LSportDatafeed
        ('database', 915, 650006),      # *database* -> same as Lsportsfeed
    ]
    
    def get_mapping_file_opposite(internal_val):
        """
        Check if the mapping file has opposite internal and account for this internal ID.
        Returns (opposite_internal, opposite_account) or None if not found.
        """
        opp_internal = None
        opp_account = None
        
        # Try to find opposite internal from mapping
        if internal_val in internal_to_opposite_internal:
            opp_internal = internal_to_opposite_internal[internal_val]
        else:
            try:
                int_val = int(internal_val)
                if int_val in internal_to_opposite_internal:
                    opp_internal = internal_to_opposite_internal[int_val]
                elif str(int_val) in internal_to_opposite_internal:
                    opp_internal = internal_to_opposite_internal[str(int_val)]
            except (ValueError, TypeError):
                pass
        
        # If we found opposite internal, get its account from mapping
        if opp_internal is not None:
            # Get the LS account for the opposite internal
            if opp_internal in internal_to_ls_account:
                opp_account = internal_to_ls_account[opp_internal]
            else:
                try:
                    int_opp = int(opp_internal)
                    if int_opp in internal_to_ls_account:
                        opp_account = internal_to_ls_account[int_opp]
                    elif str(int_opp) in internal_to_ls_account:
                        opp_account = internal_to_ls_account[str(int_opp)]
                except (ValueError, TypeError):
                    pass
            
            # Also try internal_to_opposite_account directly
            if opp_account is None:
                if internal_val in internal_to_opposite_account:
                    opp_account = internal_to_opposite_account[internal_val]
                else:
                    try:
                        int_val = int(internal_val)
                        if int_val in internal_to_opposite_account:
                            opp_account = internal_to_opposite_account[int_val]
                        elif str(int_val) in internal_to_opposite_account:
                            opp_account = internal_to_opposite_account[str(int_val)]
                    except (ValueError, TypeError):
                        pass
        
        if opp_internal is not None and opp_account is not None:
            return (opp_internal, opp_account)
        
        return None
    
    def get_hardcoded_mapping_by_internal(internal_val):
        """Check if internal ID has a hardcoded special mapping (fallback)"""
        try:
            int_val = int(internal_val)
            if int_val in SPECIAL_INTERNAL_MAPPINGS:
                return SPECIAL_INTERNAL_MAPPINGS[int_val]
        except (ValueError, TypeError):
            pass
        return None
    
    def get_special_mapping(memo_val):
        """Check if memo matches any special pattern and return (internal, account) or None
        NOTE: This is only used as a FALLBACK when mapping file doesn't have opposite values"""
        if pd.isna(memo_val):
            return None
        
        memo_lower = str(memo_val).lower().replace(' ', '')
        
        # Check for MarketingCloudAccountEngagement followed by anything
        if 'marketingcloudaccountengagement' in memo_lower:
            return (1073, 640002)
        
        # Check for lsportsup specifically (different mapping than *sup)
        if 'lsportsup' in memo_lower:
            return (377, 660001)
        
        # Check exact/partial pattern matches first
        for patterns, opp_internal, opp_account in SPECIAL_MEMO_MAPPINGS:
            for pattern in patterns:
                if pattern in memo_lower:
                    return (opp_internal, opp_account)
        
        # Check for LSportDataLtd followed by digits (like LSportDataLtd20200065) - maps to 912/650003
        if re.match(r'lsportdataltd\d+', memo_lower):
            return (912, 650003)
        
        # Check wildcard suffix patterns
        for suffix, opp_internal, opp_account in WILDCARD_SUFFIX_MAPPINGS:
            # Match if memo ends with suffix or contains it
            if memo_lower.endswith(suffix) or suffix in memo_lower:
                return (opp_internal, opp_account)
        
        return None
    
    # Apply special mappings to elimination sheet for non-IC transactions
    # Find Eliminate column
    eliminate_col = None
    for col in df_elimination.columns:
        if str(col).strip().lower() == 'eliminate':
            eliminate_col = col
            break
    
    if eliminate_col and eliminate_col in df_elimination.columns:
        # Apply mappings to rows where Eliminate = 'No'
        def apply_special_mapping(row):
            if str(row.get(eliminate_col, '')).strip().lower() == 'no':
                internal_val = row.get(internal_col, '') if internal_col and internal_col in row.index else ''
                
                # PRIORITY 1: Check the mapping file for opposite internal/account
                if internal_val:
                    mapping_file_result = get_mapping_file_opposite(internal_val)
                    if mapping_file_result:
                        return pd.Series({'_special_internal': mapping_file_result[0], '_special_account': mapping_file_result[1]})
                
                # PRIORITY 2: Check hardcoded internal ID mappings (fallback)
                if internal_val:
                    hardcoded_internal = get_hardcoded_mapping_by_internal(internal_val)
                    if hardcoded_internal:
                        return pd.Series({'_special_internal': hardcoded_internal[0], '_special_account': hardcoded_internal[1]})
                
                # PRIORITY 3: Check memo/description patterns (fallback)
                if memo_col and memo_col in row.index:
                    memo_mapping = get_special_mapping(row.get(memo_col, ''))
                    if memo_mapping:
                        return pd.Series({'_special_internal': memo_mapping[0], '_special_account': memo_mapping[1]})
            
            return pd.Series({'_special_internal': None, '_special_account': None})
        
        special_mappings = df_elimination.apply(apply_special_mapping, axis=1)
        
        # Update internal and account for matched rows
        mask = special_mappings['_special_internal'].notna()
        if mask.any():
            if internal_col and internal_col in df_elimination.columns:
                df_elimination.loc[mask, internal_col] = special_mappings.loc[mask, '_special_internal']
            if account_col and account_col in df_elimination.columns:
                df_elimination.loc[mask, account_col] = special_mappings.loc[mask, '_special_account']
            
            # Also update Account Name for these rows
            if account_name_col and account_name_col in df_elimination.columns:
                for idx in df_elimination[mask].index:
                    new_internal = special_mappings.loc[idx, '_special_internal']
                    new_name = get_account_name_from_mapping(new_internal)
                    if new_name:
                        df_elimination.loc[idx, account_name_col] = new_name
            
            matched_count = mask.sum()
            print(f"  Applied special mappings to {matched_count} non-eliminated transactions")
    else:
        print(f"  Warning: Could not find 'Eliminate' column for special mappings")
    
    # Debug: print found columns
    print(f"  IC Summary - Found columns:")
    print(f"    Internal: {internal_col}")
    print(f"    Account: {account_col}")
    print(f"    Account Name (from mapping): {account_name_col}")
    print(f"    Name: {name_col}")
    print(f"    Subsidiary Line: {subsidiary_line_col}")
    
    # === Create Pivot Table based on Elimination Transactions ===
    # Recalculate debit/credit numeric values from the elimination sheet (after swap and mappings)
    df_elimination_pivot = df_elimination.copy()
    df_elimination_pivot['_debit_num'] = pd.to_numeric(df_elimination_pivot[debit_col], errors='coerce').fillna(0)
    df_elimination_pivot['_credit_num'] = pd.to_numeric(df_elimination_pivot[credit_col], errors='coerce').fillna(0)
    
    # === Include ALL I/C accounts in pivot (for balanced elimination entry) ===
    # No filtering - use all elimination transactions for a balanced entry
    print(f"  Using all elimination transactions for balanced pivot: {len(df_elimination_pivot)} rows")
    
    # Create pivot table grouped by Internal ID, Account, and Account Name
    group_cols = []
    if internal_col and internal_col in df_elimination_pivot.columns:
        group_cols.append(internal_col)
    if account_col and account_col in df_elimination_pivot.columns:
        group_cols.append(account_col)
    if account_name_col and account_name_col in df_elimination_pivot.columns:
        group_cols.append(account_name_col)
    
    if not group_cols:
        group_cols = [df_elimination_pivot.columns[0]]  # Fallback to first column
    
    print(f"    Grouping by: {group_cols}")
    print(f"    Pivot based on Elimination Transactions (with swapped debit/credit)")
    
    # Create breakdown pivot table (using Elimination Transactions data)
    pivot_breakdown = df_elimination_pivot.groupby(group_cols, dropna=False).agg({
        '_debit_num': 'sum',
        '_credit_num': 'sum'
    }).reset_index()
    
    pivot_breakdown = pivot_breakdown.rename(columns={
        '_debit_num': 'Total Debit',
        '_credit_num': 'Total Credit'
    })
    
    # Calculate net amount (Debit - Credit)
    pivot_breakdown['Net Amount'] = pivot_breakdown['Total Debit'] - pivot_breakdown['Total Credit']
    
    # Filter out rows where both debit and credit are zero
    pivot_breakdown = pivot_breakdown[
        (pivot_breakdown['Total Debit'] != 0) | (pivot_breakdown['Total Credit'] != 0)
    ]
    
    # Sort by Internal ID
    if internal_col and internal_col in pivot_breakdown.columns:
        pivot_breakdown = pivot_breakdown.sort_values(by=internal_col)
    
    # Create summary totals
    total_debit = pivot_breakdown['Total Debit'].sum()
    total_credit = pivot_breakdown['Total Credit'].sum()
    total_net = pivot_breakdown['Net Amount'].sum()
    
    # Check if balanced
    is_balanced = abs(total_debit - total_credit) < 0.01
    print(f"  Pivot Summary - Total Debit: {total_debit:,.2f}, Total Credit: {total_credit:,.2f}")
    print(f"  Balanced: {'YES' if is_balanced else 'NO (difference: ' + str(total_debit - total_credit) + ')'}")
    
    # Remove helper columns from sheet data (AFTER pivot is created)
    helper_cols_to_remove = ['_debit_num', '_credit_num']
    df_original_clean = df_original.drop(columns=[c for c in helper_cols_to_remove if c in df_original.columns])
    df_elimination_clean = df_elimination.drop(columns=[c for c in helper_cols_to_remove if c in df_elimination.columns])
    
    # Add totals row
    totals_row = {col: '' for col in pivot_breakdown.columns}
    totals_row[pivot_breakdown.columns[0]] = 'TOTAL'
    totals_row['Total Debit'] = total_debit
    totals_row['Total Credit'] = total_credit
    totals_row['Net Amount'] = total_net
    
    pivot_with_totals = pd.concat([pivot_breakdown, pd.DataFrame([totals_row])], ignore_index=True)
    
    # Add empty rows and summary section
    empty_row = {col: '' for col in pivot_with_totals.columns}
    summary_header = {col: '' for col in pivot_with_totals.columns}
    summary_header[pivot_with_totals.columns[0]] = '=== ELIMINATION JOURNAL ENTRY ==='
    
    summary_debit = {col: '' for col in pivot_with_totals.columns}
    summary_debit[pivot_with_totals.columns[0]] = 'Journal Entry to Eliminate IC'
    summary_debit['Total Debit'] = total_credit  # Reverse: Credit becomes Debit
    summary_debit['Total Credit'] = total_debit  # Reverse: Debit becomes Credit
    summary_debit['Net Amount'] = -total_net
    
    # Combine all sections for pivot summary
    final_summary = pd.concat([
        pivot_with_totals,
        pd.DataFrame([empty_row]),
        pd.DataFrame([empty_row]),
        pd.DataFrame([summary_header]),
        pd.DataFrame([summary_debit])
    ], ignore_index=True)
    
    # Save to Excel file with 3 sheets
    output_path = os.path.join(statscore_folder, "IC_Elimination_Summary.xlsx")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # Sheet 1: Original Transactions
        df_original_clean.to_excel(writer, sheet_name='Original Transactions', index=False)
        
        # Sheet 2: Elimination Transactions
        df_elimination_clean.to_excel(writer, sheet_name='Elimination Transactions', index=False)
        
        # Sheet 3: Pivot Summary
        final_summary.to_excel(writer, sheet_name='Pivot Summary', index=False)
        
        # Set column C (internal) to TEXT format so Excel does not treat it as numeric
        wb = writer.book
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for cell in ws['C']:
                cell.number_format = '@'
    
    print(f"\n  === IC Elimination Summary ===")
    print(f"  Original Transactions: {len(df_original_clean)} rows")
    print(f"  Elimination Transactions: {len(df_elimination_clean)} rows")
    print(f"  Total Debit: {total_debit:,.2f}")
    print(f"  Total Credit: {total_credit:,.2f}")
    print(f"  Net Amount: {total_net:,.2f}")
    print(f"  Saved to: {output_path}")
    
    return output_path


def save_all_transactions_by_month(df_all):
    """
    Save ALL transactions DataFrame by month.
    Creates one file per month with all transactions.
    Files are saved to "Statscore Import" folder on the desktop.
    
    Args:
        df_all: All transactions DataFrame
    
    Returns:
        List of saved file paths and the output folder
    """
    saved_files = []
    
    if len(df_all) == 0:
        print("  No transactions to save")
        return saved_files, get_statscore_import_folder()
    
    # Get the Statscore Import folder on desktop
    statscore_folder = get_statscore_import_folder()
    
    # Find the Posting Period column
    posting_period_col = None
    for col in df_all.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    if posting_period_col is None:
        print(f"  Warning: Could not find 'Posting Period' column. Saving all to single file.")
        output_path = os.path.join(statscore_folder, "NetSuite_All_Transactions.csv")
        save_csv(df_all, output_path)
        return [('All', output_path, len(df_all))], statscore_folder
    
    # Get all unique posting periods
    all_periods = df_all[posting_period_col].dropna().unique()
    
    print(f"\n  Saving ALL transactions by month...")
    print(f"  Found {len(all_periods)} unique posting periods")
    print(f"  Saving to: {statscore_folder}")
    
    for posting_period in sorted(all_periods):
        # Get rows for this period
        period_df = df_all[df_all[posting_period_col] == posting_period].copy()
        
        if len(period_df) == 0:
            continue
        
        # Update EXTERNAL ID for this specific month
        date_col = None
        for col in period_df.columns:
            if str(col).strip().lower() == 'date':
                date_col = col
                break
        
        if date_col and 'EXTERNAL ID' in period_df.columns:
            try:
                dates = pd.to_datetime(period_df[date_col], dayfirst=True, errors='coerce')
                newest_date = dates.max()
                if pd.notna(newest_date):
                    month_num = int(newest_date.month)
                    year_short = int(newest_date.year) % 100
                    external_id = f'JE_All{month_num:02d}{year_short:02d}'
                    period_df['EXTERNAL ID'] = external_id
            except:
                pass
        
        # Get the filename for this month
        filename = get_month_filename(posting_period, prefix="NetSuite_All")
        
        if filename is None:
            safe_period = str(posting_period).replace('/', '-').replace('\\', '-').replace(' ', '_')
            filename = f"NetSuite_All_{safe_period}"
        
        output_path = os.path.join(statscore_folder, f"{filename}.csv")
        save_csv(period_df, output_path)
        
        saved_files.append((posting_period, output_path, len(period_df)))
    
    return saved_files, statscore_folder


def save_non_ic_by_month(df_non_ic):
    """
    Save Non-IC DataFrame by month.
    Creates one file per month with non-IC transactions only.
    Files are saved to "Statscore Import" folder on the desktop.
    
    Args:
        df_non_ic: Non-IC transactions DataFrame
    
    Returns:
        List of saved file paths and the output folder
    """
    saved_files = []
    
    if len(df_non_ic) == 0:
        print("  No Non-IC transactions to save")
        return saved_files, get_statscore_import_folder()
    
    # Get the Statscore Import folder on desktop
    statscore_folder = get_statscore_import_folder()
    
    # Find the Posting Period column
    posting_period_col = None
    for col in df_non_ic.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    if posting_period_col is None:
        print(f"  Warning: Could not find 'Posting Period' column. Saving all to single file.")
        output_path = os.path.join(statscore_folder, "NetSuite_NonIC_Transactions.csv")
        save_csv(df_non_ic, output_path)
        return [('All', output_path, len(df_non_ic))], statscore_folder
    
    # Get all unique posting periods
    all_periods = df_non_ic[posting_period_col].dropna().unique()
    
    print(f"\n  Saving Non-IC transactions by month (IC transactions ignored)...")
    print(f"  Found {len(all_periods)} unique posting periods")
    print(f"  Saving to: {statscore_folder}")
    
    for posting_period in sorted(all_periods):
        # Get Non-IC rows for this period
        non_ic_period = df_non_ic[df_non_ic[posting_period_col] == posting_period].copy()
        
        if len(non_ic_period) == 0:
            continue
        
        # Update EXTERNAL ID for this specific month
        date_col = None
        for col in non_ic_period.columns:
            if str(col).strip().lower() == 'date':
                date_col = col
                break
        
        if date_col and 'EXTERNAL ID' in non_ic_period.columns:
            try:
                dates = pd.to_datetime(non_ic_period[date_col], dayfirst=True, errors='coerce')
                newest_date = dates.max()
                if pd.notna(newest_date):
                    month_num = int(newest_date.month)
                    year_short = int(newest_date.year) % 100
                    external_id = f'JE_nonIC{month_num:02d}{year_short:02d}'
                    non_ic_period['EXTERNAL ID'] = external_id
            except:
                pass
        
        # Get the filename for this month
        filename = get_month_filename(posting_period, prefix="NetSuite_NonIC")
        
        if filename is None:
            safe_period = str(posting_period).replace('/', '-').replace('\\', '-').replace(' ', '_')
            filename = f"NetSuite_NonIC_{safe_period}"
        
        output_path = os.path.join(statscore_folder, f"{filename}.csv")
        save_csv(non_ic_period, output_path)
        
        saved_files.append((posting_period, output_path, len(non_ic_period)))
    
    return saved_files, statscore_folder


def save_combined_by_month(df_ic, df_non_ic):
    """
    Combine IC and Non-IC DataFrames and save by month.
    Creates one file per month with all transactions.
    Files are saved to "Statscore Import" folder on the desktop.
    
    Args:
        df_ic: IC transactions DataFrame
        df_non_ic: Non-IC transactions DataFrame
    
    Returns:
        List of saved file paths and the output folder
    """
    saved_files = []
    
    # Get the Statscore Import folder on desktop
    statscore_folder = get_statscore_import_folder()
    
    # Find the Posting Period column
    posting_period_col = None
    sample_df = df_ic if len(df_ic) > 0 else df_non_ic
    for col in sample_df.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    if posting_period_col is None:
        print(f"  Warning: Could not find 'Posting Period' column. Saving all to single file.")
        # Combine all into one file
        combined = pd.concat([df_ic, df_non_ic], ignore_index=True)
        output_path = os.path.join(statscore_folder, "NetSuite_All_Transactions.csv")
        save_csv(combined, output_path)
        return [('All', output_path, len(combined))], statscore_folder
    
    # Get all unique posting periods from both DataFrames
    all_periods = set()
    if len(df_ic) > 0 and posting_period_col in df_ic.columns:
        all_periods.update(df_ic[posting_period_col].dropna().unique())
    if len(df_non_ic) > 0 and posting_period_col in df_non_ic.columns:
        all_periods.update(df_non_ic[posting_period_col].dropna().unique())
    
    print(f"\n  Combining IC and Non-IC transactions by month...")
    print(f"  Found {len(all_periods)} unique posting periods")
    print(f"  Saving to: {statscore_folder}")
    
    for posting_period in sorted(all_periods):
        # Get IC rows for this period
        ic_period = pd.DataFrame()
        if len(df_ic) > 0 and posting_period_col in df_ic.columns:
            ic_period = df_ic[df_ic[posting_period_col] == posting_period].copy()
        
        # Get Non-IC rows for this period
        non_ic_period = pd.DataFrame()
        if len(df_non_ic) > 0 and posting_period_col in df_non_ic.columns:
            non_ic_period = df_non_ic[df_non_ic[posting_period_col] == posting_period].copy()
        
        # Combine for this month
        combined_period = pd.concat([ic_period, non_ic_period], ignore_index=True)
        
        if len(combined_period) == 0:
            continue
        
        # Update EXTERNAL ID for this specific month
        date_col = None
        for col in combined_period.columns:
            if str(col).strip().lower() == 'date':
                date_col = col
                break
        
        if date_col and 'EXTERNAL ID' in combined_period.columns:
            try:
                dates = pd.to_datetime(combined_period[date_col], dayfirst=True, errors='coerce')
                newest_date = dates.max()
                if pd.notna(newest_date):
                    month_num = int(newest_date.month)
                    year_short = int(newest_date.year) % 100
                    external_id = f'JE_All{month_num:02d}{year_short:02d}'
                    combined_period['EXTERNAL ID'] = external_id
            except:
                pass
        
        # Get the filename for this month
        filename = get_month_filename(posting_period, prefix="NetSuite_All_Transactions")
        
        if filename is None:
            safe_period = str(posting_period).replace('/', '-').replace('\\', '-').replace(' ', '_')
            filename = f"NetSuite_All_Transactions_{safe_period}"
        
        output_path = os.path.join(statscore_folder, f"{filename}.csv")
        save_csv(combined_period, output_path)
        
        ic_count = len(ic_period)
        non_ic_count = len(non_ic_period)
        saved_files.append((posting_period, output_path, len(combined_period), ic_count, non_ic_count))
    
    return saved_files, statscore_folder


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
        
        # Process CSV - returns DataFrame with ALL transactions (IC classification ignored)
        print("\nProcessing CSV file...")
        df_all, df_ic = process_csv(csv_file, mapping_dict)
        
        # Generate output directory
        output_dir = os.path.dirname(csv_file)
        
        # Save results
        print("\nSaving results...")
        
        # Save ALL transactions (IC classification is ignored)
        print(f"\n  Processing ALL transactions ({len(df_all)} rows)")
        
        all_files, statscore_folder = save_all_transactions_by_month(df_all)
        
        # Build success message for files
        files_msg = ""
        if len(all_files) > 0:
            for item in all_files:
                if isinstance(item, tuple):
                    posting_period, path, count = item[:3]
                    filename = os.path.basename(path)
                    files_msg += f"  {filename} ({count} rows)\n"
                else:
                    files_msg += f"  {os.path.basename(item)}\n"
        
        # Show success message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", 
            f"Processing complete!\n\n"
            f"ALL transactions included (IC classification ignored)\n"
            f"Total rows: {len(df_all)}\n\n"
            f"Files saved to:\n{statscore_folder}\n\n"
            f"Files by month:\n"
            f"{files_msg}")
        root.destroy()
        
        print("\nDone!")
        
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"An error occurred:\n{e}")
        root.destroy()


if __name__ == "__main__":
    main()

