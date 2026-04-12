"""
NetSuite IC Elimination Journal Entry Creator
- Creates elimination JE for specific IC accounts with opposite mappings
- Only includes transactions where Internal ID has an Opposite account defined
- Swaps Debit/Credit and maps to opposite account/internal
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

def select_files():
    """Open GUI to select CSV and mapping files"""
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
    
    root = tk.Tk()
    root.title("IC Elimination JE Creator")
    root.geometry("650x250")
    root.resizable(False, False)
    
    root.update_idletasks()
    x = (root.winfo_screenwidth() - 650) // 2
    y = (root.winfo_screenheight() - 250) // 2
    root.geometry(f"650x250+{x}+{y}")
    
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    title_label = ttk.Label(main_frame, text="IC Elimination JE Creator", font=('Helvetica', 14, 'bold'))
    title_label.pack(pady=(0, 15))
    
    csv_frame = ttk.Frame(main_frame)
    csv_frame.pack(fill=tk.X, pady=5)
    ttk.Label(csv_frame, text="CSV File:", width=15).pack(side=tk.LEFT)
    csv_entry = ttk.Entry(csv_frame, width=50)
    csv_entry.pack(side=tk.LEFT, padx=5)
    ttk.Button(csv_frame, text="Browse...", command=browse_csv).pack(side=tk.LEFT)
    
    mapping_frame = ttk.Frame(main_frame)
    mapping_frame.pack(fill=tk.X, pady=5)
    ttk.Label(mapping_frame, text="Mapping File:", width=15).pack(side=tk.LEFT)
    mapping_entry = ttk.Entry(mapping_frame, width=50)
    mapping_entry.pack(side=tk.LEFT, padx=5)
    ttk.Button(mapping_frame, text="Browse...", command=browse_mapping).pack(side=tk.LEFT)
    
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=20)
    ttk.Button(button_frame, text="Create Elimination JE", command=process, width=20).pack(side=tk.LEFT, padx=10)
    ttk.Button(button_frame, text="Cancel", command=cancel, width=15).pack(side=tk.LEFT, padx=10)
    
    root.mainloop()
    
    return result['csv'], result['mapping']


def load_mapping(mapping_file):
    """
    Load mapping file and extract accounts where "Nazwa" column contains LSPORT/LSPORTS.
    These are the accounts that trigger inclusion in elimination JE.
    Also identifies Bank type accounts for special handling.
    """
    df_mapping = pd.read_excel(mapping_file)
    
    # Find column names
    internal_col = None
    ls_account_col = None
    opposite_account_col = None
    opposite_internal_col = None
    account_name_col = None
    intercompany_col = None
    nazwa_col = None  # The key column for filtering
    type_col = None   # Account type column for Bank detection
    
    for col in df_mapping.columns:
        col_clean = str(col).strip().lower()
        if 'internal' in col_clean and 'opposite' not in col_clean and internal_col is None:
            internal_col = col
        elif col_clean == 'ls account' or col_clean == 'lsaccount':
            ls_account_col = col
        elif ('opposite' in col_clean or 'oposite' in col_clean) and 'internal' in col_clean:
            opposite_internal_col = col
        elif ('opposite' in col_clean or 'oposite' in col_clean) and 'account' in col_clean:
            opposite_account_col = col
        elif 'account name' in col_clean:
            account_name_col = col
        elif 'intercompany' in col_clean:
            intercompany_col = col
        elif col_clean == 'nazwa':
            nazwa_col = col
        elif col_clean == 'type' or col_clean == 'account type' or col_clean == 'ls account type':
            type_col = col
    
    if internal_col is None:
        raise ValueError("Could not find 'Internal' column")
    if nazwa_col is None:
        raise ValueError("Could not find 'Nazwa' column in mapping file")
    
    print(f"Mapping columns found:")
    print(f"  Internal: '{internal_col}'")
    print(f"  Nazwa: '{nazwa_col}'")
    if ls_account_col:
        print(f"  LS Account: '{ls_account_col}'")
    if opposite_account_col:
        print(f"  Opposite Account: '{opposite_account_col}'")
    if opposite_internal_col:
        print(f"  Opposite Internal: '{opposite_internal_col}'")
    if type_col:
        print(f"  Account Type: '{type_col}'")
    
    # Build mappings - filter by "Nazwa" containing LSPORT/LSPORTS
    lsport_internals = set()  # Internal IDs where Nazwa contains LSPORT
    bank_internals = set()    # Internal IDs where Type is Bank
    internal_to_opposite_internal = {}
    internal_to_opposite_account = {}
    internal_to_account_name = {}
    internal_to_ls_account = {}
    ic_mapping = {}
    
    for _, row in df_mapping.iterrows():
        internal_val = row[internal_col]
        
        # Check if Nazwa contains LSPORT or LSPORTS (case insensitive)
        nazwa_val = str(row[nazwa_col]).strip().upper() if pd.notna(row[nazwa_col]) else ""
        is_lsport = 'LSPORT' in nazwa_val  # This catches both LSPORT and LSPORTS
        
        if is_lsport:
            # Add to LSPORT internals
            lsport_internals.add(internal_val)
            try:
                lsport_internals.add(int(internal_val))
                lsport_internals.add(str(int(internal_val)))
            except (ValueError, TypeError):
                pass
            print(f"    Found LSPORT: Internal {internal_val}, Nazwa: {row[nazwa_col]}")
        
        # Check if Account Type is Bank
        if type_col and pd.notna(row[type_col]):
            type_val = str(row[type_col]).strip().lower()
            if 'bank' in type_val:
                bank_internals.add(internal_val)
                try:
                    bank_internals.add(int(internal_val))
                    bank_internals.add(str(int(internal_val)))
                except (ValueError, TypeError):
                    pass
        
        # Build opposite mappings for ALL accounts (not just LSPORT)
        # This is needed for the elimination entry transformation
        if opposite_internal_col and pd.notna(row[opposite_internal_col]):
            opp_internal = row[opposite_internal_col]
            internal_to_opposite_internal[internal_val] = opp_internal
            try:
                internal_to_opposite_internal[int(internal_val)] = int(opp_internal)
                internal_to_opposite_internal[str(int(internal_val))] = int(opp_internal)
            except (ValueError, TypeError):
                pass
        
        if opposite_account_col and pd.notna(row[opposite_account_col]):
            opp_account = row[opposite_account_col]
            internal_to_opposite_account[internal_val] = opp_account
            try:
                internal_to_opposite_account[int(internal_val)] = opp_account
                internal_to_opposite_account[str(int(internal_val))] = opp_account
            except (ValueError, TypeError):
                pass
        
        # Build IC mapping
        if intercompany_col:
            ic_val = str(row[intercompany_col]).strip().lower() == 'yes'
            ic_mapping[internal_val] = ic_val
            try:
                ic_mapping[int(internal_val)] = ic_val
                ic_mapping[str(int(internal_val))] = ic_val
            except (ValueError, TypeError):
                pass
        
        # Account name mapping
        if account_name_col and pd.notna(row[account_name_col]):
            internal_to_account_name[internal_val] = row[account_name_col]
            try:
                internal_to_account_name[int(internal_val)] = row[account_name_col]
                internal_to_account_name[str(int(internal_val))] = row[account_name_col]
            except (ValueError, TypeError):
                pass
        
        # LS Account mapping
        if ls_account_col and pd.notna(row[ls_account_col]):
            internal_to_ls_account[internal_val] = row[ls_account_col]
            try:
                internal_to_ls_account[int(internal_val)] = row[ls_account_col]
                internal_to_ls_account[str(int(internal_val))] = row[ls_account_col]
            except (ValueError, TypeError):
                pass
    
    lsport_count = len([x for x in lsport_internals if isinstance(x, int) or (isinstance(x, str) and x.isdigit())]) // 2
    bank_count = len([x for x in bank_internals if isinstance(x, int) or (isinstance(x, str) and x.isdigit())]) // 2
    
    print(f"\n  Accounts with LSPORT in Nazwa: {lsport_count}")
    print(f"  Bank type accounts: {bank_count}")
    print(f"  REF batches containing LSPORT accounts will be included in elimination JE")
    print(f"  Bank accounts will be replaced with IC opposite (111 <-> 119)")
    
    return {
        'lsport_internals': lsport_internals,
        'bank_internals': bank_internals,
        'internal_to_opposite_internal': internal_to_opposite_internal,
        'internal_to_opposite_account': internal_to_opposite_account,
        'internal_to_account_name': internal_to_account_name,
        'internal_to_ls_account': internal_to_ls_account,
        'ic_mapping': ic_mapping
    }


def create_elimination_je(csv_file, mapping_data):
    """
    Create elimination JE for IC transactions where Nazwa contains LSPORT.
    Includes ALL rows from REF batches that contain at least one LSPORT account.
    
    Bank Account Handling:
    - For batches containing Bank accounts: consolidate to exactly 2 lines
    - Replace Bank account with IC opposite (111 <-> 119)
    - 111 (account 210000) <-> 119 (account 120003)
    """
    # Read CSV
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(csv_file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='latin-1')
    
    print(f"\nLoaded {len(df)} rows from CSV")
    
    lsport_internals = mapping_data['lsport_internals']
    bank_internals = mapping_data.get('bank_internals', set())
    internal_to_opposite_internal = mapping_data['internal_to_opposite_internal']
    internal_to_opposite_account = mapping_data['internal_to_opposite_account']
    internal_to_account_name = mapping_data['internal_to_account_name']
    
    # IC Account opposite mapping (hardcoded)
    # 111 (account 210000) <-> 119 (account 120003)
    # 998 -> 783 (998 should NEVER appear in output, always replaced with 783)
    # 783 -> 964/121002 (for elimination entries)
    IC_OPPOSITE = {
        111: {'internal': 119, 'account': 120003},
        119: {'internal': 111, 'account': 210000},
        '111': {'internal': 119, 'account': 120003},
        '119': {'internal': 111, 'account': 210000},
        998: {'internal': 783, 'account': 121001},   # 998 becomes 783
        783: {'internal': 964, 'account': 121002},   # 783's opposite is 964/121002 (NOT 998)
        '998': {'internal': 783, 'account': 121001},
        '783': {'internal': 964, 'account': 121002},
    }
    
    # Find columns
    internal_col = None
    ref_col = None
    account_col = None
    account_name_col = None
    date_col = None
    debit_col = None
    credit_col = None
    
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if col_clean == 'internal' or col_clean == 'internal id':
            internal_col = col
        elif col_clean == 'ref':
            ref_col = col
        elif col_clean == 'account':
            account_col = col
        elif 'account name' in col_clean:
            account_name_col = col
        elif col_clean == 'date':
            date_col = col
        elif 'debit' in col_clean:
            debit_col = col
        elif 'credit' in col_clean:
            credit_col = col
    
    if internal_col is None:
        raise ValueError("Could not find 'Internal' column in CSV")
    if ref_col is None:
        raise ValueError("Could not find 'REF' column in CSV")
    if debit_col is None or credit_col is None:
        raise ValueError("Could not find Debit/Credit columns in CSV")
    
    print(f"Using: Internal='{internal_col}', REF='{ref_col}'")
    print(f"Debit='{debit_col}', Credit='{credit_col}'")
    
    # Helper functions
    def is_lsport_account(internal_val):
        if internal_val in lsport_internals:
            return True
        try:
            if int(internal_val) in lsport_internals:
                return True
        except (ValueError, TypeError):
            pass
        return False
    
    def is_bank_account(internal_val):
        if internal_val in bank_internals:
            return True
        try:
            if int(internal_val) in bank_internals:
                return True
        except (ValueError, TypeError):
            pass
        return False
    
    def is_ic_account(internal_val):
        """Check if this is an IC account (111, 119, 998, or 783)"""
        try:
            int_val = int(internal_val)
            return int_val in [111, 119, 998, 783]
        except (ValueError, TypeError):
            return str(internal_val) in ['111', '119', '998', '783']
    
    # Mark LSPORT and Bank rows
    df['_is_lsport'] = df[internal_col].apply(is_lsport_account)
    df['_is_bank'] = df[internal_col].apply(is_bank_account)
    df['_is_ic'] = df[internal_col].apply(is_ic_account)
    
    # Convert debit/credit to numeric
    df[debit_col] = pd.to_numeric(df[debit_col], errors='coerce').fillna(0)
    df[credit_col] = pd.to_numeric(df[credit_col], errors='coerce').fillna(0)
    
    lsport_row_count = df['_is_lsport'].sum()
    bank_row_count = df['_is_bank'].sum()
    print(f"Found {lsport_row_count} rows with LSPORT accounts")
    print(f"Found {bank_row_count} rows with Bank accounts")
    
    # Find all REF batches that contain at least one LSPORT account
    lsport_refs = df[df['_is_lsport']][ref_col].unique()
    print(f"Found {len(lsport_refs)} REF batches containing LSPORT accounts")
    
    # Include ALL rows from those REF batches
    df['_in_lsport_batch'] = df[ref_col].isin(lsport_refs)
    df_eligible = df[df['_in_lsport_batch']].copy()
    
    print(f"Total rows in LSPORT batches: {len(df_eligible)}")
    
    if len(df_eligible) == 0:
        print("No LSPORT batches found for elimination!")
        return pd.DataFrame()
    
    # === PROCESS EACH REF BATCH ===
    # For batches with Bank accounts: consolidate to 2 lines with IC opposite
    # For batches without Bank: swap debit/credit for all rows
    
    result_rows = []
    bank_batches = 0
    regular_batches = 0
    
    for ref_val in lsport_refs:
        batch = df_eligible[df_eligible[ref_col] == ref_val].copy()
        has_bank = batch['_is_bank'].any()
        
        # Get template row (first row for metadata like Date, Memo, etc.)
        template = batch.iloc[0].to_dict()
        
        if has_bank:
            bank_batches += 1
            # === BANK BATCH: Consolidate to 2 lines ===
            # Find the IC account (111 or 119) in this batch
            ic_rows = batch[batch['_is_ic']]
            
            if len(ic_rows) > 0:
                # Get the IC internal (111 or 119)
                ic_internal = ic_rows.iloc[0][internal_col]
                try:
                    ic_internal = int(ic_internal)
                except:
                    pass
                
                # Get the IC row's original amounts
                ic_debit = ic_rows[debit_col].sum()
                ic_credit = ic_rows[credit_col].sum()
                
                # SPECIAL CASE: If the IC account is 783 or 998 (which should be treated as 783)
                # and debit is bank type, then use 783/121001 and 964/121002
                # 998 should NEVER appear in output - always replace with 783
                # Calculate net amount to ensure only one of debit/credit per line
                net_amount = ic_debit - ic_credit  # Positive = net debit, Negative = net credit
                
                # Treat 998 same as 783 - both use 783/121001 and 964/121002 for elimination
                if ic_internal in [783, '783', 998, '998'] and ic_credit > ic_debit:
                    # 783/121001 has net credit, bank has debit
                    # For elimination: swap - 783 gets debit, 964/121002 gets credit
                    row1 = template.copy()
                    row1[internal_col] = 783
                    if account_col:
                        row1[account_col] = 121001
                    row1[debit_col] = abs(net_amount)  # Net credit becomes debit
                    row1[credit_col] = 0
                    
                    # Line 2: Replace bank with 964/121002
                    row2 = template.copy()
                    row2[internal_col] = 964
                    if account_col:
                        row2[account_col] = 121002
                    row2[debit_col] = 0
                    row2[credit_col] = abs(net_amount)  # Same amount on credit side to balance
                    
                    # Only add non-zero rows
                    if row1[debit_col] != 0 or row1[credit_col] != 0:
                        result_rows.append(row1)
                    if row2[debit_col] != 0 or row2[credit_col] != 0:
                        result_rows.append(row2)
                else:
                    # Normal IC opposite logic
                    # IMPORTANT: If ic_internal is 998, replace it with 783 for elimination
                    # 998 should NEVER appear in output
                    effective_ic_internal = ic_internal
                    if ic_internal in [998, '998']:
                        effective_ic_internal = 783
                    
                    opposite_info = IC_OPPOSITE.get(effective_ic_internal) or IC_OPPOSITE.get(str(effective_ic_internal))
                    
                    if opposite_info:
                        # For elimination: swap debit/credit on the IC accounts only
                        # Calculate NET amount to avoid having both debit and credit on same line
                        net_amount = ic_debit - ic_credit  # Positive = net debit, Negative = net credit
                        
                        # Create 2 lines - one for IC account, one for opposite IC
                        # Each line should have ONLY debit OR credit, not both
                        # Use effective_ic_internal (which replaces 998 with 783)
                        row1 = template.copy()
                        row1[internal_col] = effective_ic_internal
                        if account_col:
                            # Map internal to account number - 998 is already replaced with 783
                            IC_ACCOUNT_MAP = {111: 210000, 119: 120003, 783: 121001, 964: 121002}
                            ic_account = IC_ACCOUNT_MAP.get(effective_ic_internal) or IC_ACCOUNT_MAP.get(int(effective_ic_internal) if str(effective_ic_internal).isdigit() else effective_ic_internal)
                            if ic_account:
                                row1[account_col] = ic_account
                        
                        # Swap: if original was net debit, elimination is net credit (and vice versa)
                        if net_amount >= 0:
                            # Original was net debit -> Elimination: IC gets credit, Opposite gets debit
                            row1[debit_col] = 0
                            row1[credit_col] = net_amount
                        else:
                            # Original was net credit -> Elimination: IC gets debit, Opposite gets credit
                            row1[debit_col] = abs(net_amount)
                            row1[credit_col] = 0
                        
                        # Line 2: Opposite IC account (to balance the elimination)
                        row2 = template.copy()
                        row2[internal_col] = opposite_info['internal']
                        if account_col:
                            row2[account_col] = opposite_info['account']
                        
                        # Opposite gets the reverse of row1
                        if net_amount >= 0:
                            # Row1 has credit, Row2 gets debit
                            row2[debit_col] = net_amount
                            row2[credit_col] = 0
                        else:
                            # Row1 has debit, Row2 gets credit
                            row2[debit_col] = 0
                            row2[credit_col] = abs(net_amount)
                        
                        # Only add non-zero rows
                        if row1[debit_col] != 0 or row1[credit_col] != 0:
                            result_rows.append(row1)
                        if row2[debit_col] != 0 or row2[credit_col] != 0:
                            result_rows.append(row2)
                    else:
                        # Fallback: just swap all rows
                        print(f"  Warning: REF {ref_val} - no IC opposite found, swapping all rows")
                        for _, row in batch.iterrows():
                            new_row = row.to_dict()
                            orig_debit = new_row[debit_col]
                            new_row[debit_col] = new_row[credit_col]
                            new_row[credit_col] = orig_debit
                            if new_row[debit_col] != 0 or new_row[credit_col] != 0:
                                result_rows.append(new_row)
            else:
                # No IC account found in batch with bank - fallback
                print(f"  Warning: REF {ref_val} has bank but no IC account, swapping all rows")
                for _, row in batch.iterrows():
                    new_row = row.to_dict()
                    orig_debit = new_row[debit_col]
                    new_row[debit_col] = new_row[credit_col]
                    new_row[credit_col] = orig_debit
                    if new_row[debit_col] != 0 or new_row[credit_col] != 0:
                        result_rows.append(new_row)
        else:
            regular_batches += 1
            # === REGULAR BATCH: Swap debit/credit for all rows ===
            for _, row in batch.iterrows():
                new_row = row.to_dict()
                # Store original internal for opposite mapping
                orig_internal = new_row[internal_col]
                
                # Swap debit/credit
                orig_debit = new_row[debit_col]
                new_row[debit_col] = new_row[credit_col]
                new_row[credit_col] = orig_debit
                
                # SPECIAL CASE: 998 should NEVER appear in output - always replace with 783
                try:
                    int_orig = int(orig_internal)
                except (ValueError, TypeError):
                    int_orig = orig_internal
                
                if int_orig in [998, '998']:
                    new_row[internal_col] = 783
                    if account_col:
                        new_row[account_col] = 121001
                else:
                    # Apply opposite mapping if exists (from mapping file)
                    if orig_internal in internal_to_opposite_internal:
                        new_row[internal_col] = internal_to_opposite_internal[orig_internal]
                    elif int_orig in internal_to_opposite_internal:
                        new_row[internal_col] = internal_to_opposite_internal[int_orig]
                    
                    # Apply opposite account if exists
                    if account_col:
                        if orig_internal in internal_to_opposite_account:
                            new_row[account_col] = internal_to_opposite_account[orig_internal]
                        elif int_orig in internal_to_opposite_account:
                            new_row[account_col] = internal_to_opposite_account[int_orig]
                
                # Only add non-zero rows
                if new_row[debit_col] != 0 or new_row[credit_col] != 0:
                    result_rows.append(new_row)
    
    print(f"\nProcessed {bank_batches} batches with Bank accounts (consolidated to 2 lines)")
    print(f"Processed {regular_batches} regular batches (swapped all rows)")
    
    # Validate balance per REF batch
    if len(result_rows) > 0:
        temp_df = pd.DataFrame(result_rows)
        unbalanced_refs = []
        for ref_val in temp_df[ref_col].unique():
            ref_rows = temp_df[temp_df[ref_col] == ref_val]
            ref_debit = ref_rows[debit_col].sum()
            ref_credit = ref_rows[credit_col].sum()
            if abs(ref_debit - ref_credit) > 0.01:
                unbalanced_refs.append((ref_val, ref_debit, ref_credit))
        
        if unbalanced_refs:
            print(f"\n⚠ WARNING: {len(unbalanced_refs)} unbalanced batches detected:")
            for ref_val, d, c in unbalanced_refs[:5]:  # Show first 5
                print(f"    REF {ref_val}: Debit={d:.2f}, Credit={c:.2f}, Diff={abs(d-c):.2f}")
        else:
            print(f"\n✓ All {len(temp_df[ref_col].unique())} batches are balanced")
    
    if len(result_rows) == 0:
        print("No elimination rows created!")
        return pd.DataFrame()
    
    df_elim = pd.DataFrame(result_rows)
    print(f"Created {len(df_elim)} elimination rows")
    
    # Update Account Name based on new internal
    if account_name_col and account_name_col in df_elim.columns:
        def get_account_name(internal_val):
            if internal_val in internal_to_account_name:
                return internal_to_account_name[internal_val]
            try:
                int_val = int(internal_val)
                if int_val in internal_to_account_name:
                    return internal_to_account_name[int_val]
            except (ValueError, TypeError):
                pass
            return ''
        
        df_elim[account_name_col] = df_elim[internal_col].apply(get_account_name)
    
    # Set subsidiary for elimination
    if 'Subsidiary Header' in df_elim.columns:
        df_elim['Subsidiary Header'] = 5
    if 'Subsidiary Line' in df_elim.columns:
        df_elim['Subsidiary Line'] = 5
    if 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY' in df_elim.columns:
        df_elim['Subsidiary Line/ DUE TO/ FROM SUBSIDIARY'] = 6
    
    # Clear Name
    if 'Name' in df_elim.columns:
        df_elim['Name'] = ''
    
    # Remove helper columns
    helper_cols = ['_is_lsport', '_in_lsport_batch', '_is_bank', '_is_ic']
    df_elim = df_elim.drop(columns=[c for c in helper_cols if c in df_elim.columns])
    
    # Add placeholder EXTERNAL ID (will be updated per-month in save_by_month)
    df_elim.insert(0, 'EXTERNAL ID', 'JE_ELIM')
    
    # Sort by REF
    if ref_col in df_elim.columns:
        df_elim = df_elim.sort_values(by=ref_col).reset_index(drop=True)
    
    # Show summary
    total_debit = df_elim[debit_col].sum()
    total_credit = df_elim[credit_col].sum()
    
    print(f"\n{'='*50}")
    print(f"ELIMINATION JE SUMMARY")
    print(f"{'='*50}")
    print(f"Total rows: {len(df_elim)}")
    print(f"Total Debit:  {total_debit:,.2f}")
    print(f"Total Credit: {total_credit:,.2f}")
    if abs(total_debit - total_credit) < 0.01:
        print("✓ BALANCED")
    else:
        print(f"⚠ DIFFERENCE: {abs(total_debit - total_credit):,.2f}")
    
    return df_elim


def get_output_folder():
    """Get Statscore Import folder on desktop"""
    desktop = os.path.join(os.path.expanduser("~"), "Desktop")
    folder = os.path.join(desktop, "Statscore Import")
    if not os.path.exists(folder):
        os.makedirs(folder)
    return folder


def save_csv(df, path):
    """Save DataFrame to CSV"""
    df.to_csv(path, index=False, encoding='utf-8-sig')
    print(f"Saved {len(df)} rows to {path}")


def get_month_filename(posting_period, prefix="JE_IC_Elimination"):
    """
    Convert posting period to filename format.
    E.g., "Jan 2025" -> "JE_IC_Elimination_1.25", "Feb 2025" -> "JE_IC_Elimination_2.25"
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


def save_by_month(df, prefix):
    """
    Split DataFrame by month and save each month to a separate file.
    Uses Posting Period column if available, otherwise uses Date column.
    
    Each file gets:
    - Separate file per month
    - Updated EXTERNAL ID for that month (e.g., JE_ELIM0125 for Jan 2025)
    - Filename format: prefix_M.YY.csv (e.g., JE_IC_Elimination_1.25.csv)
    """
    folder = get_output_folder()
    
    MONTH_MAP = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    # Find posting period column
    pp_col = None
    for col in df.columns:
        if 'posting period' in str(col).lower():
            pp_col = col
            break
    
    # Find date column
    date_col = None
    for col in df.columns:
        if str(col).strip().lower() == 'date':
            date_col = col
            break
    
    # Find debit/credit columns for balance summary
    debit_col = None
    credit_col = None
    for col in df.columns:
        col_lower = str(col).lower()
        if 'debit' in col_lower:
            debit_col = col
        elif 'credit' in col_lower:
            credit_col = col
    
    # Determine which column to use for splitting
    use_date_column = False
    if pp_col is not None:
        print(f"\n  Splitting by Posting Period column: '{pp_col}'")
    elif date_col is not None:
        print(f"\n  Posting Period column not found. Using Date column: '{date_col}'")
        use_date_column = True
    else:
        print(f"  Warning: Could not find 'Posting Period' or 'Date' column. Saving all to single file.")
        path = os.path.join(folder, f"{prefix}.csv")
        save_csv(df, path)
        return [(None, path, len(df))], folder
    
    # Create a working copy
    df_work = df.copy()
    
    if use_date_column:
        # Parse dates and extract month/year
        df_work['_parsed_date'] = pd.to_datetime(df_work[date_col], dayfirst=True, errors='coerce')
        df_work['_month'] = df_work['_parsed_date'].dt.month
        df_work['_year'] = df_work['_parsed_date'].dt.year
        df_work['_year_short'] = df_work['_year'] % 100
        
        # Create a period key like "1.25" for Jan 2025
        df_work['_period_key'] = df_work.apply(
            lambda row: f"{int(row['_month'])}.{int(row['_year_short']):02d}" 
            if pd.notna(row['_month']) else None, 
            axis=1
        )
        
        # Get unique periods
        all_periods = df_work['_period_key'].dropna().unique()
        print(f"  Found {len(all_periods)} unique months")
        print(f"  Saving to: {folder}")
        
        saved = []
        for period_key in sorted(all_periods, key=lambda x: (int(x.split('.')[1]), int(x.split('.')[0]))):
            # Get rows for this period
            period_df = df_work[df_work['_period_key'] == period_key].copy()
            
            if len(period_df) == 0:
                continue
            
            # Extract month and year from period_key (e.g., "1.25")
            parts = period_key.split('.')
            month_num = int(parts[0])
            year_short = parts[1]
            
            # Update EXTERNAL ID for this specific month (e.g., JE_ELIM0125)
            if 'EXTERNAL ID' in period_df.columns:
                external_id = f'JE_ELIM{month_num:02d}{year_short}'
                period_df['EXTERNAL ID'] = external_id
            
            # Remove helper columns
            helper_cols = ['_parsed_date', '_month', '_year', '_year_short', '_period_key']
            period_df = period_df.drop(columns=[c for c in helper_cols if c in period_df.columns])
            
            # Generate filename
            filename = f"{prefix}_{month_num}.{year_short}"
            path = os.path.join(folder, f"{filename}.csv")
            save_csv(period_df, path)
            
            # Calculate balance for this period
            period_debit = period_df[debit_col].sum() if debit_col else 0
            period_credit = period_df[credit_col].sum() if credit_col else 0
            balance_status = "✓" if abs(period_debit - period_credit) < 0.01 else "⚠"
            
            # Get month name for display
            month_names = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                          'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
            period_name = f"{month_names[month_num]} 20{year_short}"
            
            print(f"    {period_name}: {len(period_df)} rows, Debit={period_debit:,.2f}, Credit={period_credit:,.2f} {balance_status}")
            
            saved.append((period_name, path, len(period_df)))
        
        return saved, folder
    
    else:
        # Use Posting Period column (original logic)
        all_periods = df_work[pp_col].dropna().unique()
        print(f"  Found {len(all_periods)} unique posting periods")
        print(f"  Saving to: {folder}")
        
        saved = []
        for period in sorted(all_periods):
            # Get rows for this period
            period_df = df_work[df_work[pp_col] == period].copy()
            
            if len(period_df) == 0:
                continue
            
            # Parse month/year from posting period for EXTERNAL ID
            try:
                parts = str(period).strip().split()
                month_str = parts[0].lower()[:3]
                year_str = parts[-1][-2:]
                month_num = MONTH_MAP.get(month_str, 0)
                
                # Update EXTERNAL ID for this specific month (e.g., JE_ELIM0125)
                if 'EXTERNAL ID' in period_df.columns and month_num > 0:
                    external_id = f'JE_ELIM{month_num:02d}{year_str}'
                    period_df['EXTERNAL ID'] = external_id
            except:
                month_num = 0
                year_str = "00"
            
            # Get the filename for this month
            filename = get_month_filename(period, prefix=prefix)
            
            if filename is None:
                # Fallback: sanitize the posting period for filename
                safe_period = str(period).replace('/', '-').replace('\\', '-').replace(' ', '_')
                filename = f"{prefix}_{safe_period}"
            
            path = os.path.join(folder, f"{filename}.csv")
            save_csv(period_df, path)
            
            # Calculate balance for this period
            period_debit = period_df[debit_col].sum() if debit_col else 0
            period_credit = period_df[credit_col].sum() if credit_col else 0
            balance_status = "✓" if abs(period_debit - period_credit) < 0.01 else "⚠"
            
            print(f"    {period}: {len(period_df)} rows, Debit={period_debit:,.2f}, Credit={period_credit:,.2f} {balance_status}")
            
            saved.append((period, path, len(period_df)))
        
        return saved, folder


def main():
    print("=" * 60)
    print("IC ELIMINATION JOURNAL ENTRY CREATOR")
    print("=" * 60)
    print("Creates elimination entries for IC transactions")
    print("Only includes accounts with Opposite mapping defined")
    print()
    
    csv_file, mapping_file = select_files()
    
    if not csv_file or not mapping_file:
        print("Cancelled.")
        return
    
    print(f"CSV: {csv_file}")
    print(f"Mapping: {mapping_file}")
    
    try:
        print("\n--- Loading Mapping ---")
        mapping_data = load_mapping(mapping_file)
        
        print("\n--- Creating Elimination JE ---")
        df_elim = create_elimination_je(csv_file, mapping_data)
        
        if len(df_elim) == 0:
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning("No Data", "No eligible IC transactions found.")
            root.destroy()
            return
        
        print("\n--- Saving ---")
        files, folder = save_by_month(df_elim, "JE_IC_Elimination")
        
        # Build message
        msg = ""
        if isinstance(files[0], tuple):
            for period, path, count in files:
                msg += f"  {os.path.basename(path)} ({count} rows)\n"
        else:
            msg = f"  {os.path.basename(files[0])} ({len(df_elim)} rows)\n"
        
        # Get totals
        debit_col = credit_col = None
        for col in df_elim.columns:
            if 'debit' in str(col).lower():
                debit_col = col
            elif 'credit' in str(col).lower():
                credit_col = col
        
        total_debit = pd.to_numeric(df_elim[debit_col], errors='coerce').fillna(0).sum() if debit_col else 0
        total_credit = pd.to_numeric(df_elim[credit_col], errors='coerce').fillna(0).sum() if credit_col else 0
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", 
            f"IC Elimination JE Created!\n\n"
            f"Saved to:\n{folder}\n\n"
            f"Files:\n{msg}\n"
            f"Total Debit: {total_debit:,.2f}\n"
            f"Total Credit: {total_credit:,.2f}")
        root.destroy()
        
        print("\nDone!")
        
    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"Error:\n{e}")
        root.destroy()


if __name__ == "__main__":
    main()


