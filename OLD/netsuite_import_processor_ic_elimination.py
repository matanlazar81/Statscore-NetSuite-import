"""
NetSuite IC Elimination Journal Entry Creator
- Creates elimination JE for transactions with LSPORT/LSPORTS in "Nazwa" column
- If any row in a REF batch has LSPORT account, includes ALL rows from that batch
- Swaps Debit/Credit to create reversal entries
- Output is a balanced elimination journal entry
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
    root.title("IC Elimination JE Creator")
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
    title_label = ttk.Label(main_frame, text="IC Elimination JE Creator", font=('Helvetica', 14, 'bold'))
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
    ttk.Button(button_frame, text="Create Elimination JE", command=process, width=20).pack(side=tk.LEFT, padx=10)
    ttk.Button(button_frame, text="Cancel", command=cancel, width=15).pack(side=tk.LEFT, padx=10)
    
    root.mainloop()
    
    return result['csv'], result['mapping']


def load_mapping(mapping_file):
    """
    Load the mapping file and find accounts where "Nazwa" contains LSPORT/LSPORTS.
    Also identify Bank type accounts for special handling.
    """
    df_mapping = pd.read_excel(mapping_file)
    
    # Find column names
    internal_col = None
    ls_account_col = None
    intercompany_col = None
    account_name_col = None
    nazwa_col = None  # Key column for LSPORT filtering
    type_col = None   # Key column for Bank type identification
    
    for col in df_mapping.columns:
        col_clean = str(col).strip().lower()
        if 'internal' in col_clean and 'opposite' not in col_clean and internal_col is None:
            internal_col = col
        elif col_clean == 'ls account' or col_clean == 'lsaccount':
            ls_account_col = col
        elif 'intercompany' in col_clean:
            intercompany_col = col
        elif 'account name' in col_clean:
            account_name_col = col
        elif col_clean == 'nazwa':
            nazwa_col = col
        elif col_clean == 'type' or col_clean == 'ls account type' or 'account type' in col_clean:
            type_col = col
    
    if internal_col is None:
        raise ValueError("Could not find 'Internal' column in mapping file")
    if nazwa_col is None:
        raise ValueError("Could not find 'Nazwa' column in mapping file")
    
    print(f"Using columns: Internal='{internal_col}', Nazwa='{nazwa_col}'")
    if type_col:
        print(f"  Type='{type_col}'")
    if ls_account_col:
        print(f"  LS Account='{ls_account_col}'")
    if intercompany_col:
        print(f"  Intercompany='{intercompany_col}'")
    
    # Create mappings
    lsport_internals = set()  # Internal IDs where Nazwa contains LSPORT
    bank_internals = set()    # Internal IDs where Type is Bank
    ic_mapping = {}
    internal_to_account_name = {}
    internal_to_ls_account = {}
    
    for _, row in df_mapping.iterrows():
        internal_val = row[internal_col]
        
        # Check if Nazwa contains LSPORT or LSPORTS (case insensitive)
        nazwa_val = str(row[nazwa_col]).strip().upper() if pd.notna(row[nazwa_col]) else ""
        is_lsport = 'LSPORT' in nazwa_val  # This catches both LSPORT and LSPORTS
        
        if is_lsport:
            # Add to LSPORT internals set
            lsport_internals.add(internal_val)
            try:
                lsport_internals.add(int(internal_val))
                lsport_internals.add(str(int(internal_val)))
            except (ValueError, TypeError):
                pass
            print(f"    LSPORT account: Internal {internal_val} - {row[nazwa_col]}")
        
        # Check if Type is Bank
        if type_col:
            type_val = str(row[type_col]).strip().lower() if pd.notna(row[type_col]) else ""
            if type_val == 'bank':
                bank_internals.add(internal_val)
                try:
                    bank_internals.add(int(internal_val))
                    bank_internals.add(str(int(internal_val)))
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
        
        # IC mapping (for reference)
        if intercompany_col:
            intercompany_val = str(row[intercompany_col]).strip().lower()
            ic_mapping[internal_val] = intercompany_val == 'yes'
            try:
                int_val = int(internal_val)
                ic_mapping[int_val] = intercompany_val == 'yes'
                ic_mapping[str(int_val)] = intercompany_val == 'yes'
            except (ValueError, TypeError):
                pass
        
        # Account name mapping
        if account_name_col:
            account_name = row[account_name_col]
            if pd.notna(account_name):
                internal_to_account_name[internal_val] = account_name
                try:
                    internal_to_account_name[int(internal_val)] = account_name
                    internal_to_account_name[str(int(internal_val))] = account_name
                except (ValueError, TypeError):
                    pass
    
    lsport_count = len([x for x in lsport_internals if isinstance(x, int) or (isinstance(x, str) and x.isdigit())])
    bank_count = len([x for x in bank_internals if isinstance(x, int) or (isinstance(x, str) and x.isdigit())])
    print(f"\n  Found {lsport_count} accounts with LSPORT in Nazwa")
    print(f"  Found {bank_count} Bank type accounts")
    print(f"  REF batches containing LSPORT accounts will be included in elimination JE")
    print(f"  Bank accounts will be replaced with opposite IC account (111 <-> 119)")
    
    return {
        'lsport_internals': lsport_internals,
        'bank_internals': bank_internals,
        'ic_mapping': ic_mapping,
        'internal_to_account_name': internal_to_account_name,
        'internal_to_ls_account': internal_to_ls_account
    }


def create_elimination_je(csv_file, mapping_data):
    """
    Create an elimination journal entry to reverse LSPORT transactions.
    Only includes REF batches that contain at least one account with "LSPORT" in Nazwa.
    Bank accounts are replaced with opposite IC account (111 <-> 119).
    Each REF batch results in exactly 2 lines.
    Returns a balanced DataFrame with swapped debit/credit.
    """
    # Read CSV
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(csv_file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='latin-1')
    
    print(f"Loaded {len(df)} rows from CSV")
    
    lsport_internals = mapping_data['lsport_internals']
    bank_internals = mapping_data['bank_internals']
    internal_to_ls_account = mapping_data.get('internal_to_ls_account', {})
    internal_to_account_name = mapping_data.get('internal_to_account_name', {})
    
    # IC account mapping: 111 (210000) <-> 119 (120003)
    IC_OPPOSITE = {
        111: {'internal': 119, 'account': 120003},
        119: {'internal': 111, 'account': 210000},
        '111': {'internal': 119, 'account': 120003},
        '119': {'internal': 111, 'account': 210000},
    }
    
    # Find required columns
    internal_col = None
    account_col = None
    ref_col = None
    date_col = None
    debit_col = None
    credit_col = None
    
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if col_clean == 'internal' or col_clean == 'internal id':
            internal_col = col
        elif col_clean == 'account':
            account_col = col
        elif col_clean == 'ref':
            ref_col = col
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
    
    print(f"Using columns: Internal='{internal_col}', REF='{ref_col}'")
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
    
    # Mark LSPORT and Bank rows
    df['_is_lsport'] = df[internal_col].apply(is_lsport_account)
    df['_is_bank'] = df[internal_col].apply(is_bank_account)
    
    lsport_row_count = df['_is_lsport'].sum()
    bank_row_count = df['_is_bank'].sum()
    print(f"Found {lsport_row_count} rows with LSPORT accounts")
    print(f"Found {bank_row_count} rows with Bank accounts")
    
    # Find all REF batches that contain at least one LSPORT account
    lsport_refs = df[df['_is_lsport']][ref_col].unique()
    print(f"Found {len(lsport_refs)} REF batches containing LSPORT accounts")
    
    # Include ALL rows from those REF batches
    df['_in_lsport_batch'] = df[ref_col].isin(lsport_refs)
    df_ic = df[df['_in_lsport_batch']].copy()
    
    print(f"Total rows in LSPORT batches: {len(df_ic)}")
    
    if len(df_ic) == 0:
        print("No IC transactions found!")
        return pd.DataFrame()
    
    # === PROCESS EACH REF BATCH ===
    # For each batch: consolidate to 2 lines, replace bank with opposite IC account
    result_rows = []
    
    for ref_val in lsport_refs:
        batch = df_ic[df_ic[ref_col] == ref_val].copy()
        
        # Get debit and credit totals
        batch[debit_col] = pd.to_numeric(batch[debit_col], errors='coerce').fillna(0)
        batch[credit_col] = pd.to_numeric(batch[credit_col], errors='coerce').fillna(0)
        
        total_debit = batch[debit_col].sum()
        total_credit = batch[credit_col].sum()
        
        # Find the non-bank IC account in this batch (111 or 119)
        non_bank_rows = batch[~batch['_is_bank']]
        bank_rows = batch[batch['_is_bank']]
        
        # Get template row (first row of batch for metadata)
        template = batch.iloc[0].copy()
        
        if len(bank_rows) > 0 and len(non_bank_rows) > 0:
            # Batch has bank account - replace with opposite IC
            # Find the IC account (111 or 119) in non-bank rows
            ic_internal = None
            for _, row in non_bank_rows.iterrows():
                int_val = row[internal_col]
                try:
                    int_val = int(int_val)
                except:
                    pass
                if int_val in IC_OPPOSITE or str(int_val) in IC_OPPOSITE:
                    ic_internal = int_val
                    break
            
            if ic_internal is not None:
                # Get the opposite IC account
                opposite = IC_OPPOSITE.get(ic_internal) or IC_OPPOSITE.get(str(ic_internal))
                
                # Create 2 lines: one with original IC, one with opposite IC
                # Line 1: Original IC account (non-bank)
                row1 = template.copy()
                row1[internal_col] = ic_internal
                if account_col:
                    row1[account_col] = internal_to_ls_account.get(ic_internal, row1.get(account_col, ''))
                
                # Line 2: Opposite IC account (replaces bank)
                row2 = template.copy()
                row2[internal_col] = opposite['internal']
                if account_col:
                    row2[account_col] = opposite['account']
                
                # Assign debit/credit based on original IC position
                # If original IC had credit, it gets debit in elimination (swap)
                ic_row = non_bank_rows.iloc[0]
                ic_had_credit = ic_row[credit_col] > 0
                
                if ic_had_credit:
                    # Original: IC=Credit, Bank=Debit
                    # Elimination (swapped): IC=Debit, Opposite=Credit
                    row1[debit_col] = total_credit
                    row1[credit_col] = 0
                    row2[debit_col] = 0
                    row2[credit_col] = total_debit
                else:
                    # Original: IC=Debit, Bank=Credit
                    # Elimination (swapped): IC=Credit, Opposite=Debit
                    row1[debit_col] = 0
                    row1[credit_col] = total_debit
                    row2[debit_col] = total_credit
                    row2[credit_col] = 0
                
                result_rows.append(row1)
                result_rows.append(row2)
            else:
                # No IC account found, just swap and include non-bank rows
                for _, row in non_bank_rows.iterrows():
                    new_row = row.copy()
                    orig_debit = new_row[debit_col]
                    new_row[debit_col] = new_row[credit_col]
                    new_row[credit_col] = orig_debit
                    result_rows.append(new_row)
        else:
            # No bank account in batch - include all rows with swapped debit/credit
            for _, row in batch.iterrows():
                new_row = row.copy()
                orig_debit = new_row[debit_col]
                new_row[debit_col] = new_row[credit_col]
                new_row[credit_col] = orig_debit
                result_rows.append(new_row)
    
    if len(result_rows) == 0:
        print("No elimination rows created!")
        return pd.DataFrame()
    
    df_elim = pd.DataFrame(result_rows)
    print(f"Created {len(df_elim)} elimination rows")
    
    # Update subsidiary for elimination
    if 'Subsidiary Line' in df_elim.columns:
        df_elim['Subsidiary Line'] = 5  # xElimination - Parent
    
    if 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY' in df_elim.columns:
        df_elim['Subsidiary Line/ DUE TO/ FROM SUBSIDIARY'] = 6
    
    # Set Subsidiary Header to 5 for elimination
    if 'Subsidiary Header' in df_elim.columns:
        df_elim['Subsidiary Header'] = 5
    
    # Clear Name for elimination
    if 'Name' in df_elim.columns:
        df_elim['Name'] = ''
    
    # Remove helper columns
    helper_cols = ['_is_lsport', '_in_lsport_batch', '_is_bank']
    df_elim = df_elim.drop(columns=[c for c in helper_cols if c in df_elim.columns])
    
    # Add EXTERNAL ID
    if date_col and date_col in df_elim.columns:
        try:
            dates = pd.to_datetime(df_elim[date_col], dayfirst=True, errors='coerce')
            newest_date = dates.max()
            if pd.notna(newest_date):
                month_num = int(newest_date.month)
                year_short = int(newest_date.year) % 100
                external_id = f'JE_ELIM{month_num:02d}{year_short:02d}'
                df_elim.insert(0, 'EXTERNAL ID', external_id)
        except:
            df_elim.insert(0, 'EXTERNAL ID', 'JE_ELIM')
    else:
        df_elim.insert(0, 'EXTERNAL ID', 'JE_ELIM')
    
    # Filter out zero rows
    debit_vals = pd.to_numeric(df_elim[debit_col], errors='coerce').fillna(0)
    credit_vals = pd.to_numeric(df_elim[credit_col], errors='coerce').fillna(0)
    rows_before = len(df_elim)
    df_elim = df_elim[(debit_vals != 0) | (credit_vals != 0)]
    if rows_before - len(df_elim) > 0:
        print(f"  Filtered out {rows_before - len(df_elim)} zero rows")
    
    # Sort by REF
    if ref_col in df_elim.columns:
        df_elim = df_elim.sort_values(by=ref_col).reset_index(drop=True)
    
    # Verify balance
    total_debit = pd.to_numeric(df_elim[debit_col], errors='coerce').fillna(0).sum()
    total_credit = pd.to_numeric(df_elim[credit_col], errors='coerce').fillna(0).sum()
    
    print(f"\n=== ELIMINATION JE SUMMARY ===")
    print(f"Total rows: {len(df_elim)}")
    print(f"Total Debit:  {total_debit:,.2f}")
    print(f"Total Credit: {total_credit:,.2f}")
    
    if abs(total_debit - total_credit) < 0.01:
        print("✓ Journal Entry is BALANCED")
    else:
        print(f"⚠ Difference: {abs(total_debit - total_credit):,.2f}")
    
    return df_elim


def get_statscore_import_folder():
    """Create and return the path to Statscore Import folder on desktop"""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    statscore_folder = os.path.join(desktop_path, "Statscore Import")
    
    if not os.path.exists(statscore_folder):
        os.makedirs(statscore_folder)
        print(f"Created folder: {statscore_folder}")
    
    return statscore_folder


def save_csv(df, output_path):
    """Save DataFrame to CSV"""
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved {len(df)} rows to {output_path}")


def get_month_filename(posting_period, prefix="JE_IC_Elimination"):
    """Convert posting period to filename format"""
    MONTH_TO_NUM = {
        'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
        'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12
    }
    
    try:
        parts = str(posting_period).strip().split()
        if len(parts) >= 2:
            month_str = parts[0].lower()[:3]
            year_str = parts[-1]
            month_num = MONTH_TO_NUM.get(month_str)
            if month_num:
                year_short = year_str[-2:]
                return f"{prefix}_{month_num}.{year_short}"
    except:
        pass
    return None


def save_by_month(df, file_prefix):
    """Split and save by posting period"""
    saved_files = []
    statscore_folder = get_statscore_import_folder()
    
    # Find Posting Period column
    posting_period_col = None
    for col in df.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    if posting_period_col is None:
        # Save as single file
        output_path = os.path.join(statscore_folder, f"{file_prefix}.csv")
        save_csv(df, output_path)
        return [output_path], statscore_folder
    
    # Group by posting period
    grouped = df.groupby(posting_period_col)
    
    print(f"\nSplitting by Posting Period...")
    print(f"Saving to: {statscore_folder}")
    
    for posting_period, group_df in grouped:
        group_df = group_df.copy().reset_index(drop=True)
        
        filename = get_month_filename(posting_period, prefix=file_prefix)
        if filename is None:
            safe_period = str(posting_period).replace('/', '-').replace('\\', '-').replace(' ', '_')
            filename = f"{file_prefix}_{safe_period}"
        
        output_path = os.path.join(statscore_folder, f"{filename}.csv")
        save_csv(group_df, output_path)
        saved_files.append((posting_period, output_path, len(group_df)))
    
    return saved_files, statscore_folder


def main():
    print("=" * 60)
    print("IC ELIMINATION JOURNAL ENTRY CREATOR")
    print("=" * 60)
    print("Creates elimination entries for LSPORT transactions")
    print("(Filters by 'Nazwa' column containing LSPORT/LSPORTS)")
    print("(Includes entire REF batch if any row has LSPORT account)")
    print()
    
    # Select files
    csv_file, mapping_file = select_files()
    
    if not csv_file or not mapping_file:
        print("Cancelled.")
        return
    
    print(f"CSV: {csv_file}")
    print(f"Mapping: {mapping_file}")
    
    try:
        # Load mapping
        print("\nLoading mapping...")
        mapping_dict = load_mapping(mapping_file)
        
        # Create elimination JE
        print("\nCreating elimination JE...")
        df_elim = create_elimination_je(csv_file, mapping_dict)
        
        if len(df_elim) == 0:
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning("No Data", "No LSPORT transactions found for elimination.\n\nMake sure the mapping file has a 'Nazwa' column\nwith accounts containing 'LSPORT' or 'LSPORTS'.")
            root.destroy()
            return
        
        # Save
        print("\nSaving...")
        elim_files, statscore_folder = save_by_month(df_elim, "JE_IC_Elimination")
        
        # Build message
        files_msg = ""
        if isinstance(elim_files[0], tuple):
            for period, path, count in elim_files:
                files_msg += f"  {os.path.basename(path)} ({count} rows)\n"
        else:
            files_msg = f"  {os.path.basename(elim_files[0])} ({len(df_elim)} rows)\n"
        
        # Calculate totals for message
        debit_col = None
        credit_col = None
        for col in df_elim.columns:
            if 'debit' in str(col).lower():
                debit_col = col
            elif 'credit' in str(col).lower():
                credit_col = col
        
        total_debit = pd.to_numeric(df_elim[debit_col], errors='coerce').fillna(0).sum() if debit_col else 0
        total_credit = pd.to_numeric(df_elim[credit_col], errors='coerce').fillna(0).sum() if credit_col else 0
        
        # Show success
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", 
            f"IC Elimination JE Created!\n\n"
            f"Saved to:\n{statscore_folder}\n\n"
            f"Files:\n{files_msg}\n"
            f"Total Debit: {total_debit:,.2f}\n"
            f"Total Credit: {total_credit:,.2f}\n\n"
            f"{'✓ BALANCED' if abs(total_debit - total_credit) < 0.01 else '⚠ NOT BALANCED'}")
        root.destroy()
        
        print("\nDone!")
        
    except Exception as e:
        print(f"\nError: {e}")
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Error", f"Error:\n{e}")
        root.destroy()


if __name__ == "__main__":
    main()

