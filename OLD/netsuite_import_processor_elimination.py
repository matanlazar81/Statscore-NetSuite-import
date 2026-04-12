"""
NetSuite Import Processor - IC Elimination Only
- Opens a GUI to select CSV file and Excel mapping file
- Identifies Intercompany (IC) transactions using the mapping file
- Creates ONLY the opposite account entries for elimination purposes
- Output contains ONLY the elimination entries (opposite accounts)
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
    root.title("NetSuite IC Elimination Processor")
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
    title_label = ttk.Label(main_frame, text="NetSuite IC Elimination Processor", font=('Helvetica', 14, 'bold'))
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
    - Internal -> Opposite Internal
    - Internal -> Opposite Account
    - Internal -> Account Name
    - Internal -> Account Type (for filtering Expenses/Income)
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
        if 'internal' in col_clean and 'opposite' not in col_clean and 'oposite' not in col_clean and internal_col is None:
            internal_col = col
        elif col_clean == 'intercompany' or 'intercompany' in col_clean:
            intercompany_col = col
        elif col_clean == 'ls account' or col_clean == 'lsaccount':
            ls_account_col = col
        elif ('oposite' in col_clean or 'opposite' in col_clean) and 'internal' in col_clean:
            opposite_internal_col = col
        elif ('oposite' in col_clean or 'opposite' in col_clean) and 'account' in col_clean:
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
    ic_mapping = {}
    internal_to_opposite_account = {}
    internal_to_opposite_internal = {}
    internal_to_account_name = {}
    internal_to_ls_account = {}
    internal_to_account_type = {}  # For filtering by Expenses/Income
    
    for _, row in df_mapping.iterrows():
        internal_val = row[internal_col]
        intercompany_val = str(row[intercompany_col]).strip().lower()
        ic_mapping[internal_val] = intercompany_val == 'yes'
        
        # Build Internal -> Account Type mapping
        if account_type_col:
            account_type = row[account_type_col]
            if pd.notna(account_type):
                internal_to_account_type[internal_val] = str(account_type).strip()
                try:
                    internal_to_account_type[int(internal_val)] = str(account_type).strip()
                    internal_to_account_type[str(int(internal_val))] = str(account_type).strip()
                except (ValueError, TypeError):
                    pass
        
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
        
        # Build Internal -> Opposite Account mapping
        if opposite_account_col:
            opposite_account = row[opposite_account_col]
            if pd.notna(opposite_account):
                internal_to_opposite_account[internal_val] = opposite_account
                try:
                    internal_to_opposite_account[int(internal_val)] = opposite_account
                    internal_to_opposite_account[str(int(internal_val))] = opposite_account
                except (ValueError, TypeError):
                    pass
    
    # Build Internal -> Opposite Internal mapping
    if opposite_internal_col:
        print(f"  Using 'Opposite Internal' column for direct mapping")
        for _, row in df_mapping.iterrows():
            internal_val = row[internal_col]
            opposite_internal_val = row[opposite_internal_col]
            if pd.notna(opposite_internal_val):
                internal_to_opposite_internal[internal_val] = opposite_internal_val
                try:
                    int_internal = int(internal_val)
                    int_opp_internal = int(float(opposite_internal_val))
                    internal_to_opposite_internal[int_internal] = int_opp_internal
                    internal_to_opposite_internal[str(int_internal)] = int_opp_internal
                except (ValueError, TypeError):
                    pass
    
    print(f"  IC accounts found: {sum(1 for v in ic_mapping.values() if v)}")
    print(f"  Opposite internal mappings: {len(internal_to_opposite_internal) // 3 if internal_to_opposite_internal else 0}")
    
    # Count expense and income accounts
    expense_count = sum(1 for v in internal_to_account_type.values() if v.lower() in ['expense', 'expenses'])
    income_count = sum(1 for v in internal_to_account_type.values() if v.lower() == 'income')
    print(f"  Expense accounts: {expense_count // 3 if expense_count else 0}")
    print(f"  Income accounts: {income_count // 3 if income_count else 0}")
    
    return {
        'ic_mapping': ic_mapping,
        'internal_to_opposite_account': internal_to_opposite_account,
        'internal_to_account_type': internal_to_account_type,
        'internal_to_opposite_internal': internal_to_opposite_internal,
        'internal_to_account_name': internal_to_account_name,
        'internal_to_ls_account': internal_to_ls_account
    }


def process_csv_elimination(csv_file, mapping_data):
    """
    Process CSV file and extract ONLY opposite account entries for IC transactions.
    Returns DataFrame with elimination entries only.
    """
    # Read CSV with UTF-8 encoding
    try:
        df = pd.read_csv(csv_file, encoding='utf-8')
    except UnicodeDecodeError:
        try:
            df = pd.read_csv(csv_file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            df = pd.read_csv(csv_file, encoding='latin-1')
    
    print(f"Loaded {len(df)} rows from CSV")
    
    # Extract mapping dictionaries
    ic_mapping = mapping_data['ic_mapping']
    internal_to_opposite_account = mapping_data['internal_to_opposite_account']
    internal_to_opposite_internal = mapping_data['internal_to_opposite_internal']
    internal_to_account_name = mapping_data['internal_to_account_name']
    internal_to_ls_account = mapping_data['internal_to_ls_account']
    internal_to_account_type = mapping_data.get('internal_to_account_type', {})
    
    # Find required columns
    internal_col = None
    ref_col = None
    account_col = None
    account_name_col = None
    date_col = None
    
    for col in df.columns:
        col_clean = str(col).strip().lower()
        if col_clean == 'internal' or col_clean == 'internal id':
            internal_col = col
        elif col_clean == 'ref':
            ref_col = col
        elif col_clean == 'account':
            account_col = col
        elif 'account name' in col_clean or col_clean == 'account name':
            account_name_col = col
        elif col_clean == 'date':
            date_col = col
    
    if internal_col is None:
        raise ValueError("Could not find 'Internal' column in CSV file")
    if ref_col is None:
        raise ValueError("Could not find 'REF' column in CSV file")
    
    print(f"Internal column: '{internal_col}'")
    print(f"REF column: '{ref_col}'")
    
    # Determine I/C status for each row
    def is_intercompany(internal_val):
        if internal_val in ic_mapping:
            return ic_mapping[internal_val]
        try:
            int_val = int(internal_val)
            if int_val in ic_mapping:
                return ic_mapping[int_val]
            if str(int_val) in ic_mapping:
                return ic_mapping[str(int_val)]
        except (ValueError, TypeError):
            pass
        return False
    
    # Mark each row as IC or not
    df['_is_ic_row'] = df[internal_col].apply(is_intercompany)
    
    # Find all REF values that contain at least one IC transaction
    ic_refs = df[df['_is_ic_row']][ref_col].unique()
    print(f"Found {len(ic_refs)} unique REF batches containing IC transactions")
    
    # Mark all rows belonging to IC batches
    df['_is_ic_batch'] = df[ref_col].isin(ic_refs)
    
    # Get only IC transactions
    df_ic = df[df['_is_ic_batch']].copy()
    
    if len(df_ic) == 0:
        print("No IC transactions found!")
        return pd.DataFrame()
    
    # Store original internal values
    df_ic['_original_internal'] = df_ic[internal_col]
    
    print(f"Total IC rows before filtering: {len(df_ic)}")
    
    # === Filter for only Expense and Income accounts ===
    def is_expense_or_income(internal_val):
        """Check if account type is Expense or Income"""
        account_type = None
        if internal_val in internal_to_account_type:
            account_type = internal_to_account_type[internal_val]
        else:
            try:
                int_val = int(internal_val)
                if int_val in internal_to_account_type:
                    account_type = internal_to_account_type[int_val]
                elif str(int_val) in internal_to_account_type:
                    account_type = internal_to_account_type[str(int_val)]
            except (ValueError, TypeError):
                pass
        
        if account_type:
            account_type_lower = account_type.lower()
            return account_type_lower in ['expense', 'expenses', 'income']
        return False
    
    # Filter to only include Expense and Income accounts
    df_ic['_is_expense_income'] = df_ic[internal_col].apply(is_expense_or_income)
    rows_before_filter = len(df_ic)
    df_ic = df_ic[df_ic['_is_expense_income']].copy()
    rows_after_filter = len(df_ic)
    
    print(f"  Filtered to Expense/Income accounts: {rows_after_filter} rows (removed {rows_before_filter - rows_after_filter} non-expense/income rows)")
    
    if len(df_ic) == 0:
        print("No Expense or Income IC transactions found!")
        return pd.DataFrame()
    
    # Clean up helper column
    df_ic = df_ic.drop(columns=['_is_expense_income'])
    
    print(f"Processing {len(df_ic)} IC Expense/Income rows for elimination entries...")
    
    # === Create elimination entries (opposite accounts only) ===
    df_elimination = df_ic.copy()
    
    # Set subsidiary values for elimination
    if 'Subsidiary Line' in df_elimination.columns:
        df_elimination['Subsidiary Line'] = 5  # xElimination - Parent
    
    if 'Subsidiary Line/ DUE TO/ FROM SUBSIDIARY' in df_elimination.columns:
        df_elimination['Subsidiary Line/ DUE TO/ FROM SUBSIDIARY'] = 6
    
    # Clear Name column for elimination entries
    df_elimination['Name'] = ''
    
    # Replace account with opposite account
    if account_col:
        def get_opposite_account(row):
            original_internal = row['_original_internal']
            opposite_account = None
            
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
            
            if opposite_account is not None:
                try:
                    return int(float(opposite_account))
                except (ValueError, TypeError):
                    return opposite_account
            
            return row[account_col]
        
        df_elimination[account_col] = df_elimination.apply(get_opposite_account, axis=1)
    
    # Replace internal with opposite internal
    def get_opposite_internal(row):
        original_internal = row['_original_internal']
        
        if original_internal in internal_to_opposite_internal:
            return internal_to_opposite_internal[original_internal]
        
        try:
            int_internal = int(original_internal)
            if int_internal in internal_to_opposite_internal:
                return internal_to_opposite_internal[int_internal]
            if str(int_internal) in internal_to_opposite_internal:
                return internal_to_opposite_internal[str(int_internal)]
        except (ValueError, TypeError):
            pass
        
        return original_internal
    
    df_elimination[internal_col] = df_elimination.apply(get_opposite_internal, axis=1)
    
    # Update Account Name based on new internal ID
    if account_name_col and internal_to_account_name:
        def get_account_name(row):
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
        
        df_elimination[account_name_col] = df_elimination.apply(get_account_name, axis=1)
    
    # === SWAP DEBIT AND CREDIT to create reversing/balancing entries ===
    # Find debit and credit columns
    debit_col = None
    credit_col = None
    for col in df_elimination.columns:
        col_clean = str(col).strip().lower()
        if 'debit' in col_clean:
            debit_col = col
        elif 'credit' in col_clean:
            credit_col = col
    
    if debit_col and credit_col:
        print(f"  Swapping Debit and Credit columns to create elimination entries...")
        # Store original values
        original_debit = df_elimination[debit_col].copy()
        original_credit = df_elimination[credit_col].copy()
        # Swap: debit becomes credit, credit becomes debit
        df_elimination[debit_col] = original_credit
        df_elimination[credit_col] = original_debit
        print(f"  Debit/Credit swapped for {len(df_elimination)} rows")
        
        # Show totals
        final_debit = pd.to_numeric(df_elimination[debit_col], errors='coerce').fillna(0).sum()
        final_credit = pd.to_numeric(df_elimination[credit_col], errors='coerce').fillna(0).sum()
        print(f"  Totals - Debit: {final_debit:,.2f}, Credit: {final_credit:,.2f}")
    
    # Remove helper columns
    helper_cols = ['_is_ic_row', '_is_ic_batch', '_original_internal']
    df_elimination = df_elimination.drop(columns=[c for c in helper_cols if c in df_elimination.columns])
    
    # Add EXTERNAL ID column
    if date_col and date_col in df_elimination.columns:
        try:
            dates = pd.to_datetime(df_elimination[date_col], dayfirst=True, errors='coerce')
            newest_date = dates.max()
            if pd.notna(newest_date):
                month_num = int(newest_date.month)
                year_short = int(newest_date.year) % 100
                external_id = f'JE_ELIM{month_num:02d}{year_short:02d}'
                df_elimination.insert(0, 'EXTERNAL ID', external_id)
        except:
            df_elimination.insert(0, 'EXTERNAL ID', 'JE_ELIM')
    else:
        df_elimination.insert(0, 'EXTERNAL ID', 'JE_ELIM')
    
    # Filter out rows where both debit and credit are zero
    debit_col = None
    credit_col = None
    for col in df_elimination.columns:
        col_clean = str(col).strip().lower()
        if 'debit' in col_clean:
            debit_col = col
        elif 'credit' in col_clean:
            credit_col = col
    
    if debit_col and credit_col:
        rows_before = len(df_elimination)
        debit_vals = pd.to_numeric(df_elimination[debit_col], errors='coerce').fillna(0)
        credit_vals = pd.to_numeric(df_elimination[credit_col], errors='coerce').fillna(0)
        df_elimination = df_elimination[(debit_vals != 0) | (credit_vals != 0)]
        rows_filtered = rows_before - len(df_elimination)
        if rows_filtered > 0:
            print(f"  Filtered out {rows_filtered} rows with zero debit and credit")
    
    # Sort by REF
    if ref_col and ref_col in df_elimination.columns:
        df_elimination = df_elimination.sort_values(by=ref_col).reset_index(drop=True)
    
    print(f"\nElimination entries created: {len(df_elimination)} rows")
    
    return df_elimination


def get_statscore_import_folder():
    """Create and return the path to "Statscore Import" folder on the desktop."""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    statscore_folder = os.path.join(desktop_path, "Statscore Import")
    
    if not os.path.exists(statscore_folder):
        os.makedirs(statscore_folder)
        print(f"  Created folder: {statscore_folder}")
    
    return statscore_folder


def save_csv(df, output_path):
    """Save DataFrame to CSV file with UTF-8 encoding"""
    df.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"Saved {len(df)} rows to {output_path}")


def get_month_filename(posting_period, prefix="CSV_Statscore_elimination"):
    """Convert posting period to filename format."""
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
    """Split DataFrame by Posting Period and save each month to a separate file."""
    saved_files = []
    statscore_folder = get_statscore_import_folder()
    
    # Find the Posting Period column
    posting_period_col = None
    for col in df.columns:
        if 'posting period' in str(col).lower():
            posting_period_col = col
            break
    
    if posting_period_col is None:
        print(f"  Warning: Could not find 'Posting Period' column. Saving all to single file.")
        output_path = os.path.join(statscore_folder, f"{file_prefix}.csv")
        save_csv(df, output_path)
        return [output_path], statscore_folder
    
    # Group by Posting Period
    grouped = df.groupby(posting_period_col)
    
    print(f"\n  Splitting {file_prefix} by Posting Period...")
    print(f"  Saving to: {statscore_folder}")
    
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
    print("NetSuite IC Elimination Processor")
    print("=" * 50)
    print("This processor creates ONLY elimination entries (opposite accounts)")
    print()
    
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
        
        # Process CSV - returns elimination entries only
        print("\nProcessing CSV file for elimination entries...")
        df_elimination = process_csv_elimination(csv_file, mapping_dict)
        
        if len(df_elimination) == 0:
            root = tk.Tk()
            root.withdraw()
            messagebox.showwarning("No Data", "No IC transactions found for elimination.")
            root.destroy()
            return
        
        # Save results
        print("\nSaving elimination entries...")
        
        elim_files, statscore_folder = save_by_month(df_elimination, "CSV_Statscore_elimination")
        
        # Build success message
        elim_files_msg = ""
        if isinstance(elim_files[0], tuple):
            for posting_period, path, count in elim_files:
                filename = os.path.basename(path)
                elim_files_msg += f"  {filename} ({count} rows)\n"
        else:
            elim_files_msg = f"  {os.path.basename(elim_files[0])} ({len(df_elimination)} rows)\n"
        
        # Show success message
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Success", 
            f"Elimination processing complete!\n\n"
            f"All files saved to:\n{statscore_folder}\n\n"
            f"Elimination entries ({len(df_elimination)} rows):\n{elim_files_msg}")
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

