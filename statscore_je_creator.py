"""
Statscore JE Creator
--------------------
Opens a GUI to select a Statscore CSV file, previews the Journal Entry,
and posts it directly to NetSuite via the REST API.

All CSV lines become one single JE under Statscore subsidiary (ID 6).

Requires:
    - .env file with NS OAuth credentials at \\\\mainsrv\\d\\Cloudpay_Script\\Banks Dashboard\\.env
    - Statscore-specific tokens: NETSUITE_TOKEN_ID_6, NETSUITE_TOKEN_SECRET_6
"""

import sys
import os
import time
import hmac
import hashlib
import base64
import urllib.parse
import urllib.request
import urllib.error
import json
import secrets
import webbrowser
import threading
import calendar
from datetime import datetime

import subprocess
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ---------------------------------------------------------------------------
# NetSuite constants
# ---------------------------------------------------------------------------
NS_ENV_PATH = r"\\mainsrv\d\Cloudpay_Script\Banks Dashboard\.env"

NS_SUBSIDIARY_ID = 6          # Statscore
NS_CURRENCY_EUR  = 1          # EUR
NS_DEPT_ID       = 24         # Statscore import
NS_LOC_POLAND    = 5          # Poland
NS_CSEG_OTHER    = 3          # Location (Expenses) = Other

# ---------------------------------------------------------------------------
# NetSuite helpers (reused from JE_Processor / bank_commission_je)
# ---------------------------------------------------------------------------

def load_ns_env():
    """Load .env file. Override tokens with Statscore-specific ones (_6 suffix)."""
    env = {}
    if not os.path.exists(NS_ENV_PATH):
        return env
    with open(NS_ENV_PATH) as f:
        for line in f:
            line = line.strip()
            if '=' in line and not line.startswith('#'):
                k, v = line.split('=', 1)
                env[k.strip()] = v.strip()
    # Use Statscore tokens (subsidiary 6)
    if env.get('NETSUITE_TOKEN_ID_6'):
        env['NETSUITE_TOKEN_ID'] = env['NETSUITE_TOKEN_ID_6']
    if env.get('NETSUITE_TOKEN_SECRET_6'):
        env['NETSUITE_TOKEN_SECRET'] = env['NETSUITE_TOKEN_SECRET_6']
    return env


def ns_gen_auth(env, method, url):
    """OAuth 1.0 HMAC-SHA256 header."""
    parsed = urllib.parse.urlparse(url)
    base_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
    ts = str(int(time.time()))
    nonce = secrets.token_hex(16)
    oauth_params = {
        'oauth_consumer_key': env['NETSUITE_CONSUMER_KEY'],
        'oauth_token': env['NETSUITE_TOKEN_ID'],
        'oauth_nonce': nonce, 'oauth_timestamp': ts,
        'oauth_signature_method': 'HMAC-SHA256', 'oauth_version': '1.0',
    }
    all_params = dict(oauth_params)
    if parsed.query:
        for qp in parsed.query.split('&'):
            if '=' in qp:
                qk, qv = qp.split('=', 1)
                all_params[qk] = qv
    ps = '&'.join(
        f"{urllib.parse.quote(k, safe='')}={urllib.parse.quote(v, safe='')}"
        for k, v in sorted(all_params.items()))
    bs = f"{method.upper()}&{urllib.parse.quote(base_url, safe='')}&{urllib.parse.quote(ps, safe='')}"
    sk = (urllib.parse.quote(env['NETSUITE_CONSUMER_SECRET'], safe='') + '&' +
          urllib.parse.quote(env['NETSUITE_TOKEN_SECRET'], safe=''))
    sig = base64.b64encode(
        hmac.new(sk.encode(), bs.encode(), hashlib.sha256).digest()).decode()
    oauth_params['oauth_signature'] = sig
    parts = ', '.join(
        f'{k}="{urllib.parse.quote(v, safe="")}"'
        for k, v in sorted(oauth_params.items()))
    return f'OAuth realm="{env["NETSUITE_ACCOUNT_ID"]}", {parts}'


def ns_suiteql(env, query):
    """Execute a SuiteQL query, return list of result items."""
    url = f"https://{env['NETSUITE_ACCOUNT_ID']}.suitetalk.api.netsuite.com/services/rest/query/v1/suiteql"
    auth = ns_gen_auth(env, 'POST', url)
    data = json.dumps({"q": query}).encode()
    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', auth)
    req.add_header('Content-Type', 'application/json')
    req.add_header('Prefer', 'transient')
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode()).get('items', [])


def ns_post_je(env, payload):
    """POST a Journal Entry to NetSuite REST Record API.
    Returns (internal_id, tran_id, error)."""
    url = f"https://{env['NETSUITE_ACCOUNT_ID']}.suitetalk.api.netsuite.com/services/rest/record/v1/journalEntry"
    auth = ns_gen_auth(env, 'POST', url)
    data = json.dumps(payload).encode()
    req = urllib.request.Request(url, data=data, method='POST')
    req.add_header('Authorization', auth)
    req.add_header('Content-Type', 'application/json')
    try:
        with urllib.request.urlopen(req, timeout=600) as resp:
            loc = resp.getheader('Location', '')
            resp.read()
            je_id = loc.rstrip('/').split('/')[-1] if loc else ''
            tran_id = ''
            if je_id:
                try:
                    items = ns_suiteql(env,
                        f"SELECT tranid FROM transaction WHERE id = {je_id} "
                        "OFFSET 0 ROWS FETCH NEXT 1 ROWS ONLY")
                    if items:
                        tran_id = items[0].get('tranid', '')
                except Exception:
                    pass
            return je_id, tran_id, None
    except urllib.error.HTTPError as e:
        err_body = e.read().decode()
        return None, None, f"HTTP {e.code}: {err_body[:1000]}"


def ns_je_url(env, je_id):
    acct = (env.get('NETSUITE_ACCOUNT_ID') or '').replace('_', '-').lower()
    return f"https://{acct}.app.netsuite.com/app/accounting/transactions/journal.nl?id={je_id}"


def open_in_chrome(url):
    """Open a URL in Chrome browser."""
    chrome_paths = [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
    ]
    for chrome in chrome_paths:
        if os.path.exists(chrome):
            subprocess.Popen([chrome, url])
            return
    # Fallback to default browser
    webbrowser.open(url)


# ---------------------------------------------------------------------------
# Account validation
# ---------------------------------------------------------------------------

def validate_accounts_ns(env, internal_ids):
    """Validate that all internal IDs exist as accounts in NetSuite.
    Returns (valid_dict, invalid_list) where valid_dict maps id -> {acctnumber, fullname}."""
    valid = {}
    invalid = []

    # Query in batches of 100
    id_list = list(internal_ids)
    for i in range(0, len(id_list), 100):
        batch = id_list[i:i+100]
        ids_str = ','.join(str(x) for x in batch)
        try:
            items = ns_suiteql(env,
                f"SELECT a.id, a.acctnumber, a.fullname "
                f"FROM account a WHERE a.id IN ({ids_str})")
            for item in items:
                aid = int(item['id'])
                valid[aid] = {
                    'acctnumber': item.get('acctnumber', ''),
                    'fullname': item.get('fullname', ''),
                }
        except Exception as e:
            raise RuntimeError(f"SuiteQL query failed: {e}")

    for aid in id_list:
        if int(aid) not in valid:
            invalid.append(int(aid))

    return valid, invalid


# ---------------------------------------------------------------------------
# CSV parsing & JE building
# ---------------------------------------------------------------------------

def parse_csv(csv_path):
    """Read the Statscore CSV, return DataFrame with cleaned columns."""
    for enc in ('windows-1252', 'utf-8-sig', 'utf-8', 'iso-8859-1'):
        try:
            df = pd.read_csv(csv_path, encoding=enc)
            break
        except (UnicodeDecodeError, UnicodeError):
            continue
    else:
        raise ValueError(f"Could not read CSV with any supported encoding")

    # Clean column names
    df.columns = [c.strip() for c in df.columns]

    # Drop fully empty rows
    df = df.dropna(how='all').reset_index(drop=True)

    # Detect column names (flexible matching)
    col_map = {}
    for c in df.columns:
        cl = c.lower().strip()
        if cl == 'internal':
            col_map['internal'] = c
        elif cl == 'account':
            col_map['account'] = c
        elif cl == 'debit':
            col_map['debit'] = c
        elif cl == 'credit':
            col_map['credit'] = c
        elif cl in ('linememo', 'line memo', 'line_memo'):
            col_map['memo'] = c
        elif cl == 'ref':
            col_map['ref'] = c
        elif cl == 'date':
            col_map['date'] = c
        elif 'class' in cl:
            col_map['class'] = c
        elif 'location' in cl and 'expense' in cl:
            col_map['loc_exp'] = c
        elif cl in ('currecny', 'currency'):
            col_map['currency'] = c

    required = ['internal', 'debit', 'credit', 'memo', 'date']
    missing = [k for k in required if k not in col_map]
    if missing:
        raise ValueError(f"Missing required columns: {missing}. Found: {list(df.columns)}")

    return df, col_map


def detect_month_year(df, col_map):
    """Detect the month and year from the Date column. Returns (month, year) as ints."""
    date_col = col_map['date']
    dates = pd.to_datetime(df[date_col], format='%d/%m/%y', errors='coerce')
    dates = dates.dropna()
    if dates.empty:
        raise ValueError("No valid dates found in the Date column")
    # Use the most common month
    month = dates.dt.month.mode().iloc[0]
    year = dates.dt.year.mode().iloc[0]
    return int(month), int(year)


def build_je_payload(df, col_map, month, year):
    """Build the NetSuite JE payload from the parsed CSV.
    Skips lines where both Debit and Credit are zero."""
    # Transaction date = last day of month
    last_day = calendar.monthrange(year, month)[1]
    tran_date = f"{year}-{month:02d}-{last_day:02d}"
    memo = f"Statscore {month:02d}/{year}"

    je_lines = []
    skipped_zero = 0
    for _, row in df.iterrows():
        internal_id = row[col_map['internal']]
        if pd.isna(internal_id):
            continue
        internal_id = str(int(float(internal_id)))

        debit = float(row[col_map['debit']]) if pd.notna(row[col_map['debit']]) else 0.0
        credit = float(row[col_map['credit']]) if pd.notna(row[col_map['credit']]) else 0.0

        # Skip lines where both debit and credit are zero
        if debit == 0 and credit == 0:
            skipped_zero += 1
            continue

        line_memo = str(row[col_map['memo']]) if pd.notna(row[col_map['memo']]) else ''

        line = {
            "account": {"id": internal_id},
            "memo": line_memo,
            "department": {"id": str(NS_DEPT_ID)},
            "location": {"id": str(NS_LOC_POLAND)},
            "cseg_location_exp": {"id": str(NS_CSEG_OTHER)},
        }

        # Handle positive, negative, and zero amounts
        # Negative debit = credit (reversal), negative credit = debit (reversal)
        if debit > 0:
            line["debit"] = round(debit, 2)
        elif debit < 0:
            line["credit"] = round(abs(debit), 2)

        if credit > 0:
            line["credit"] = round(credit, 2)
        elif credit < 0:
            line["debit"] = round(abs(credit), 2)

        je_lines.append(line)

    payload = {
        "subsidiary": {"id": str(NS_SUBSIDIARY_ID)},
        "currency": {"id": str(NS_CURRENCY_EUR)},
        "tranDate": tran_date,
        "memo": memo,
        "approved": True,
        "line": {"items": je_lines},
    }

    return payload, skipped_zero


def validate_ref_balance(df, col_map):
    """Validate that each REF group is balanced (total debit == total credit).
    Returns (ok_count, errors) where errors is a list of (ref, debit, credit, diff)."""
    if 'ref' not in col_map:
        return 0, []

    ref_col = col_map['ref']
    debit_col = col_map['debit']
    credit_col = col_map['credit']

    errors = []
    ok_count = 0

    for ref_val, group in df.groupby(ref_col):
        # Sum raw debit/credit (positive and negative)
        total_d = group[debit_col].fillna(0).sum()
        total_c = group[credit_col].fillna(0).sum()
        diff = abs(total_d - total_c)
        if diff >= 0.02:
            errors.append((ref_val, total_d, total_c, diff))
        else:
            ok_count += 1

    return ok_count, errors


# ---------------------------------------------------------------------------
# GUI Application
# ---------------------------------------------------------------------------

class StatscoreJECreator:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Statscore JE Creator")

        win_w, win_h = 780, 600
        sx = self.root.winfo_screenwidth() // 2 - win_w // 2
        sy = self.root.winfo_screenheight() // 2 - win_h // 2
        self.root.geometry(f"{win_w}x{win_h}+{sx}+{sy}")
        self.root.resizable(True, True)

        self.csv_var = tk.StringVar()
        self.payload = None
        self.env = None

        self._build_gui()
        self._load_env()

    def _build_gui(self):
        # Title bar
        tf = tk.Frame(self.root, bg="#2c3e50", height=44)
        tf.pack(fill="x")
        tf.pack_propagate(False)
        tk.Label(tf, text="Statscore JE Creator",
                 font=("Segoe UI", 13, "bold"), fg="white", bg="#2c3e50").pack(expand=True)

        # CSV file selector
        sf = tk.LabelFrame(self.root, text="  CSV File  ", padx=8, pady=6,
                           font=("Segoe UI", 10))
        sf.pack(fill="x", padx=12, pady=(8, 4))
        tk.Entry(sf, textvariable=self.csv_var, font=("Segoe UI", 10), width=60
                 ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        tk.Button(sf, text="Browse...", command=self._browse_csv,
                  font=("Segoe UI", 9), width=9).pack(side="right")

        # Log area
        self.log = scrolledtext.ScrolledText(self.root, height=24,
                                              font=("Consolas", 9),
                                              state="disabled", wrap="word")
        self.log.pack(fill="both", expand=True, padx=12, pady=(4, 4))

        # Configure tag for links
        self.log.tag_configure("link", foreground="blue", underline=True)
        self.log.tag_bind("link", "<Button-1>", self._on_link_click)
        self.log.tag_bind("link", "<Enter>",
                          lambda e: self.log.config(cursor="hand2"))
        self.log.tag_bind("link", "<Leave>",
                          lambda e: self.log.config(cursor=""))

        # Buttons
        bf = tk.Frame(self.root)
        bf.pack(pady=(2, 10))

        tk.Button(bf, text="Preview", command=self._do_preview,
                  font=("Segoe UI", 10, "bold"), width=12,
                  bg="#2980b9", fg="white", cursor="hand2").pack(side="left", padx=6)

        tk.Button(bf, text="Post to NetSuite", command=self._do_post,
                  font=("Segoe UI", 10, "bold"), width=16,
                  bg="#c0392b", fg="white", cursor="hand2").pack(side="left", padx=6)

        tk.Button(bf, text="Close", command=self.root.destroy,
                  font=("Segoe UI", 9), width=8).pack(side="left", padx=6)

    def _browse_csv(self):
        p = filedialog.askopenfilename(
            title="Select Statscore CSV File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if p:
            self.csv_var.set(p)

    def _load_env(self):
        try:
            self.env = load_ns_env()
            if not self.env.get('NETSUITE_ACCOUNT_ID'):
                self._log("WARNING: No NS credentials found at " + NS_ENV_PATH + "\n")
            else:
                self._log("NS credentials loaded (Statscore tokens).\n")
        except Exception as e:
            self._log(f"ERROR loading credentials: {e}\n")

    def _log(self, text):
        self.log.config(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.config(state="disabled")

    def _log_link(self, text, url):
        self.log.config(state="normal")
        self.log.insert("end", text, "link")
        self.log.config(state="disabled")
        self._link_url = url

    def _on_link_click(self, event):
        if hasattr(self, '_link_url'):
            webbrowser.open(self._link_url)

    def _get_csv_path(self):
        path = self.csv_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("No file", "Please select a valid CSV file.")
            return None
        return path

    def _do_preview(self):
        path = self._get_csv_path()
        if not path:
            return

        self._log("\n" + "=" * 65 + "\n")
        self._log("PREVIEW - Parsing CSV...\n")

        try:
            df, col_map = parse_csv(path)
            month, year = detect_month_year(df, col_map)
            self._log(f"  CSV rows: {len(df)}\n")
            self._log(f"  Detected period: {month:02d}/{year}\n")

            # Validate per-REF balance first
            self._log("\n  Validating REF group balances...\n")
            ok_count, ref_errors = validate_ref_balance(df, col_map)
            if ref_errors:
                self._log(f"\n  *** {len(ref_errors)} UNBALANCED REF groups ***\n")
                self._log(f"  {'REF':>6s}  {'Debit':>14s}  {'Credit':>14s}  {'Diff':>14s}\n")
                self._log(f"  {'-'*6}  {'-'*14}  {'-'*14}  {'-'*14}\n")
                for ref_val, rd, rc, rdiff in ref_errors[:30]:
                    self._log(f"  {str(ref_val):>6s}  {rd:>14,.2f}  {rc:>14,.2f}  {rdiff:>14,.2f}\n")
                if len(ref_errors) > 30:
                    self._log(f"  ... and {len(ref_errors) - 30} more\n")
                self._log(f"\n  BLOCKED: Fix unbalanced REFs before posting.\n")
                self._log("=" * 65 + "\n")
                self.payload = None
                return
            else:
                self._log(f"  All {ok_count} REF groups are balanced.\n")

            payload, skipped_zero = build_je_payload(df, col_map, month, year)
            self.payload = payload

            je_lines = payload['line']['items']
            total_d = sum(ln.get('debit', 0) for ln in je_lines)
            total_c = sum(ln.get('credit', 0) for ln in je_lines)
            diff = abs(total_d - total_c)
            bal = "BALANCED" if diff < 0.02 else f"*** UNBALANCED (diff={diff:,.2f}) ***"

            self._log(f"\n  JE Header:\n")
            self._log(f"    Subsidiary: Statscore (ID {NS_SUBSIDIARY_ID})\n")
            self._log(f"    Currency:   EUR (ID {NS_CURRENCY_EUR})\n")
            self._log(f"    Date:       {payload['tranDate']}\n")
            self._log(f"    Memo:       {payload['memo']}\n")
            self._log(f"    Approved:   {payload['approved']}\n")
            self._log(f"\n  Lines: {len(je_lines)} (skipped {skipped_zero} zero-amount lines)\n")
            self._log(f"  Total Debit:  {total_d:>16,.2f}\n")
            self._log(f"  Total Credit: {total_c:>16,.2f}\n")
            self._log(f"  Balance:      {bal}\n")

            if diff >= 0.02:
                self._log(f"\n  BLOCKED: JE is unbalanced. Cannot post.\n")
                self._log("=" * 65 + "\n")
                self.payload = None
                return

            # Show first 20 lines as sample
            self._log(f"\n  Sample lines (first 20):\n")
            self._log(f"  {'Acct ID':>8s}  {'Debit':>14s}  {'Credit':>14s}  Memo\n")
            self._log(f"  {'-'*8}  {'-'*14}  {'-'*14}  {'-'*30}\n")
            for ln in je_lines[:20]:
                acct = ln['account']['id']
                d = ln.get('debit', 0)
                c = ln.get('credit', 0)
                m = ln.get('memo', '')[:40]
                self._log(f"  {acct:>8s}  {d:>14,.2f}  {c:>14,.2f}  {m}\n")
            if len(je_lines) > 20:
                self._log(f"  ... and {len(je_lines) - 20} more lines\n")

            # Count unique accounts and validate against NS
            unique_accts = set(ln['account']['id'] for ln in je_lines)
            self._log(f"\n  Unique accounts: {len(unique_accts)}\n")

            # Validate accounts in NetSuite
            if self.env and self.env.get('NETSUITE_ACCOUNT_ID'):
                self._log("  Validating accounts in NetSuite...\n")
                try:
                    valid, invalid = validate_accounts_ns(self.env, unique_accts)
                    if invalid:
                        self._log(f"\n  *** WARNING: {len(invalid)} INVALID account IDs ***\n")
                        for aid in invalid:
                            self._log(f"    Account ID {aid} - NOT FOUND in NetSuite\n")
                        self._log("  These lines will FAIL when posting!\n")
                        self.payload = None  # Block posting
                    else:
                        self._log(f"  All {len(valid)} accounts validated OK\n")
                        # Show account mapping
                        self._log("\n  Account mapping:\n")
                        for aid in sorted(valid.keys()):
                            v = valid[aid]
                            self._log(f"    {aid:>6d}  {v['acctnumber']:>10s}  {v['fullname']}\n")
                except Exception as e:
                    self._log(f"  WARNING: Could not validate accounts: {e}\n")

            # Show unique REF count if available
            if 'ref' in col_map:
                ref_count = df[col_map['ref']].nunique()
                self._log(f"  Unique REF groups: {ref_count}\n")

            self._log("\n" + "=" * 65 + "\n")
            self._log("PREVIEW COMPLETE - Ready to post.\n")

        except Exception as e:
            self._log(f"\nERROR: {e}\n")
            import traceback
            self._log(traceback.format_exc())

    def _do_post(self):
        if not self.env or not self.env.get('NETSUITE_ACCOUNT_ID'):
            messagebox.showerror("Error", "No NetSuite credentials loaded.")
            return

        if self.payload is None:
            messagebox.showwarning("Preview first",
                "Please run Preview first to validate the CSV.")
            return

        je_lines = self.payload['line']['items']
        total_d = sum(ln.get('debit', 0) for ln in je_lines)
        total_c = sum(ln.get('credit', 0) for ln in je_lines)

        if abs(total_d - total_c) >= 0.02:
            messagebox.showerror("Unbalanced",
                f"JE is UNBALANCED - cannot post.\n\n"
                f"Debit: {total_d:,.2f}\nCredit: {total_c:,.2f}\n"
                f"Diff: {abs(total_d - total_c):,.2f}")
            return

        if not messagebox.askyesno("Confirm",
                f"Post 1 JE to NetSuite?\n\n"
                f"Lines: {len(je_lines)}\n"
                f"Debit: {total_d:,.2f}\n"
                f"Credit: {total_c:,.2f}\n"
                f"Memo: {self.payload['memo']}\n"
                f"Date: {self.payload['tranDate']}\n\n"
                f"This may take a few minutes for large JEs."):
            return

        # Post in background thread
        self._log("\n" + "=" * 65 + "\n")
        self._log(f"POSTING to NetSuite ({len(je_lines)} lines, please wait)...\n")
        threading.Thread(target=self._post_thread, daemon=True).start()

    def _post_thread(self):
        """Post the single JE payload to NetSuite."""
        try:
            je_id, tran_id, err = ns_post_je(self.env, self.payload)

            if err:
                self.root.after(0, self._log, f"\nERROR: {err}\n")
                return

            url = ns_je_url(self.env, je_id)
            self.root.after(0, self._log,
                f"\nSUCCESS!\n"
                f"  Internal ID: {je_id}\n"
                f"  JE Number:   {tran_id}\n"
                f"  URL: ")
            self.root.after(0, self._log_link, url + "\n", url)
            self.root.after(0, self._log, "\n" + "=" * 65 + "\n")

            # Open the JE in Chrome
            self.root.after(500, lambda: open_in_chrome(url))

        except Exception as e:
            self.root.after(0, self._log, f"\nERROR: {e}\n")
            import traceback
            self.root.after(0, self._log, traceback.format_exc())

    def run(self):
        # Pre-fill CSV if passed as command line arg
        if len(sys.argv) > 1 and os.path.exists(sys.argv[1]):
            self.csv_var.set(sys.argv[1])
        self.root.mainloop()


if __name__ == '__main__':
    app = StatscoreJECreator()
    app.run()
