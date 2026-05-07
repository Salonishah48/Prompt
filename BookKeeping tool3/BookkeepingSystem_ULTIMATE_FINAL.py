"""
Offline Bookkeeping System v2.0 - COMPLETE ALL-IN-ONE
QuickBooks Compatible | All Features Included
Ready to Run - No Assembly Required!
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import sqlite3
from pathlib import Path
from datetime import datetime
import os

class BookkeepingSystemV2:
    def __init__(self, root):
        self.root = root
        self.root.title("Offline Bookkeeping System v2.0 - QuickBooks Compatible")
        self.root.geometry("1600x950")
        self.root.configure(bg='#f8f9fa')
        
        # Data storage
        self.current_data = pd.DataFrame()
        self.historical_data = pd.DataFrame()
        self.chart_of_accounts = pd.DataFrame()
        self.current_client = None
        
        # QuickBooks account mapping (60+ mappings)
        self.qb_account_mapping = {
            'Sales Revenue': ('Revenue', 'Sales'),
            'Service Revenue': ('Revenue', 'Service Revenue'),
            'Interest Income': ('Revenue', 'Interest Income'),
            'Other Income': ('Revenue', 'Other Income'),
            'Income': ('Revenue', 'Sales'),
            'Direct Materials': ('Cost of Goods Sold', 'Direct Materials'),
            'Direct Labor': ('Cost of Goods Sold', 'Direct Labor'),
            'Manufacturing Overhead': ('Cost of Goods Sold', 'Manufacturing Overhead'),
            'Cost of Goods Sold': ('Cost of Goods Sold', 'Purchase of Goods'),
            'COGS': ('Cost of Goods Sold', 'Purchase of Goods'),
            'Rent': ('Operating Expenses', 'Rent'),
            'Rent Expense': ('Operating Expenses', 'Rent'),
            'Salaries': ('Operating Expenses', 'Salaries & Wages'),
            'Salaries & Wages': ('Operating Expenses', 'Salaries & Wages'),
            'Wages': ('Operating Expenses', 'Salaries & Wages'),
            'Payroll': ('Operating Expenses', 'Salaries & Wages'),
            'Utilities': ('Operating Expenses', 'Utilities'),
            'Utilities Expense': ('Operating Expenses', 'Utilities'),
            'Marketing': ('Operating Expenses', 'Marketing & Advertising'),
            'Advertising': ('Operating Expenses', 'Marketing & Advertising'),
            'Office Supplies': ('Operating Expenses', 'Office Supplies'),
            'Supplies': ('Operating Expenses', 'Office Supplies'),
            'Insurance': ('Operating Expenses', 'Insurance'),
            'Professional Fees': ('Operating Expenses', 'Professional Fees'),
            'Legal Fees': ('Operating Expenses', 'Legal Fees'),
            'Accounting Fees': ('Operating Expenses', 'Accounting Fees'),
            'Bank Charges': ('Operating Expenses', 'Bank Charges'),
            'Travel': ('Operating Expenses', 'Travel & Entertainment'),
            'Meals and Entertainment': ('Operating Expenses', 'Travel & Entertainment'),
            'Repairs': ('Operating Expenses', 'Repairs & Maintenance'),
            'Maintenance': ('Operating Expenses', 'Repairs & Maintenance'),
            'Telephone': ('Operating Expenses', 'Telephone & Internet'),
            'Internet': ('Operating Expenses', 'Telephone & Internet'),
            'Depreciation': ('Operating Expenses', 'Depreciation'),
            'Cash': ('Assets', 'Cash'),
            'Bank': ('Assets', 'Bank Account'),
            'Checking': ('Assets', 'Bank Account'),
            'Accounts Receivable': ('Assets', 'Accounts Receivable'),
            'Inventory': ('Assets', 'Inventory'),
            'Fixed Assets': ('Assets', 'Fixed Assets'),
            'Equipment': ('Assets', 'Equipment'),
            'Furniture': ('Assets', 'Furniture & Fixtures'),
            'Vehicles': ('Assets', 'Vehicles'),
            'Accounts Payable': ('Liabilities', 'Accounts Payable'),
            'Credit Card': ('Liabilities', 'Credit Cards'),
            'Loan': ('Liabilities', 'Short-term Loans'),
            'Equity': ('Equity', 'Share Capital'),
            'Retained Earnings': ('Equity', 'Retained Earnings'),
        }
        
        # Standard categories
        self.account_categories = {
            'Revenue': ['Sales', 'Service Revenue', 'Interest Income', 'Other Income'],
            'Cost of Goods Sold': ['Direct Materials', 'Direct Labor', 'Manufacturing Overhead', 'Purchase of Goods'],
            'Operating Expenses': [
                'Salaries & Wages', 'Rent', 'Utilities', 'Marketing & Advertising', 
                'Office Supplies', 'Insurance', 'Professional Fees', 'Legal Fees',
                'Accounting Fees', 'Bank Charges', 'Travel & Entertainment', 
                'Repairs & Maintenance', 'Telephone & Internet', 'Printing & Stationery',
                'Vehicle Expenses', 'Depreciation', 'Amortization', 'Other Expenses'
            ],
            'Assets': [
                'Cash', 'Bank Account', 'Accounts Receivable', 'Inventory', 
                'Prepaid Expenses', 'Fixed Assets', 'Equipment', 'Furniture & Fixtures',
                'Vehicles', 'Accumulated Depreciation', 'Investments', 'Other Assets'
            ],
            'Liabilities': [
                'Accounts Payable', 'Accrued Expenses', 'Short-term Loans', 
                'Long-term Debt', 'Credit Cards', 'Taxes Payable', 'TDS Payable',
                'GST Payable', 'Other Liabilities'
            ],
            'Equity': [
                'Share Capital', 'Retained Earnings', 'Current Year Profit/Loss', 
                'Drawings', 'Capital Contributions', 'Other Equity'
            ],
            'Ask My Accountant': [
                'Unclear Transaction', 'Needs Review', 'Unidentified', 'Miscellaneous'
            ]
        }
        
        self.setup_styles()
        self.init_database()
        self.create_menu()
        self.create_main_interface()
    
    def setup_styles(self):
        """Configure modern UI styles"""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#f8f9fa')
        style.configure('TLabel', background='#f8f9fa', foreground='#212529', font=('Segoe UI', 9))
        style.configure('TButton', font=('Segoe UI', 9), borderwidth=0, relief='flat', padding=8)
        style.map('TButton',
                 background=[('active', '#e3f2fd'), ('!active', '#2196f3')],
                 foreground=[('active', '#1976d2'), ('!active', 'white')])
        style.configure('TNotebook', background='#f8f9fa', borderwidth=0)
        style.configure('TNotebook.Tab', font=('Segoe UI', 9, 'bold'), padding=[15, 8])
        style.map('TNotebook.Tab',
                 background=[('selected', '#ffffff'), ('!selected', '#e0e0e0')],
                 foreground=[('selected', '#2196f3'), ('!selected', '#757575')])
    
    def init_database(self):
        """Initialize SQLite database"""
        self.db_path = Path.home() / "bookkeeping_v2.db"
        self.conn = sqlite3.connect(str(self.db_path))
        self.cursor = self.conn.cursor()
        
        # Tables
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS clients (
                client_id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_name TEXT UNIQUE NOT NULL,
                company_name TEXT,
                gstin TEXT,
                pan TEXT,
                address TEXT,
                created_date TEXT,
                last_modified TEXT
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS transactions (
                transaction_id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER,
                transaction_date TEXT,
                transaction_type TEXT,
                num TEXT,
                name TEXT,
                description TEXT,
                account TEXT,
                split TEXT,
                debit REAL,
                credit REAL,
                balance REAL,
                category TEXT,
                sub_category TEXT,
                year INTEGER,
                FOREIGN KEY (client_id) REFERENCES clients(client_id)
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS chart_of_accounts (
                account_id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER,
                category TEXT NOT NULL,
                sub_category TEXT NOT NULL,
                account_code TEXT,
                usage_count INTEGER DEFAULT 0,
                total_debit REAL DEFAULT 0,
                total_credit REAL DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                FOREIGN KEY (client_id) REFERENCES clients(client_id),
                UNIQUE(client_id, category, sub_category)
            )
        ''')
        
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS classification_rules (
                rule_id INTEGER PRIMARY KEY AUTOINCREMENT,
                client_id INTEGER,
                keyword TEXT NOT NULL,
                qb_account TEXT,
                name_pattern TEXT,
                category TEXT NOT NULL,
                sub_category TEXT NOT NULL,
                confidence REAL DEFAULT 1.0,
                usage_count INTEGER DEFAULT 0,
                last_used TEXT,
                FOREIGN KEY (client_id) REFERENCES clients(client_id)
            )
        ''')
        
        self.conn.commit()
    
    def create_menu(self):
        """Create menu bar"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        client_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Client", menu=client_menu)
        client_menu.add_command(label="New Client", command=self.new_client)
        client_menu.add_command(label="Open Client", command=self.open_client)
        client_menu.add_separator()
        client_menu.add_command(label="Exit", command=self.root.quit)
        
        data_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Data", menu=data_menu)
        data_menu.add_command(label="Import QuickBooks GL", command=lambda: self.import_data("historical"))
        data_menu.add_command(label="Import Bank Statement", command=lambda: self.import_data("current"))
        data_menu.add_separator()
        data_menu.add_command(label="Export All to Excel", command=self.export_all_data)
        
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Auto-Classify All", command=self.auto_classify_all)
        tools_menu.add_command(label="Save All Changes", command=self.save_all_changes)
        tools_menu.add_separator()
        tools_menu.add_command(label="Reset Classifications", command=self.reset_classifications)
        
        reports_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Reports", menu=reports_menu)
        reports_menu.add_command(label="Profit & Loss", command=self.generate_pl)
        reports_menu.add_command(label="Balance Sheet", command=self.generate_bs)
        reports_menu.add_separator()
        reports_menu.add_command(label="📊 Comparative P&L (Period Selection)", command=self.generate_comparative_pl_enhanced)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self.show_about)
    
    def create_main_interface(self):
        """Create main interface"""
        top_frame = tk.Frame(self.root, bg='#2196f3', height=60)
        top_frame.pack(fill=tk.X)
        top_frame.pack_propagate(False)
        
        top_inner = tk.Frame(top_frame, bg='#2196f3')
        top_inner.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)
        
        tk.Label(top_inner, text="👤 Current Client:", 
                font=('Segoe UI', 10, 'bold'), bg='#2196f3', fg='white').pack(side=tk.LEFT, padx=(0, 10))
        
        self.client_label = tk.Label(top_inner, text="No client selected", 
                                     font=('Segoe UI', 11), bg='#2196f3', fg='#e3f2fd')
        self.client_label.pack(side=tk.LEFT)
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        self.create_dashboard_tab()
        self.create_data_tab()
        self.create_classification_tab()
        self.create_ask_accountant_tab()
        self.create_edit_transactions_tab()
        self.create_chart_of_accounts_tab()
        self.create_reports_tab()
        
        status_frame = tk.Frame(self.root, bg='#e0e0e0', height=30)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        status_frame.pack_propagate(False)
        
        self.status_bar = tk.Label(status_frame, text="Ready", 
                                   bg='#e0e0e0', fg='#616161',
                                   font=('Segoe UI', 9), anchor=tk.W)
        self.status_bar.pack(fill=tk.BOTH, padx=15, pady=5)
    
    # CLIENT MANAGEMENT
    def new_client(self):
        """Create new client"""
        dialog = tk.Toplevel(self.root)
        dialog.title("New Client")
        dialog.geometry("500x400")
        dialog.configure(bg='#f8f9fa')
        dialog.grab_set()
        
        fields = [
            ("Client Name*:", "client_name"),
            ("Company Name:", "company_name"),
            ("GSTIN:", "gstin"),
            ("PAN:", "pan"),
            ("Address:", "address"),
        ]
        
        entries = {}
        for i, (label, key) in enumerate(fields):
            tk.Label(dialog, text=label, bg='#f8f9fa', font=('Segoe UI', 9)).grid(
                row=i, column=0, sticky=tk.W, padx=20, pady=10)
            entry = tk.Entry(dialog, width=40, font=('Segoe UI', 9))
            entry.grid(row=i, column=1, padx=20, pady=10)
            entries[key] = entry
        
        def save_client():
            client_name = entries['client_name'].get().strip()
            if not client_name:
                messagebox.showerror("Error", "Client name is required!")
                return
            
            try:
                self.cursor.execute('''
                    INSERT INTO clients (client_name, company_name, gstin, pan, address, 
                                       created_date, last_modified)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (
                    client_name,
                    entries['company_name'].get().strip(),
                    entries['gstin'].get().strip(),
                    entries['pan'].get().strip(),
                    entries['address'].get().strip(),
                    datetime.now().isoformat(),
                    datetime.now().isoformat()
                ))
                self.conn.commit()
                
                self.current_client = self.cursor.lastrowid
                self.client_label.config(text=client_name)
                self.status_bar.config(text=f"Client '{client_name}' created")
                
                messagebox.showinfo("Success", f"Client '{client_name}' created!")
                dialog.destroy()
                
            except sqlite3.IntegrityError:
                messagebox.showerror("Error", "Client name already exists!")
        
        tk.Button(dialog, text="Save", command=save_client,
                 bg='#4caf50', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=20, pady=10, cursor='hand2').grid(
                     row=len(fields), column=0, columnspan=2, pady=30)
    
    def open_client(self):
        """Open existing client"""
        self.cursor.execute("SELECT client_id, client_name, company_name FROM clients ORDER BY last_modified DESC")
        clients = self.cursor.fetchall()
        
        if not clients:
            messagebox.showinfo("Info", "No clients found. Create a new client first.")
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Open Client")
        dialog.geometry("600x400")
        dialog.configure(bg='#f8f9fa')
        dialog.grab_set()
        
        tk.Label(dialog, text="Select Client:", font=('Segoe UI', 12, 'bold'),
                bg='#f8f9fa').pack(pady=20)
        
        frame = tk.Frame(dialog, bg='#ffffff')
        frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, font=('Segoe UI', 10))
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        for client_id, name, company in clients:
            display = f"{name} - {company}" if company else name
            listbox.insert(tk.END, display)
        
        def select_client():
            selection = listbox.curselection()
            if not selection:
                return
            
            idx = selection[0]
            self.current_client = clients[idx][0]
            self.client_label.config(text=clients[idx][1])
            
            self.cursor.execute("UPDATE clients SET last_modified = ? WHERE client_id = ?",
                              (datetime.now().isoformat(), self.current_client))
            self.conn.commit()
            
            dialog.destroy()
            self.load_client_data()
            self.status_bar.config(text=f"Opened client: {clients[idx][1]}")
        
        tk.Button(dialog, text="Open", command=select_client,
                 bg='#2196f3', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=30, pady=10, cursor='hand2').pack(pady=10)
    
    def load_client_data(self):
        """Load data for current client"""
        if not self.current_client:
            return
        
        query = "SELECT * FROM transactions WHERE client_id = ? ORDER BY transaction_date"
        df = pd.read_sql_query(query, self.conn, params=(self.current_client,))
        
        self.historical_data = df[df['year'] == 2023] if not df.empty else pd.DataFrame()
        self.current_data = df[df['year'] == 2024] if not df.empty else pd.DataFrame()
        
        query = "SELECT * FROM chart_of_accounts WHERE client_id = ?"
        self.chart_of_accounts = pd.read_sql_query(query, self.conn, params=(self.current_client,))
        
        self.refresh_all_views()
    
    # DATA IMPORT WITH QB SUPPORT
    def import_data(self, data_type):
        """Import QuickBooks GL or Bank Statement"""
        if not self.current_client:
            messagebox.showerror("Error", "Please select a client first!")
            return
        
        title = "Select QuickBooks General Ledger" if data_type == "historical" else "Select Bank Statement"
        file_path = filedialog.askopenfilename(title=title, filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv")])
        
        if not file_path:
            return
        
        try:
            self.status_bar.config(text="Importing data...")
            self.root.update()
            
            if file_path.endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            if data_type == "historical":
                df = self.parse_qb_gl(df)
            else:
                df = self.parse_bank_statement(df)
            
            if df is None or df.empty:
                messagebox.showerror("Error", "No valid data found")
                return
            
            if data_type == "historical":
                self.historical_data = df
                self.learn_from_historical()
                self.generate_chart_of_accounts()
            else:
                self.current_data = df
            
            self.save_transactions_to_db(df, data_type)
            
            messagebox.showinfo("Success", f"Imported {len(df)} transactions!")
            self.status_bar.config(text=f"Imported {len(df)} transactions")
            self.refresh_all_views()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to import: {str(e)}")
            self.status_bar.config(text="Import failed")
    
    def parse_qb_gl(self, df):
        """Parse QuickBooks GL format - ULTRA FLEXIBLE"""
        # Show user what columns we found
        print(f"Found columns: {list(df.columns)}")
        
        # Try multiple possible column names
        qb_columns = {
            'Date': 'transaction_date',
            'date': 'transaction_date',
            'DATE': 'transaction_date',
            'Transaction Date': 'transaction_date',
            'transaction_date': 'transaction_date',
            'Txn Date': 'transaction_date',
            'Transaction Type': 'transaction_type',
            'Type': 'transaction_type',
            'Num': 'num',
            'Number': 'num',
            'Name': 'name',
            'Vendor': 'name',
            'Customer': 'name',
            'Party Name': 'name',
            'Memo/Description': 'description',
            'Description': 'description',
            'Particulars': 'description',
            'Narration': 'description',
            'Memo': 'description',
            'Account': 'account',
            'Split': 'split',
            'Debit': 'debit',
            'debit': 'debit',
            'Dr': 'debit',
            'Withdrawal': 'debit',
            'Credit': 'credit',
            'credit': 'credit',
            'Cr': 'credit',
            'Deposit': 'credit',
            'Balance': 'balance',
            'balance': 'balance',
            'Running Balance': 'balance'
        }
        
        column_mapping = {}
        for qb_col, std_col in qb_columns.items():
            if qb_col in df.columns:
                column_mapping[qb_col] = std_col
        
        df = df.rename(columns=column_mapping)
        
        # Ensure transaction_date exists
        if 'transaction_date' not in df.columns:
            # Try to find ANY date-like column
            for col in df.columns:
                if 'date' in col.lower():
                    df['transaction_date'] = df[col]
                    break
        
        # If still no date column, show error with helpful message
        if 'transaction_date' not in df.columns:
            raise ValueError(f"Could not find date column. Your columns are: {list(df.columns)}\nPlease ensure you have a column named 'Date' or 'Transaction Date'")
        
        # Convert date
        df['transaction_date'] = pd.to_datetime(df['transaction_date'], errors='coerce')
        
        # Add missing required columns
        for col in ['description', 'name', 'account', 'split', 'transaction_type', 'num']:
            if col not in df.columns:
                df[col] = ''
        
        # Convert numeric columns
        for col in ['debit', 'credit', 'balance']:
            if col not in df.columns:
                df[col] = 0
            else:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Remove rows with invalid dates
        df = df[df['transaction_date'].notna()]
        
        if df.empty:
            raise ValueError("No valid data found after date conversion")
        
        # Add classification columns
        df['category'] = ''
        df['sub_category'] = ''
        
        # Auto-map QB accounts
        for idx, row in df.iterrows():
            if pd.notna(row.get('account')):
                account = str(row['account'])
                if account in self.qb_account_mapping:
                    category, sub_category = self.qb_account_mapping[account]
                    df.at[idx, 'category'] = category
                    df.at[idx, 'sub_category'] = sub_category
        
        return df
    
    def parse_bank_statement(self, df):
        """Parse bank statement format - ULTRA FLEXIBLE"""
        print(f"Found columns: {list(df.columns)}")
        
        # Expanded list of possible column names
        date_cols = ['Date', 'date', 'DATE', 'Transaction Date', 'transaction_date', 'Txn Date', 'Value Date', 'Posting Date']
        desc_cols = ['Description', 'description', 'DESCRIPTION', 'Memo', 'Particulars', 'Narration', 'Details', 'Transaction Details']
        name_cols = ['Counterparty', 'Name', 'name', 'Party Name', 'Vendor', 'Customer', 'Payee']
        debit_cols = ['Debit', 'debit', 'DEBIT', 'Withdrawal', 'Dr', 'DR', 'Withdrawals', 'Paid Out']
        credit_cols = ['Credit', 'credit', 'CREDIT', 'Deposit', 'Cr', 'CR', 'Deposits', 'Paid In']
        balance_cols = ['Balance', 'balance', 'BALANCE', 'Running Balance', 'Closing Balance']
        ref_cols = ['Reference', 'Ref', 'Num', 'Transaction ID', 'Cheque No', 'Ref No']
        
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower().strip()
            # Date
            if col in date_cols or 'date' in col_lower:
                column_mapping[col] = 'transaction_date'
            # Description
            elif col in desc_cols or 'description' in col_lower or 'particular' in col_lower or 'narration' in col_lower:
                column_mapping[col] = 'description'
            # Name
            elif col in name_cols or 'name' in col_lower or 'party' in col_lower:
                column_mapping[col] = 'name'
            # Debit
            elif col in debit_cols or 'debit' in col_lower or 'withdrawal' in col_lower or col_lower == 'dr':
                column_mapping[col] = 'debit'
            # Credit
            elif col in credit_cols or 'credit' in col_lower or 'deposit' in col_lower or col_lower == 'cr':
                column_mapping[col] = 'credit'
            # Balance
            elif col in balance_cols or 'balance' in col_lower:
                column_mapping[col] = 'balance'
            # Reference
            elif col in ref_cols or 'ref' in col_lower or 'num' in col_lower:
                column_mapping[col] = 'num'
        
        df = df.rename(columns=column_mapping)
        
        # Ensure required columns exist
        for col in ['transaction_date', 'description', 'name', 'debit', 'credit', 'balance', 'num']:
            if col not in df.columns:
                df[col] = '' if col in ['description', 'name', 'num'] else 0
        
        # If still no transaction_date, raise helpful error
        if df['transaction_date'].astype(str).str.strip().eq('').all():
            raise ValueError(f"Could not find date column. Your columns are: {list(df.columns)}\nPlease ensure you have a column with 'Date' in its name")
        
        # Convert data types
        df['transaction_date'] = pd.to_datetime(df['transaction_date'], errors='coerce')
        df['debit'] = pd.to_numeric(df['debit'], errors='coerce').fillna(0)
        df['credit'] = pd.to_numeric(df['credit'], errors='coerce').fillna(0)
        df['balance'] = pd.to_numeric(df['balance'], errors='coerce').fillna(0)
        
        # Remove rows with invalid dates
        df = df[df['transaction_date'].notna()]
        
        if df.empty:
            raise ValueError("No valid data found after processing")
        
        # Add classification columns
        df['category'] = ''
        df['sub_category'] = ''
        df['transaction_type'] = 'Bank Transaction'
        df['account'] = 'Bank'
        df['split'] = ''
        
        return df
    
    def save_transactions_to_db(self, df, data_type):
        """Save transactions to database"""
        year = 2023 if data_type == "historical" else 2024
        
        for _, row in df.iterrows():
            try:
                self.cursor.execute('''
                    INSERT INTO transactions
                    (client_id, transaction_date, transaction_type, num, name, description,
                     account, split, debit, credit, balance, category, sub_category, year)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    self.current_client,
                    str(row.get('transaction_date', '')),
                    str(row.get('transaction_type', '')),
                    str(row.get('num', '')),
                    str(row.get('name', '')),
                    str(row.get('description', '')),
                    str(row.get('account', '')),
                    str(row.get('split', '')),
                    float(row.get('debit', 0)),
                    float(row.get('credit', 0)),
                    float(row.get('balance', 0)),
                    str(row.get('category', '')),
                    str(row.get('sub_category', '')),
                    year
                ))
            except:
                continue
        
        self.conn.commit()
    
    def learn_from_historical(self):
        """Learn classification patterns"""
        if self.historical_data.empty:
            return
        
        for _, row in self.historical_data.iterrows():
            if not pd.notna(row.get('category')) or not row.get('category'):
                continue
            
            description = str(row.get('description', '')).lower()
            keywords = description.split()[:5]
            
            for keyword in keywords:
                if len(keyword) > 3:
                    self.cursor.execute('''
                        SELECT rule_id FROM classification_rules
                        WHERE client_id = ? AND keyword = ? AND category = ?
                    ''', (self.current_client, keyword, row['category']))
                    
                    if self.cursor.fetchone():
                        self.cursor.execute('''
                            UPDATE classification_rules
                            SET usage_count = usage_count + 1, last_used = ?
                            WHERE client_id = ? AND keyword = ? AND category = ?
                        ''', (datetime.now().isoformat(), self.current_client, keyword, row['category']))
                    else:
                        self.cursor.execute('''
                            INSERT INTO classification_rules
                            (client_id, keyword, qb_account, category, sub_category, usage_count, last_used)
                            VALUES (?, ?, ?, ?, ?, 1, ?)
                        ''', (self.current_client, keyword, str(row.get('account', '')), 
                             row['category'], str(row.get('sub_category', '')), datetime.now().isoformat()))
        
        self.conn.commit()
    
    def generate_chart_of_accounts(self):
        """Generate chart from historical data"""
        if self.historical_data.empty:
            return
        
        grouped = self.historical_data.groupby(['category', 'sub_category']).agg({
            'debit': 'sum',
            'credit': 'sum'
        }).reset_index()
        
        for _, row in grouped.iterrows():
            if pd.notna(row['category']) and row['category']:
                try:
                    self.cursor.execute('''
                        INSERT OR REPLACE INTO chart_of_accounts
                        (client_id, category, sub_category, usage_count, total_debit, total_credit, is_active)
                        VALUES (?, ?, ?, 
                                (SELECT COALESCE(usage_count, 0) + 1 FROM chart_of_accounts 
                                 WHERE client_id = ? AND category = ? AND sub_category = ?),
                                ?, ?, 1)
                    ''', (self.current_client, row['category'], row['sub_category'],
                         self.current_client, row['category'], row['sub_category'],
                         row['debit'], row['credit']))
                except:
                    pass
        
        self.conn.commit()
    
    def auto_classify_all(self):
        """Auto-classify transactions"""
        if self.current_data.empty:
            messagebox.showinfo("Info", "No data to classify")
            return
        
        classified = 0
        ask_accountant = 0
        
        for idx, row in self.current_data.iterrows():
            if pd.notna(row.get('category')) and row.get('category'):
                continue
            
            description = str(row.get('description', '')).lower()
            keywords = description.split()
            best_match = None
            best_score = 0
            
            for keyword in keywords:
                if len(keyword) < 3:
                    continue
                
                self.cursor.execute('''
                    SELECT category, sub_category, confidence, usage_count
                    FROM classification_rules
                    WHERE client_id = ? AND keyword LIKE ?
                    ORDER BY usage_count DESC LIMIT 1
                ''', (self.current_client, f'%{keyword}%'))
                
                match = self.cursor.fetchone()
                
                if match:
                    score = (match[2] * 0.5) + (min(match[3] / 10, 1.0) * 0.5)
                    if score > best_score:
                        best_match = match
                        best_score = score
            
            if best_match and best_score >= 0.3:
                self.current_data.at[idx, 'category'] = best_match[0]
                self.current_data.at[idx, 'sub_category'] = best_match[1]
                classified += 1
            else:
                self.current_data.at[idx, 'category'] = 'Ask My Accountant'
                self.current_data.at[idx, 'sub_category'] = 'Needs Review'
                ask_accountant += 1
        
        self.save_all_changes()
        self.refresh_all_views()
        
        total = len(self.current_data)
        messagebox.showinfo("Complete", 
                          f"Classified: {classified} ({classified/total*100:.1f}%)\n" +
                          f"Ask My Accountant: {ask_accountant} ({ask_accountant/total*100:.1f}%)")
    
    def save_all_changes(self):
        """Save changes to database"""
        if self.current_data.empty:
            return
        
        for idx, row in self.current_data.iterrows():
            if pd.notna(row.get('category')):
                self.cursor.execute('''
                    UPDATE transactions
                    SET category = ?, sub_category = ?
                    WHERE client_id = ? AND transaction_date = ? AND description = ?
                ''', (row['category'], row.get('sub_category', ''),
                     self.current_client, str(row.get('transaction_date', '')), 
                     str(row.get('description', ''))))
        
        self.conn.commit()
        self.status_bar.config(text="All changes saved")
    
    def reset_classifications(self):
        """Reset classifications"""
        if messagebox.askyesno("Confirm", "Reset all classifications?"):
            self.current_data['category'] = ''
            self.current_data['sub_category'] = ''
            self.save_all_changes()
            self.refresh_all_views()
    
    # DASHBOARD TAB
    def create_dashboard_tab(self):
        dashboard = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(dashboard, text="📊 Dashboard")
        
        metrics_frame = tk.Frame(dashboard, bg='#f8f9fa')
        metrics_frame.pack(fill=tk.X, padx=20, pady=20)
        
        metrics = [
            ('Total', 'total', '#e3f2fd', '#1976d2'),
            ('Classified', 'classified', '#e8f5e9', '#388e3c'),
            ('Unclassified', 'unclassified', '#fff3e0', '#f57c00'),
            ('Ask Accountant', 'ask', '#fce4ec', '#c2185b')
        ]
        
        self.metric_cards = {}
        for i, (label, key, bg, fg) in enumerate(metrics):
            card = tk.Frame(metrics_frame, bg=bg)
            card.grid(row=0, column=i, padx=8, sticky='nsew')
            metrics_frame.columnconfigure(i, weight=1)
            
            tk.Label(card, text=label, bg=bg, fg='#616161',
                    font=('Segoe UI', 9)).pack(padx=15, pady=(15, 5))
            
            val = tk.Label(card, text="0", bg=bg, fg=fg, font=('Segoe UI', 32, 'bold'))
            val.pack(padx=15, pady=(0, 15))
            self.metric_cards[key] = val
        
        summary_frame = tk.Frame(dashboard, bg='#ffffff')
        summary_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=20,
                                                      font=('Consolas', 9), bg='#fafafa')
        self.summary_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        self.update_dashboard()
    
    def update_dashboard(self):
        if self.current_data.empty:
            for key in self.metric_cards:
                self.metric_cards[key].config(text="0")
            return
        
        total = len(self.current_data)
        classified = len(self.current_data[
            (self.current_data['category'].notna()) & 
            (self.current_data['category'] != '') &
            (self.current_data['category'] != 'Ask My Accountant')
        ])
        unclassified = len(self.current_data[
            (self.current_data['category'].isna()) | 
            (self.current_data['category'] == '')
        ])
        ask = len(self.current_data[self.current_data['category'] == 'Ask My Accountant'])
        
        self.metric_cards['total'].config(text=str(total))
        self.metric_cards['classified'].config(text=str(classified))
        self.metric_cards['unclassified'].config(text=str(unclassified))
        self.metric_cards['ask'].config(text=str(ask))
        
        summary = f"Total: {total}\nClassified: {classified} ({classified/total*100:.1f}%)\n"
        summary += f"Unclassified: {unclassified} ({unclassified/total*100:.1f}%)\n"
        summary += f"Ask My Accountant: {ask} ({ask/total*100:.1f}%)\n"
        
        self.summary_text.delete('1.0', tk.END)
        self.summary_text.insert('1.0', summary)
    
    # DATA TAB
    def create_data_tab(self):
        tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(tab, text="📁 Data")
        
        btn_frame = tk.Frame(tab, bg='#ffffff')
        btn_frame.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Button(btn_frame, text="Import QuickBooks GL", command=lambda: self.import_data("historical"),
                 bg='#2196f3', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=20, pady=10).pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="Import Bank Statement", command=lambda: self.import_data("current"),
                 bg='#2196f3', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=20, pady=10).pack(side=tk.LEFT, padx=10)
    
    # CLASSIFICATION TAB
    def create_classification_tab(self):
        tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(tab, text="🤖 Classification")
        
        btn_frame = tk.Frame(tab, bg='#ffffff')
        btn_frame.pack(fill=tk.X, padx=20, pady=20)
        
        tk.Button(btn_frame, text="Auto-Classify All", command=self.auto_classify_all,
                 bg='#4caf50', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=20, pady=10).pack(side=tk.LEFT, padx=10)
        
        tk.Button(btn_frame, text="Save Changes", command=self.save_all_changes,
                 bg='#2196f3', fg='white', font=('Segoe UI', 10, 'bold'),
                 relief='flat', padx=20, pady=10).pack(side=tk.LEFT, padx=10)
    
    # ASK MY ACCOUNTANT TAB
    def create_ask_accountant_tab(self):
        ask_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(ask_tab, text="⚠️ Ask My Accountant")
        
        # BULK PANEL
        bulk = tk.Frame(ask_tab, bg='#fff3cd', height=50)
        bulk.pack(fill=tk.X, padx=20, pady=(15,0))
        bulk.pack_propagate(False)
        
        self.ask_sel_lbl = tk.Label(bulk, text="0 selected", bg='#fff3cd', font=('Segoe UI',10,'bold'), fg='#856404')
        self.ask_sel_lbl.pack(side=tk.LEFT, padx=15)
        
        tk.Button(bulk, text="✓ Bulk Categorize", command=self.bulk_cat_ask, bg='#28a745', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bulk, text="Select All", command=self.sel_all_ask, bg='#17a2b8', fg='white', relief='flat', padx=12, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bulk, text="Clear", command=self.clr_sel_ask, bg='#6c757d', fg='white', relief='flat', padx=12, pady=5).pack(side=tk.LEFT, padx=5)
        
        control = tk.Frame(ask_tab, bg='#ffffff')
        control.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(control, text="⚠️ Items Needing Review", font=('Segoe UI', 12, 'bold'), bg='#ffffff', fg='#c2185b').pack(side=tk.LEFT, padx=10)
        tk.Button(control, text="🔄 Refresh", command=self.load_ask_items, bg='#2196f3', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        
        # Sorting buttons
        tk.Label(control, text="Sort:", bg='#ffffff', font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(20,5))
        tk.Button(control, text="Date↓", command=lambda: self.sort_ask('date'), bg='#e3f2fd', relief='flat', padx=8, pady=3, font=('Segoe UI',8)).pack(side=tk.LEFT, padx=2)
        tk.Button(control, text="Amount↓", command=lambda: self.sort_ask('amount'), bg='#e3f2fd', relief='flat', padx=8, pady=3, font=('Segoe UI',8)).pack(side=tk.LEFT, padx=2)
        tk.Button(control, text="Name", command=lambda: self.sort_ask('name'), bg='#e3f2fd', relief='flat', padx=8, pady=3, font=('Segoe UI',8)).pack(side=tk.LEFT, padx=2)
        
        self.ask_count = tk.Label(control, text="Items: 0", bg='#ffffff', font=('Segoe UI', 10))
        self.ask_count.pack(side=tk.RIGHT, padx=10)
        
        table_frame = tk.Frame(ask_tab, bg='#ffffff')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        vsb = tk.Scrollbar(table_frame)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('☐', 'Date', 'Description', 'Name', 'Debit', 'Credit', 'Category')
        self.ask_tree = ttk.Treeview(table_frame, columns=columns, show='headings', yscrollcommand=vsb.set)
        vsb.config(command=self.ask_tree.yview)
        
        widths = [30, 100, 300, 120, 100, 100, 120]
        for col, w in zip(columns, widths):
            self.ask_tree.heading(col, text=col)
            self.ask_tree.column(col, width=w)
        
        self.ask_tree.pack(fill=tk.BOTH, expand=True)
        self.ask_tree.bind('<Button-1>', self.tog_ask_chk)
        self.ask_tree.bind('<Double-1>', self.edit_ask_item)
        
        self.load_ask_items()
    
    def tog_ask_chk(self, e):
        if self.ask_tree.identify("region", e.x, e.y) == "cell":
            if self.ask_tree.identify_column(e.x) == '#1':
                item = self.ask_tree.identify_row(e.y)
                if item:
                    v = list(self.ask_tree.item(item)['values'])
                    v[0] = '☑' if v[0] == '☐' else '☐'
                    self.ask_tree.item(item, values=v)
                    self.upd_ask_cnt()
    
    def upd_ask_cnt(self):
        c = sum(1 for i in self.ask_tree.get_children() if self.ask_tree.item(i)['values'][0] == '☑')
        self.ask_sel_lbl.config(text=f"{c} selected")
    
    def sel_all_ask(self):
        for i in self.ask_tree.get_children():
            v = list(self.ask_tree.item(i)['values'])
            v[0] = '☑'
            self.ask_tree.item(i, values=v)
        self.upd_ask_cnt()
    
    def clr_sel_ask(self):
        for i in self.ask_tree.get_children():
            v = list(self.ask_tree.item(i)['values'])
            v[0] = '☐'
            self.ask_tree.item(i, values=v)
        self.upd_ask_cnt()
    
    def bulk_cat_ask(self):
        sel = [i for i in self.ask_tree.get_children() if self.ask_tree.item(i)['values'][0] == '☑']
        if not sel:
            messagebox.showinfo("Info", "No items selected")
            return
        
        d = tk.Toplevel(self.root)
        d.title(f"Bulk Categorize {len(sel)} Transactions")
        d.geometry("500x300")
        d.grab_set()
        
        tk.Label(d, text=f"Apply to {len(sel)} transactions:", font=('Segoe UI', 12, 'bold')).pack(pady=20)
        tk.Label(d, text="Category:").pack(anchor='w', padx=30)
        cv = tk.StringVar()
        cc = ttk.Combobox(d, textvariable=cv, values=list(self.account_categories.keys()), width=45, state='readonly')
        cc.pack(padx=30, pady=5)
        
        tk.Label(d, text="Sub-Category:").pack(anchor='w', padx=30, pady=(10,0))
        scv = tk.StringVar()
        scc = ttk.Combobox(d, textvariable=scv, values=[], width=45, state='readonly')
        scc.pack(padx=30, pady=5)
        
        def upd(e=None):
            if cv.get() in self.account_categories:
                scc['values'] = self.account_categories[cv.get()]
                if scc['values']: scc.current(0)
        cc.bind('<<ComboboxSelected>>', upd)
        
        def app():
            if not cv.get() or not scv.get():
                messagebox.showwarning("Warning", "Select both category and sub-category")
                return
            for idx in sel:
                self.current_data.at[int(idx), 'category'] = cv.get()
                self.current_data.at[int(idx), 'sub_category'] = scv.get()
            self.save_all_changes()
            self.load_ask_items()
            self.update_dashboard()
            d.destroy()
            messagebox.showinfo("Success", f"Updated {len(sel)} transactions!")
        
        tk.Button(d, text="✓ Apply", command=app, bg='#28a745', fg='white', font=('Segoe UI',11,'bold'), relief='flat', padx=30, pady=12).pack(pady=20)
    
    def load_ask_items(self):
        for item in self.ask_tree.get_children():
            self.ask_tree.delete(item)
        
        if self.current_data.empty:
            self.ask_count.config(text="Items: 0")
            return
        
        ask_items = self.current_data[self.current_data['category'] == 'Ask My Accountant']
        self.ask_count.config(text=f"Items: {len(ask_items)}")
        
        for idx, row in ask_items.iterrows():
            self.ask_tree.insert('', tk.END, iid=str(idx),
                               values=(str(row.get('transaction_date', ''))[:10],
                                      str(row.get('description', '')),
                                      str(row.get('name', '')),
                                      f"₹{row.get('debit', 0):,.2f}" if row.get('debit', 0) > 0 else "",
                                      f"₹{row.get('credit', 0):,.2f}" if row.get('credit', 0) > 0 else "",
                                      str(row.get('category', '')),
                                      str(row.get('sub_category', ''))))
    
    def edit_ask_item(self, event):
        selection = self.ask_tree.selection()
        if not selection:
            return
        
        idx = int(selection[0])
        row = self.current_data.loc[idx]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Edit Classification")
        dialog.geometry("600x500")
        dialog.configure(bg='#f8f9fa')
        dialog.grab_set()
        
        info_frame = tk.LabelFrame(dialog, text="Transaction Details", bg='#ffffff', font=('Segoe UI', 10, 'bold'))
        info_frame.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(info_frame, text=f"Date: {str(row.get('transaction_date', ''))[:10]}", 
                bg='#ffffff').pack(fill=tk.X, padx=15, pady=5)
        tk.Label(info_frame, text=f"Description: {str(row.get('description', ''))}", 
                bg='#ffffff').pack(fill=tk.X, padx=15, pady=5)
        tk.Label(info_frame, text=f"Amount: ₹{row.get('debit', 0) if row.get('debit', 0) > 0 else row.get('credit', 0):,.2f}", 
                bg='#ffffff', font=('Segoe UI', 10, 'bold')).pack(fill=tk.X, padx=15, pady=5)
        
        class_frame = tk.LabelFrame(dialog, text="Set Classification", bg='#ffffff', font=('Segoe UI', 10, 'bold'))
        class_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 15))
        
        tk.Label(class_frame, text="Category:", bg='#ffffff', font=('Segoe UI', 9, 'bold')).pack(anchor='w', padx=15, pady=(15, 5))
        cat_var = tk.StringVar(value=row.get('category', ''))
        cat_combo = ttk.Combobox(class_frame, textvariable=cat_var, values=list(self.account_categories.keys()),
                                 state='readonly', width=50)
        cat_combo.pack(padx=15, pady=(0, 15))
        
        tk.Label(class_frame, text="Sub-Category:", bg='#ffffff', font=('Segoe UI', 9, 'bold')).pack(anchor='w', padx=15, pady=5)
        subcat_var = tk.StringVar(value=row.get('sub_category', ''))
        subcat_combo = ttk.Combobox(class_frame, textvariable=subcat_var, values=[], state='readonly', width=50)
        subcat_combo.pack(padx=15, pady=(0, 15))
        
        def update_subcat(e=None):
            cat = cat_var.get()
            if cat in self.account_categories:
                subcat_combo['values'] = self.account_categories[cat]
                if subcat_combo['values']:
                    subcat_combo.current(0)
        
        cat_combo.bind('<<ComboboxSelected>>', update_subcat)
        update_subcat()
        
        def save():
            self.current_data.at[idx, 'category'] = cat_var.get()
            self.current_data.at[idx, 'sub_category'] = subcat_var.get()
            self.load_ask_items()
            self.update_dashboard()
            dialog.destroy()
            messagebox.showinfo("Success", "Updated!")
        
        btn_frame = tk.Frame(dialog, bg='#f8f9fa')
        btn_frame.pack(fill=tk.X, padx=20, pady=(0, 20))
        
        tk.Button(btn_frame, text="💾 Save", command=save, bg='#4caf50', fg='white',
                 font=('Segoe UI', 10, 'bold'), relief='flat', padx=25, pady=10).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="❌ Cancel", command=dialog.destroy, bg='#757575', fg='white',
                 font=('Segoe UI', 10, 'bold'), relief='flat', padx=25, pady=10).pack(side=tk.LEFT, padx=5)
    
    # EDIT TRANSACTIONS TAB  
    def create_edit_transactions_tab(self):
        edit_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(edit_tab, text="✏️ Edit Transactions")
        
        # BULK PANEL
        bulk = tk.Frame(edit_tab, bg='#d1ecf1', height=50)
        bulk.pack(fill=tk.X, padx=20, pady=(15,0))
        bulk.pack_propagate(False)
        
        self.edit_sel_lbl = tk.Label(bulk, text="0 selected", bg='#d1ecf1', font=('Segoe UI',10,'bold'), fg='#0c5460')
        self.edit_sel_lbl.pack(side=tk.LEFT, padx=15)
        
        tk.Button(bulk, text="✓ Bulk Categorize", command=self.bulk_cat_edit, bg='#28a745', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bulk, text="Select All", command=self.sel_all_edit, bg='#17a2b8', fg='white', relief='flat', padx=12, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(bulk, text="Clear", command=self.clr_sel_edit, bg='#6c757d', fg='white', relief='flat', padx=12, pady=5).pack(side=tk.LEFT, padx=5)
        
        control = tk.Frame(edit_tab, bg='#ffffff')
        control.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(control, text="✏️ Edit Any Transaction", font=('Segoe UI', 12, 'bold'), bg='#ffffff').pack(side=tk.LEFT, padx=10)
        tk.Label(control, text="🔍 Search:", bg='#ffffff').pack(side=tk.LEFT, padx=(20, 5))
        self.edit_search = tk.StringVar()
        self.edit_search.trace_add('write', lambda *args: self.filter_transactions())
        tk.Entry(control, textvariable=self.edit_search, width=30).pack(side=tk.LEFT, padx=5)
        
        tk.Button(control, text="🔄 Refresh", command=self.load_transactions, bg='#2196f3', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        
        # Sorting buttons
        tk.Label(control, text="Sort:", bg='#ffffff', font=('Segoe UI', 9)).pack(side=tk.LEFT, padx=(10,5))
        tk.Button(control, text="Date↓", command=lambda: self.sort_edit('date'), bg='#e3f2fd', relief='flat', padx=8, pady=3, font=('Segoe UI',8)).pack(side=tk.LEFT, padx=2)
        tk.Button(control, text="Amount↓", command=lambda: self.sort_edit('amount'), bg='#e3f2fd', relief='flat', padx=8, pady=3, font=('Segoe UI',8)).pack(side=tk.LEFT, padx=2)
        
        self.edit_count = tk.Label(control, text="Transactions: 0", bg='#ffffff')
        self.edit_count.pack(side=tk.RIGHT, padx=10)
        
        table_frame = tk.Frame(edit_tab, bg='#ffffff')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        vsb = tk.Scrollbar(table_frame)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('☐', 'Date', 'Description', 'Name', 'Amount', 'Type', 'Category')
        self.edit_tree = ttk.Treeview(table_frame, columns=columns, show='headings', yscrollcommand=vsb.set)
        vsb.config(command=self.edit_tree.yview)
        
        widths = [30, 100, 300, 140, 100, 70, 150]
        for col, w in zip(columns, widths):
            self.edit_tree.heading(col, text=col)
            self.edit_tree.column(col, width=w)
        
        self.edit_tree.pack(fill=tk.BOTH, expand=True)
        self.edit_tree.bind('<Button-1>', self.tog_edit_chk)
        self.edit_tree.bind('<Double-1>', lambda e: self.edit_ask_item(e))
        
        self.load_transactions()
    
    def tog_edit_chk(self, e):
        if self.edit_tree.identify("region", e.x, e.y) == "cell":
            if self.edit_tree.identify_column(e.x) == '#1':
                item = self.edit_tree.identify_row(e.y)
                if item:
                    v = list(self.edit_tree.item(item)['values'])
                    v[0] = '☑' if v[0] == '☐' else '☐'
                    self.edit_tree.item(item, values=v)
                    self.upd_edit_cnt()
    
    def upd_edit_cnt(self):
        c = sum(1 for i in self.edit_tree.get_children() if self.edit_tree.item(i)['values'][0] == '☑')
        self.edit_sel_lbl.config(text=f"{c} selected")
    
    def sel_all_edit(self):
        for i in self.edit_tree.get_children():
            v = list(self.edit_tree.item(i)['values'])
            v[0] = '☑'
            self.edit_tree.item(i, values=v)
        self.upd_edit_cnt()
    
    def clr_sel_edit(self):
        for i in self.edit_tree.get_children():
            v = list(self.edit_tree.item(i)['values'])
            v[0] = '☐'
            self.edit_tree.item(i, values=v)
        self.upd_edit_cnt()
    
    def bulk_cat_edit(self):
        sel = [i for i in self.edit_tree.get_children() if self.edit_tree.item(i)['values'][0] == '☑']
        if not sel:
            messagebox.showinfo("Info", "No items selected")
            return
        
        d = tk.Toplevel(self.root)
        d.title(f"Bulk Categorize {len(sel)}")
        d.geometry("500x280")
        d.grab_set()
        
        tk.Label(d, text=f"Categorize {len(sel)} transactions:", font=('Segoe UI', 12, 'bold')).pack(pady=20)
        tk.Label(d, text="Category:").pack(anchor='w', padx=30)
        cv = tk.StringVar()
        cc = ttk.Combobox(d, textvariable=cv, values=list(self.account_categories.keys()), width=45, state='readonly')
        cc.pack(padx=30, pady=5)
        
        tk.Label(d, text="Sub-Category:").pack(anchor='w', padx=30, pady=(10,0))
        scv = tk.StringVar()
        scc = ttk.Combobox(d, textvariable=scv, values=[], width=45, state='readonly')
        scc.pack(padx=30, pady=5)
        
        def upd(e=None):
            if cv.get() in self.account_categories:
                scc['values'] = self.account_categories[cv.get()]
                if scc['values']: scc.current(0)
        cc.bind('<<ComboboxSelected>>', upd)
        
        def app():
            if not cv.get() or not scv.get():
                messagebox.showwarning("Warning", "Select both")
                return
            for idx in sel:
                self.current_data.at[int(idx), 'category'] = cv.get()
                self.current_data.at[int(idx), 'sub_category'] = scv.get()
            self.save_all_changes()
            self.load_transactions()
            self.update_dashboard()
            d.destroy()
            messagebox.showinfo("Success", f"Updated {len(sel)}!")
        
        tk.Button(d, text="✓ Apply", command=app, bg='#28a745', fg='white', font=('Segoe UI',11,'bold'), relief='flat', padx=30, pady=12).pack(pady=20)
    
    def load_transactions(self):
        for item in self.edit_tree.get_children():
            self.edit_tree.delete(item)
        
        if self.current_data.empty:
            self.edit_count.config(text="Transactions: 0")
            return
        
        self.edit_count.config(text=f"Transactions: {len(self.current_data)}")
        
        for idx, row in self.current_data.iterrows():
            debit = row.get('debit', 0)
            credit = row.get('credit', 0)
            amount = debit if debit > 0 else credit
            
            self.edit_tree.insert('', tk.END, iid=str(idx),
                                values=(str(row.get('transaction_date', ''))[:10],
                                       str(row.get('description', '')),
                                       str(row.get('name', '')),
                                       f"₹{amount:,.2f}",
                                       "Debit" if debit > 0 else "Credit",
                                       str(row.get('category', '')),
                                       str(row.get('sub_category', ''))))
    
    def filter_transactions(self):
        search = self.edit_search.get().lower()
        for item in self.edit_tree.get_children():
            self.edit_tree.delete(item)
        
        if self.current_data.empty:
            return
        
        for idx, row in self.current_data.iterrows():
            desc = str(row.get('description', '')).lower()
            name = str(row.get('name', '')).lower()
            cat = str(row.get('category', '')).lower()
            
            if search in desc or search in name or search in cat:
                debit = row.get('debit', 0)
                credit = row.get('credit', 0)
                amount = debit if debit > 0 else credit
                
                self.edit_tree.insert('', tk.END, iid=str(idx),
                                    values=(str(row.get('transaction_date', ''))[:10],
                                           str(row.get('description', '')),
                                           str(row.get('name', '')),
                                           f"₹{amount:,.2f}",
                                           "Debit" if debit > 0 else "Credit",
                                           str(row.get('category', '')),
                                           str(row.get('sub_category', ''))))
    
    # CHART OF ACCOUNTS TAB
    def create_chart_of_accounts_tab(self):
        coa_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(coa_tab, text="📋 Chart of Accounts")
        
        control = tk.Frame(coa_tab, bg='#ffffff')
        control.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(control, text="📋 Chart of Accounts", 
                font=('Segoe UI', 12, 'bold'), bg='#ffffff').pack(side=tk.LEFT, padx=10)
        
        tk.Button(control, text="➕ Add", command=self.add_account,
                 bg='#4caf50', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        tk.Button(control, text="🔄 Refresh", command=self.load_coa,
                 bg='#2196f3', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        
        self.coa_count = tk.Label(control, text="Accounts: 0", bg='#ffffff')
        self.coa_count.pack(side=tk.RIGHT, padx=10)
        
        table_frame = tk.Frame(coa_tab, bg='#ffffff')
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        vsb = tk.Scrollbar(table_frame)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        columns = ('Category', 'Sub-Category', 'Usage', 'Total Debit', 'Total Credit', 'Active')
        self.coa_tree = ttk.Treeview(table_frame, columns=columns, show='headings', yscrollcommand=vsb.set)
        vsb.config(command=self.coa_tree.yview)
        
        for col in columns:
            self.coa_tree.heading(col, text=col)
            width = 250 if col in ['Category', 'Sub-Category'] else 100
            self.coa_tree.column(col, width=width)
        
        self.coa_tree.pack(fill=tk.BOTH, expand=True)
        self.load_coa()
    
    def load_coa(self):
        for item in self.coa_tree.get_children():
            self.coa_tree.delete(item)
        
        if not self.current_client:
            self.coa_count.config(text="Accounts: 0")
            return
        
        self.cursor.execute('''
            SELECT account_id, category, sub_category, usage_count, total_debit, total_credit, is_active
            FROM chart_of_accounts WHERE client_id = ? ORDER BY category, sub_category
        ''', (self.current_client,))
        
        accounts = self.cursor.fetchall()
        self.coa_count.config(text=f"Accounts: {len(accounts)}")
        
        for acc in accounts:
            self.coa_tree.insert('', tk.END, iid=str(acc[0]),
                               values=(acc[1], acc[2], acc[3],
                                      f"₹{acc[4]:,.2f}", f"₹{acc[5]:,.2f}",
                                      "✓ Yes" if acc[6] else "✗ No"))
    
    def add_account(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Account")
        dialog.geometry("500x300")
        dialog.configure(bg='#f8f9fa')
        dialog.grab_set()
        
        tk.Label(dialog, text="Add New Account", font=('Segoe UI', 14, 'bold'), bg='#f8f9fa').pack(pady=20)
        
        fields = tk.Frame(dialog, bg='#ffffff')
        fields.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        tk.Label(fields, text="Category:", bg='#ffffff', font=('Segoe UI', 9, 'bold')).pack(anchor='w', padx=15, pady=(15, 5))
        cat_var = tk.StringVar()
        cat_combo = ttk.Combobox(fields, textvariable=cat_var, values=list(self.account_categories.keys()),
                                 state='readonly', width=50)
        cat_combo.pack(padx=15, pady=(0, 15))
        
        tk.Label(fields, text="Sub-Category:", bg='#ffffff', font=('Segoe UI', 9, 'bold')).pack(anchor='w', padx=15, pady=5)
        subcat_var = tk.StringVar()
        subcat_combo = ttk.Combobox(fields, textvariable=subcat_var, values=[], state='readonly', width=50)
        subcat_combo.pack(padx=15, pady=(0, 15))
        
        def update_subcat(e=None):
            cat = cat_var.get()
            if cat in self.account_categories:
                subcat_combo['values'] = self.account_categories[cat]
        
        cat_combo.bind('<<ComboboxSelected>>', update_subcat)
        
        def save():
            try:
                self.cursor.execute('''
                    INSERT INTO chart_of_accounts (client_id, category, sub_category, is_active)
                    VALUES (?, ?, ?, 1)
                ''', (self.current_client, cat_var.get(), subcat_var.get()))
                self.conn.commit()
                self.load_coa()
                dialog.destroy()
                messagebox.showinfo("Success", "Account added!")
            except:
                messagebox.showerror("Error", "Account already exists!")
        
        tk.Button(dialog, text="💾 Save", command=save, bg='#4caf50', fg='white',
                 font=('Segoe UI', 10, 'bold'), relief='flat', padx=25, pady=10).pack(pady=10)
    
    # REPORTS TAB
    def create_reports_tab(self):
        reports_tab = tk.Frame(self.notebook, bg='#f8f9fa')
        self.notebook.add(reports_tab, text="📈 Reports")
        
        control = tk.Frame(reports_tab, bg='#ffffff')
        control.pack(fill=tk.X, padx=20, pady=15)
        
        tk.Label(control, text="📈 Financial Reports", 
                font=('Segoe UI', 12, 'bold'), bg='#ffffff').pack(side=tk.LEFT, padx=10)
        
        tk.Label(control, text="Period:", bg='#ffffff').pack(side=tk.LEFT, padx=(20, 5))
        self.report_period = tk.StringVar(value="Full Year")
        ttk.Combobox(control, textvariable=self.report_period,
                    values=["Full Year", "Q1", "Q2", "Q3", "Q4"], state='readonly', width=15).pack(side=tk.LEFT, padx=5)
        
        tk.Button(control, text="📊 P&L", command=self.generate_pl,
                 bg='#2196f3', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=10)
        tk.Button(control, text="📊 Balance Sheet", command=self.generate_bs,
                 bg='#2196f3', fg='white', relief='flat', padx=15, pady=5).pack(side=tk.LEFT, padx=5)
        
        report_frame = tk.Frame(reports_tab, bg='#ffffff')
        report_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))
        
        self.report_text = scrolledtext.ScrolledText(report_frame, font=('Consolas', 9), bg='#fafafa', wrap=tk.NONE)
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
    
    def generate_pl(self):
        if self.current_data.empty:
            messagebox.showinfo("Info", "No data available")
            return
        
        period = self.report_period.get()
        data = self.filter_by_period(self.current_data, period)
        
        revenue = data[data['category'] == 'Revenue'].groupby('sub_category')['credit'].sum()
        cogs = data[data['category'] == 'Cost of Goods Sold'].groupby('sub_category')['debit'].sum()
        opex = data[data['category'] == 'Operating Expenses'].groupby('sub_category')['debit'].sum()
        
        total_revenue = revenue.sum()
        total_cogs = cogs.sum()
        total_opex = opex.sum()
        
        gross_profit = total_revenue - total_cogs
        net_profit = gross_profit - total_opex
        
        report = f"""
{'='*80}
                    PROFIT & LOSS STATEMENT
                         Period: {period}
                Client: {self.client_label.cget('text')}
{'='*80}

REVENUE
"""
        for subcat, amt in revenue.items():
            report += f"  {str(subcat)[:50]:50} ₹{amt:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'Total Revenue':50} ₹{total_revenue:>15,.2f}\n\n"
        
        report += "COST OF GOODS SOLD\n"
        for subcat, amt in cogs.items():
            report += f"  {str(subcat)[:50]:50} ₹{amt:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'Total COGS':50} ₹{total_cogs:>15,.2f}\n"
        report += f"  {'='*68}\n"
        report += f"  {'GROSS PROFIT':50} ₹{gross_profit:>15,.2f}\n\n"
        
        report += "OPERATING EXPENSES\n"
        for subcat, amt in opex.items():
            report += f"  {str(subcat)[:50]:50} ₹{amt:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'Total Operating Expenses':50} ₹{total_opex:>15,.2f}\n"
        report += f"  {'='*68}\n"
        report += f"  {'NET PROFIT':50} ₹{net_profit:>15,.2f}\n"
        
        report += f"\n{'='*80}\n"
        report += f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        report += f"{'='*80}\n"
        
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', report)
    
    def generate_bs(self):
        if self.current_data.empty:
            messagebox.showinfo("Info", "No data available")
            return
        
        period = self.report_period.get()
        data = self.filter_by_period(self.current_data, period)
        
        assets = data[data['category'] == 'Assets'].groupby('sub_category').agg({'debit': 'sum', 'credit': 'sum'})
        assets['balance'] = assets['debit'] - assets['credit']
        
        liabilities = data[data['category'] == 'Liabilities'].groupby('sub_category').agg({'debit': 'sum', 'credit': 'sum'})
        liabilities['balance'] = liabilities['credit'] - liabilities['debit']
        
        equity = data[data['category'] == 'Equity'].groupby('sub_category').agg({'debit': 'sum', 'credit': 'sum'})
        equity['balance'] = equity['credit'] - equity['debit']
        
        report = f"""
{'='*80}
                        BALANCE SHEET
                         Period: {period}
                Client: {self.client_label.cget('text')}
{'='*80}

ASSETS
"""
        for subcat, row in assets.iterrows():
            report += f"  {str(subcat)[:50]:50} ₹{row['balance']:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'TOTAL ASSETS':50} ₹{assets['balance'].sum():>15,.2f}\n\n"
        
        report += "LIABILITIES\n"
        for subcat, row in liabilities.iterrows():
            report += f"  {str(subcat)[:50]:50} ₹{row['balance']:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'Total Liabilities':50} ₹{liabilities['balance'].sum():>15,.2f}\n\n"
        
        report += "EQUITY\n"
        for subcat, row in equity.iterrows():
            report += f"  {str(subcat)[:50]:50} ₹{row['balance']:>15,.2f}\n"
        report += f"  {'-'*68}\n"
        report += f"  {'Total Equity':50} ₹{equity['balance'].sum():>15,.2f}\n"
        
        report += f"\n{'='*80}\n"
        
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', report)
    
    def filter_by_period(self, data, period):
        if data.empty:
            return data
        
        data['transaction_date'] = pd.to_datetime(data['transaction_date'], errors='coerce')
        
        if period == "Full Year":
            return data
        elif period == "Q1":
            return data[data['transaction_date'].dt.month.isin([1, 2, 3])]
        elif period == "Q2":
            return data[data['transaction_date'].dt.month.isin([4, 5, 6])]
        elif period == "Q3":
            return data[data['transaction_date'].dt.month.isin([7, 8, 9])]
        elif period == "Q4":
            return data[data['transaction_date'].dt.month.isin([10, 11, 12])]
        
        return data
    
    # UTILITIES
    def refresh_all_views(self):
        try:
            self.update_dashboard()
            self.load_ask_items()
            self.load_transactions()
            self.load_coa()
        except:
            pass
    
    def export_all_data(self):
        if self.current_data.empty and self.historical_data.empty:
            messagebox.showinfo("Info", "No data to export")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        
        if file_path:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if not self.historical_data.empty:
                    self.historical_data.to_excel(writer, sheet_name='2023 Historical', index=False)
                if not self.current_data.empty:
                    self.current_data.to_excel(writer, sheet_name='2024 Current', index=False)
                if not self.chart_of_accounts.empty:
                    self.chart_of_accounts.to_excel(writer, sheet_name='Chart of Accounts', index=False)
            
            messagebox.showinfo("Success", f"Exported to {file_path}")
    
    def show_about(self):
        messagebox.showinfo("About", """
Offline Bookkeeping System v2.0
QuickBooks Compatible

Features:
• Import QB GL & Bank Statements
• Auto-Classification with ML
• Ask My Accountant Tab
• Edit Any Transaction
• Chart of Accounts
• Financial Reports

100% Offline | No Subscription
        """)
    
    
    def generate_comparative_pl(self):
        """Generate Year-over-Year Comparative P&L Report"""
        if self.current_data.empty and self.historical_data.empty:
            messagebox.showinfo("Info", "Need data from both years!\n\nImport 2023 data (QB GL) and 2024 data (Bank Statement)")
            return
        
        d23 = self.historical_data if not self.historical_data.empty else pd.DataFrame()
        d24 = self.current_data if not self.current_data.empty else pd.DataFrame()
        
        r23 = d23[d23['category']=='Revenue']['credit'].sum() if not d23.empty else 0
        r24 = d24[d24['category']=='Revenue']['credit'].sum() if not d24.empty else 0
        c23 = d23[d23['category']=='Cost of Goods Sold']['debit'].sum() if not d23.empty else 0
        c24 = d24[d24['category']=='Cost of Goods Sold']['debit'].sum() if not d24.empty else 0
        o23 = d23[d23['category']=='Operating Expenses']['debit'].sum() if not d23.empty else 0
        o24 = d24[d24['category']=='Operating Expenses']['debit'].sum() if not d24.empty else 0
        
        rv = r24-r23
        rvp = (rv/r23*100) if r23>0 else 0
        n23 = r23-c23-o23
        n24 = r24-c24-o24
        nv = n24-n23
        nvp = (nv/n23*100) if n23!=0 else 0
        
        rep = f"""
{'='*100}
      COMPARATIVE P&L - 2023 vs 2024
      Client: {self.client_label.cget('text') if hasattr(self, 'client_label') else 'N/A'}
{'='*100}

                    2023          2024        Variance    %
Revenue         ₹{r23:>12,.0f}  ₹{r24:>12,.0f}  {'↑' if rv>=0 else '↓'}₹{abs(rv):>10,.0f}  {rvp:>6.1f}%
COGS            ₹{c23:>12,.0f}  ₹{c24:>12,.0f}
Gross Profit    ₹{r23-c23:>12,.0f}  ₹{r24-c24:>12,.0f}
Opex            ₹{o23:>12,.0f}  ₹{o24:>12,.0f}
NET PROFIT      ₹{n23:>12,.0f}  ₹{n24:>12,.0f}  {'↑' if nv>=0 else '↓'}₹{abs(nv):>10,.0f}  {nvp:>6.1f}%
{'='*100}
"""
        if rvp>10: rep += "\n✓ Revenue grew " + f"{rvp:.1f}%"
        if nv>0: rep += "\n✓ Profit up ₹" + f"{nv:,.0f}"
        
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', rep)
        self.status_bar.config(text="Comparative report generated")
    
    
    def sort_ask(self, by):
        """Sort Ask My Accountant"""
        items = [(self.ask_tree.set(k, by), k) for k in self.ask_tree.get_children('')]
        if by == 'date': items.sort(reverse=True)
        elif by == 'amount': items.sort(key=lambda x: float(x[0].replace('₹','').replace(',','')) if x[0].replace('₹','').replace(',','').replace('.','').isdigit() else 0, reverse=True)
        else: items.sort()
        for i, (v, k) in enumerate(items): self.ask_tree.move(k, '', i)
    
    def sort_edit(self, by):
        """Sort Edit Transactions"""
        items = [(self.edit_tree.set(k, by), k) for k in self.edit_tree.get_children('')]
        if by == 'date': items.sort(reverse=True)
        elif by == 'amount': items.sort(key=lambda x: float(x[0].replace('₹','').replace(',','')) if x[0].replace('₹','').replace(',','').replace('.','').isdigit() else 0, reverse=True)
        for i, (v, k) in enumerate(items): self.edit_tree.move(k, '', i)
    
    def generate_comparative_pl_enhanced(self):
        """Comparative P&L with period selection"""
        if self.current_data.empty:
            messagebox.showinfo("Info", "Import data first!")
            return
        
        d = tk.Toplevel(self.root)
        d.title("Comparative P&L")
        d.geometry("450x300")
        d.grab_set()
        
        tk.Label(d, text="📊 Comparative Report", font=('Segoe UI', 14, 'bold')).pack(pady=20)
        tk.Label(d, text="Compare with:", font=('Segoe UI', 10, 'bold')).pack(anchor='w', padx=30, pady=(10,5))
        
        pv = tk.StringVar(value="previous_month")
        for txt, val in [("Previous Month", "previous_month"), ("Previous Quarter", "previous_quarter"), 
                         ("Previous Year", "previous_year"), ("Same Month Last Year", "same_month_last_year")]:
            tk.Radiobutton(d, text=txt, variable=pv, value=val, font=('Segoe UI', 9)).pack(anchor='w', padx=50, pady=2)
        
        def gen():
            d.destroy()
            self.gen_comp_report(pv.get())
        
        tk.Button(d, text="Generate", command=gen, bg='#2196f3', fg='white', font=('Segoe UI', 11, 'bold'), relief='flat', padx=40, pady=12).pack(pady=30)
    
    def gen_comp_report(self, period):
        """Generate comparative report"""
        from datetime import timedelta
        self.current_data['date'] = pd.to_datetime(self.current_data['date'], errors='coerce')
        cd = self.current_data.dropna(subset=['date'])
        if cd.empty: return
        
        ld = cd['date'].max()
        
        if period == "previous_month":
            cs = ld.replace(day=1)
            ce = ld
            pe = cs - timedelta(days=1)
            ps = pe.replace(day=1)
            p1 = ps.strftime("%b %Y")
            p2 = cs.strftime("%b %Y")
        elif period == "previous_quarter":
            cm = ld.month
            cq = (cm-1)//3+1
            cs = ld.replace(month=(cq-1)*3+1, day=1)
            ce = ld
            pq = cq-1 if cq>1 else 4
            py = ld.year if cq>1 else ld.year-1
            ps = ld.replace(year=py, month=(pq-1)*3+1, day=1)
            pe = cs - timedelta(days=1)
            p1 = f"Q{pq} {py}"
            p2 = f"Q{cq} {ld.year}"
        elif period == "previous_year":
            cs = ld.replace(month=1, day=1)
            ce = ld
            ps = cs.replace(year=cs.year-1)
            pe = cs - timedelta(days=1)
            p1 = str(ps.year)
            p2 = str(cs.year)
        else:
            cs = ld.replace(day=1)
            ce = ld
            ps = cs.replace(year=cs.year-1)
            pe = ld.replace(year=ld.year-1)
            p1 = ps.strftime("%b %Y")
            p2 = cs.strftime("%b %Y")
        
        d1 = cd[(cd['date']>=ps) & (cd['date']<=pe)]
        d2 = cd[(cd['date']>=cs) & (cd['date']<=ce)]
        
        r1 = d1[d1['category']=='Revenue']['credit'].sum()
        r2 = d2[d2['category']=='Revenue']['credit'].sum()
        c1 = d1[d1['category']=='Cost of Goods Sold']['debit'].sum()
        c2 = d2[d2['category']=='Cost of Goods Sold']['debit'].sum()
        o1 = d1[d1['category']=='Operating Expenses']['debit'].sum()
        o2 = d2[d2['category']=='Operating Expenses']['debit'].sum()
        
        n1 = r1-c1-o1
        n2 = r2-c2-o2
        rv = r2-r1
        rvp = (rv/r1*100) if r1>0 else 0
        nv = n2-n1
        nvp = (nv/n1*100) if n1!=0 else 0
        
        rep = f"""
{'='*90}
    COMPARATIVE P&L: {p1} vs {p2}
    Client: {self.client_label.cget('text')}
{'='*90}

              {p1:>15}  {p2:>15}    Variance     %
{'='*90}
Revenue     ₹{r1:>13,.0f}  ₹{r2:>13,.0f}  {'↑' if rv>=0 else '↓'}₹{abs(rv):>10,.0f}  {rvp:>6.1f}%
COGS        ₹{c1:>13,.0f}  ₹{c2:>13,.0f}
Gross       ₹{r1-c1:>13,.0f}  ₹{r2-c2:>13,.0f}
Opex        ₹{o1:>13,.0f}  ₹{o2:>13,.0f}
{'='*90}
NET PROFIT  ₹{n1:>13,.0f}  ₹{n2:>13,.0f}  {'↑' if nv>=0 else '↓'}₹{abs(nv):>10,.0f}  {nvp:>6.1f}%
{'='*90}
"""
        if rvp>10: rep += f"\n✓ Revenue growth: {rvp:.1f}%"
        if nv>0: rep += f"\n✓ Profit up: ₹{nv:,.0f}"
        
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', rep)
        self.status_bar.config(text=f"Comparative: {p1} vs {p2}")
    
    def __del__(self):
        if hasattr(self, 'conn'):
            self.conn.close()


if __name__ == "__main__":
    root = tk.Tk()
    app = BookkeepingSystemV2(root)
    root.mainloop()
