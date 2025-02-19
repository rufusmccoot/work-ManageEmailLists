import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import re
import configparser
from openpyxl import load_workbook

class EmailListFreshener:
    def __init__(self):
        """Initialize the application."""
        self.window = tk.Tk()
        self.window.title("Email List Freshener")
        self.window.geometry("800x550")  # Made window wider
        
        # Initialize variables
        self.excel_file_path = tk.StringVar(self.window)
        self.csv_folder_path = tk.StringVar(self.window)
        self.status_var = tk.StringVar(self.window)
        self.progress_var = tk.DoubleVar(self.window)
        
        # Create GUI elements
        self.create_gui()
        
        # Load config and exclusions
        self.config = configparser.ConfigParser()
        self.config.read('configuration.ini')
        self.load_exclusions()
        
        # Set default paths from config
        self.excel_file_path.set(self.config.get('Paths', 'DefaultHostedList', fallback=''))
        self.csv_folder_path.set(self.config.get('Paths', 'DefaultCSVFolder', fallback=''))
        
        # Configure styles
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=5)
        
        # Initialize summary path
        self.summary_path = None

    def load_exclusions(self):
        """Load exclusions from configuration.ini and exclusions.txt"""
        self.excluded_emails = set()
        self.excluded_domains = set()
        
        # Load from exclusions.txt
        try:
            with open('exclusions.txt', 'r') as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith('#'):
                        # Remove any inline comments
                        line = line.split('#')[0].strip()
                        if '@' in line:
                            self.excluded_emails.add(line.lower())
                        else:
                            self.excluded_domains.add(line.lower())
        except Exception as e:
            print(f"Warning: Could not load exclusions.txt: {e}")

    def create_gui(self):
        """Create the GUI elements."""
        # Excel file selection
        excel_frame = ttk.Frame(self.window)
        excel_frame.pack(pady=10, padx=10, fill='x')
        
        ttk.Label(excel_frame, text="Excel File:").pack(side='left')
        self.excel_file_path = tk.StringVar()
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_path, width=80)  # Made entry wider
        excel_entry.pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(excel_frame, text="Browse", command=self.browse_excel).pack(side='left')
        
        # CSV folder selection
        csv_frame = ttk.Frame(self.window)
        csv_frame.pack(pady=10, padx=10, fill='x')
        
        ttk.Label(csv_frame, text="CSV Folder:").pack(side='left')
        self.csv_folder_path = tk.StringVar()
        csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_folder_path, width=80)  # Made entry wider
        csv_entry.pack(side='left', padx=5, fill='x', expand=True)
        ttk.Button(csv_frame, text="Browse", command=self.browse_csv).pack(side='left')
        
        # Process button
        process_button = ttk.Button(self.window, text="Process Files", command=self.process_csvs)
        process_button.pack(pady=10)
        
        # Progress bar - make it 3x longer with a frame
        progress_frame = ttk.Frame(self.window)
        progress_frame.pack(fill='x', padx=10)
        
        # Configure equal weight for the columns to center the progress bar
        progress_frame.columnconfigure(0, weight=1)  # Left spacer
        progress_frame.columnconfigure(2, weight=1)  # Right spacer
        
        # Left spacer
        ttk.Frame(progress_frame).grid(row=0, column=0, sticky='ew')
        
        # Progress bar in center column
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100, length=600)
        self.progress_bar.grid(row=0, column=1)
        
        # Right spacer
        ttk.Frame(progress_frame).grid(row=0, column=2, sticky='ew')
        
        self.progress_bar.grid_remove()  # Hide initially
        
        # Status label - centered
        status_frame = ttk.Frame(self.window)
        status_frame.pack(fill='x')
        ttk.Frame(status_frame).pack(side='left', expand=True)
        status_label = ttk.Label(status_frame, textvariable=self.status_var)
        status_label.pack(side='left')
        ttk.Frame(status_frame).pack(side='left', expand=True)
        
        # Total count label
        total_count_label = ttk.Label(status_frame, textvariable=tk.StringVar(value="Total: 0"))
        total_count_label.pack(side='left')
        
        # Tree view for results
        self.tree = ttk.Treeview(self.window, columns=('Value',), height=15, show='tree')
        self.tree.heading('#0', text='Category')
        self.tree.column('#0', width=200, anchor='w', stretch=False)
        self.tree.column('Value', width=50, anchor='w', stretch=True)
        self.tree.pack(padx=5, pady=5, fill='x')
        
        # Remove grid weights as we're using pack
        # self.window.grid_columnconfigure(0, weight=1)
        # self.window.grid_columnconfigure(1, weight=1)

    def browse_excel(self):
        """Open file dialog to select hosted Excel list."""
        filename = filedialog.askopenfilename(
            title="Select Hosted List Excel File",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if filename:
            self.excel_file_path.set(filename)
            
    def browse_csv(self):
        """Open folder dialog to select CSV folder."""
        folder = filedialog.askdirectory(
            title="Select CSV Folder"
        )
        if folder:
            self.csv_folder_path.set(folder)

    def find_active_column(self, df):
        """Find the column that indicates if a user is active."""
        if 'Status' in df.columns:  # Direct match first
            return 'Status'
            
        # Fallback to pattern matching
        active_patterns = ['active', 'status', 'enabled']
        for col in df.columns:
            if any(pattern in col.lower() for pattern in active_patterns):
                return col
        return None
        
    def is_user_active(self, value):
        """Check if a user is active based on the value in the active column.
        Only 'Active' is considered active, everything else (Deleted, Inactive, etc.) is inactive."""
        if pd.isna(value):
            return False
            
        # Convert to string and check case-insensitive
        str_value = str(value).strip().lower()
        return str_value == 'active'

    def is_excluded(self, email):
        """Check if email is excluded based on exclusions list"""
        email = email.lower()
        if email in self.excluded_emails:
            return True
            
        domain = email.split('@')[1] if '@' in email else ''
        return domain in self.excluded_domains

    def find_email_column(self, df, sheet_name=""):
        """Find the email column in a dataframe"""
        email_column_names = ['email', 'emailaddress', 'e-mail', 'e_mail', 'Email', 'EmailAddress']
        for col in df.columns:
            if str(col).lower().replace(" ", "") in [name.lower().replace(" ", "") for name in email_column_names]:
                print(f"Found email column in {sheet_name}: {col}")
                return col
        return None

    def get_domain_from_email(self, email):
        """Extract domain from email address."""
        try:
            return email.split('@')[1].lower()
        except:
            return None
            
    def find_matching_domain_record(self, domain, todo_df, email_col):
        """Find first record in ToDo sheet with matching domain and return its Company/MailRoom/OCP."""
        if not domain:
            return None
            
        for _, row in todo_df.iterrows():
            if pd.notna(row[email_col]):
                existing_domain = self.get_domain_from_email(str(row[email_col]).lower())
                if existing_domain == domain:
                    return {
                        'Company': row.get('Company', ''),
                        'MailRoom': row.get('MailRoom', ''),
                        'OCP': row.get('OCP', '')
                    }
        return None

    def process_csvs(self):
        """Process all CSVs in the configured folder"""
        try:
            # Get paths from config
            csv_folder = self.csv_folder_path.get()
            excel_path = self.excel_file_path.get()
            
            # Show progress bar
            self.progress_bar.grid()
            self.progress_var.set(0)
            self.status_var.set("Loading files...")
            
            # Initialize stats
            stats = {
                'csv_total': 0,
                'xlsx_initial': 0,
                'added': 0,
                'added_correct_company': 0,
                'added_bad_company': 0,
                'inactive': 0,
                'invalid_format': 0,
                'excluded': 0,
                'already_exists': 0,
                'previously_removed': 0
            }
            
            # Get list of CSV files
            csv_files = [f for f in os.listdir(csv_folder) if f.lower().endswith('.csv')]
            if not csv_files:
                messagebox.showinfo("No Files", "No CSV files found!")
                return
                
            # Load Excel file once
            print("Loading Excel file...")
            todo_df = pd.read_excel(excel_path, sheet_name='To Do')
            removed_df = pd.read_excel(excel_path, sheet_name='ZRemoved')
            
            # Find the email column in the Excel files - case insensitive
            todo_email_col = next((col for col in todo_df.columns if col.lower() == 'email'), None)
            removed_email_col = next((col for col in removed_df.columns if col.lower() == 'email'), None)
            
            if not todo_email_col or not removed_email_col:
                raise ValueError(f"Could not find email columns. To Do columns: {list(todo_df.columns)}, ZRemoved columns: {list(removed_df.columns)}")
            
            # Convert email columns to lowercase for case-insensitive comparison
            todo_emails = set(todo_df[todo_email_col].str.lower().dropna())
            removed_emails = set(removed_df[removed_email_col].str.lower().dropna())
            
            print(f"Loaded {len(todo_emails)} existing emails")
            print(f"Loaded {len(removed_emails)} removed emails")
            
            stats['xlsx_initial'] = len(todo_emails)
            
            # Process each CSV file
            records_to_add = []
            total_processed = 0
            
            for csv_file in csv_files:
                try:
                    full_path = os.path.join(csv_folder, csv_file)
                    print(f"\nProcessing CSV file: {full_path}")
                    df = pd.read_csv(full_path, encoding='utf-8-sig')
                    print(f"CSV columns: {list(df.columns)}")
                    
                    # Find UserLoginId column (this is our email column in CSV)
                    email_col = 'UserLoginId'  # We know this is the correct column name
                    if email_col not in df.columns:
                        print(f"No UserLoginId column found in {csv_file}")
                        continue
                    
                    print(f"Using UserLoginId column for emails")
                    
                    # Find active column
                    active_col = self.find_active_column(df)
                    if active_col:
                        print(f"Found active column: {active_col}")
                    
                    # Process each row
                    for _, row in df.iterrows():
                        stats['csv_total'] += 1
                        total_processed += 1
                        
                        # Update progress every 1000 rows
                        if total_processed % 1000 == 0:
                            self.progress_var.set((total_processed / len(df)) * 100)
                            self.status_var.set(f"Processing row {total_processed} of {len(df)}")
                            self.window.update()
                        
                        # 1. Check Active Status
                        if active_col and not self.is_user_active(row[active_col]):
                            stats['inactive'] += 1
                            continue
                        
                        # 2. Get email and validate format
                        email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
                        if not email or '@' not in email or '.' not in email.split('@')[1]:
                            stats['invalid_format'] += 1
                            continue
                        
                        # Convert email to lowercase for consistency
                        email = email.lower()
                        
                        # 3. Check exclusions
                        if self.is_excluded(email):
                            stats['excluded'] += 1
                            continue
                        
                        # 4. Check if already in current emails
                        if email in todo_emails:
                            stats['already_exists'] += 1
                            continue
                        
                        # 5. Check if in removed emails
                        if email in removed_emails:
                            stats['previously_removed'] += 1
                            continue
                        
                        # 6. Add to records
                        domain = self.get_domain_from_email(email)
                        
                        # Look for matching domain in ToDo sheet
                        matching_record = self.find_matching_domain_record(domain, todo_df, todo_email_col)
                        
                        record = {
                            todo_email_col: row[email_col],  # Original case of email
                            'First Name': str(row['FirstName']).strip() if pd.notna(row.get('FirstName')) else "",
                            'Last Name': str(row['LastName']).strip() if pd.notna(row.get('LastName')) else "",
                            'Extracted from hosted DBs': 'Yes',
                            'Date': datetime.now().strftime('%m/%d/%Y')
                        }
                        
                        if matching_record:
                            # Use values from matching domain record
                            record['Company'] = matching_record['Company']
                            record['MailRoom'] = matching_record['MailRoom']
                            record['OCP'] = matching_record['OCP']
                            stats['added_correct_company'] += 1
                        else:
                            # No matching domain found - use special company format
                            org_name = str(row['OrganizationName']).strip() if pd.notna(row.get('OrganizationName')) else ""
                            record['Company'] = f"zz_EmailListFreshen could not find company based on email address. Tracker PRO Org = {org_name}"
                            record['MailRoom'] = ""
                            record['OCP'] = ""
                            stats['added_bad_company'] += 1
                        
                        # Add empty strings for other columns
                        for col in todo_df.columns:
                            if col not in record:
                                if col == 'Extracted from hosted DBs':
                                    record[col] = 'Yes'
                                elif col == 'Date':
                                    record[col] = datetime.now().strftime('%m/%d/%Y')
                                else:
                                    record[col] = ""
                        
                        records_to_add.append(record)
                        stats['added'] += 1
                
                except Exception as e:
                    print(f"Error processing {csv_file}: {str(e)}")
            
            # Batch add records to Excel
            if records_to_add:
                print(f"Adding {len(records_to_add)} records to Excel...")
                wb = load_workbook(excel_path, data_only=False)
                sheet = wb["To Do"]
                
                # Find last row
                last_row = sheet.max_row
                while last_row > 1:
                    if sheet.cell(row=last_row, column=1).value:
                        break
                    last_row -= 1
                
                # Get column indices
                headers = {sheet.cell(row=1, column=col).value: col 
                          for col in range(1, sheet.max_column + 1)
                          if sheet.cell(row=1, column=col).value}
                print(f"Excel headers: {headers}")
                
                # Add all records at once
                for i, record in enumerate(records_to_add, 1):
                    row_num = last_row + i
                    for header, col_idx in headers.items():
                        sheet.cell(row=row_num, column=col_idx, value=record.get(header, ''))
                
                wb.save(excel_path)
            
            # Display summary
            self.display_summary(stats)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            print(f"Full error details: {str(e)}")
        finally:
            self.progress_bar.grid_remove()
            self.status_var.set("")

    def display_summary(self, stats):
        """Display processing summary in tree view."""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Add main summary node
        parent = self.tree.insert('', 'end', text='Processing Summary', open=True)
        
        # Add counts
        self.tree.insert(parent, 'end', text='Total emails in CSV files', values=(stats['csv_total'],))
        self.tree.insert(parent, 'end', text='Total emails in XLSX', values=(stats['xlsx_initial'],))
        
        # Add records node with subitems
        records = self.tree.insert(parent, 'end', text='Records added', values=(stats['added'],), open=True)
        self.tree.insert(records, 'end', text='Correct company', values=(stats['added_correct_company'],))
        self.tree.insert(records, 'end', text='Bad company', values=(stats['added_bad_company'],))
        
        # Show final count
        self.tree.insert(parent, 'end', text='Final emails in XLSX', values=(stats['xlsx_initial'] + stats['added'],))
        
        # Add skipped records node with total
        total_skipped = (stats['already_exists'] + stats['previously_removed'] + 
                        stats['invalid_format'] + stats['inactive'] + stats['excluded'])
        skipped = self.tree.insert(parent, 'end', text='Skipped Records', values=(total_skipped,), open=True)
        self.tree.insert(skipped, 'end', text='Already exists', values=(stats['already_exists'],))
        self.tree.insert(skipped, 'end', text='Previously removed', values=(stats['previously_removed'],))
        self.tree.insert(skipped, 'end', text='Invalid format', values=(stats['invalid_format'],))
        self.tree.insert(skipped, 'end', text='Inactive users', values=(stats['inactive'],))
        self.tree.insert(skipped, 'end', text='Domain excluded', values=(stats['excluded'],))

    def run(self):
        """Start the application."""
        self.window.mainloop()

if __name__ == "__main__":
    app = EmailListFreshener()
    app.run()
