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
        self.window.geometry("600x550")
        
        # Initialize variables
        self.excel_file_path = tk.StringVar(self.window)
        self.csv_folder_path = tk.StringVar(self.window)
        self.status_var = tk.StringVar(self.window)
        self.progress_var = tk.DoubleVar(self.window)
        
        # Create GUI elements
        self.create_gui()
        
        # Load config
        self.config = configparser.ConfigParser()
        self.config.read('configuration.ini')
        
        # Set default paths from config
        self.excel_file_path.set(self.config.get('Paths', 'DefaultHostedList', fallback=''))
        self.csv_folder_path.set(self.config.get('Paths', 'DefaultCSVFolder', fallback=''))
        
        # Configure styles
        style = ttk.Style()
        style.configure('TButton', padding=5)
        style.configure('TLabel', padding=5)
        
        # Initialize summary path
        self.summary_path = None
        
    def create_gui(self):
        """Create the GUI elements."""
        # Excel file selection
        excel_frame = ttk.LabelFrame(self.window, text="Hosted List Excel File", padding=5)
        excel_frame.grid(row=0, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        excel_entry = ttk.Entry(excel_frame, textvariable=self.excel_file_path, width=50)
        excel_entry.grid(row=0, column=0, padx=5)
        
        excel_button = ttk.Button(excel_frame, text="Browse", command=self.browse_excel)
        excel_button.grid(row=0, column=1, padx=5)
        
        # CSV folder selection
        csv_frame = ttk.LabelFrame(self.window, text="CSV Folder", padding=5)
        csv_frame.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        
        csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_folder_path, width=50)
        csv_entry.grid(row=0, column=0, padx=5)
        
        csv_button = ttk.Button(csv_frame, text="Browse", command=self.browse_csv)
        csv_button.grid(row=0, column=1, padx=5)
        
        # Process button
        process_button = ttk.Button(self.window, text="Process Files", command=self.process_csvs)
        process_button.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(self.window, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky="ew")
        self.progress_bar.grid_remove()  # Hide initially
        
        # Status label
        status_label = ttk.Label(self.window, textvariable=self.status_var)
        status_label.grid(row=4, column=0, columnspan=2, pady=5)
        
        # Tree view for results
        self.tree = ttk.Treeview(self.window, columns=('Value',), height=15, show='tree')
        self.tree.column('#0', width=400)  # Make text column wider
        self.tree.column('Value', width=50, anchor='w')  # Left align values
        self.tree.grid(row=5, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # Configure grid weights
        self.window.grid_columnconfigure(0, weight=1)
        self.window.grid_columnconfigure(1, weight=1)
        
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
        """Check if a user is active based on the value in the active column."""
        if pd.isna(value):
            return True
            
        str_value = str(value).lower()
        
        # If value is explicitly "active", user is active
        if str_value == "active":
            return True
            
        # If value is explicitly "inactive", user is inactive
        if str_value == "inactive":
            return False
            
        # Check other common inactive values
        inactive_values = ['disabled', 'false', '0', 'no']
        return not any(inactive in str_value for inactive in inactive_values)
        
    def find_matching_domain_record(self, email, todo_df):
        """Check if email matches any excluded domain or pattern."""
        # Get exclusions from config
        try:
            excluded_domains = [d.strip() for d in self.config.get('EmailExclusions', 'ExcludedDomains').split(',')]
            excluded_patterns = [p.strip() for p in self.config.get('EmailExclusions', 'ExcludedPatterns').split(',')]
        except:
            excluded_domains = []
            excluded_patterns = []
            
        # Check domains
        email_domain = email.split('@')[1] if '@' in email else ''
        if email_domain in excluded_domains:
            return True, f"Matched excluded domain: {email_domain}"
            
        # Check patterns
        for pattern in excluded_patterns:
            if pattern in email:
                return True, f"Matched excluded pattern: {pattern}"
                
        return False, ""
        
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
                'invalid_format': 0,
                'excluded': 0,
                'already_exists': 0,
                'previously_removed': 0,
                'inactive': 0,
                'added': 0,
                'added_with_org': 0,
                'added_without_org': 0
            }
            
            # Get list of CSV files
            csv_files = []
            for file in os.listdir(csv_folder):
                if file.lower().endswith('.csv'):
                    csv_files.append(os.path.join(csv_folder, file))
            
            if not csv_files:
                print("No CSV files found!")
                return
                
            print(f"\nFound {len(csv_files)} CSV files")
            
            # Process each CSV
            records_to_add = []
            total_rows = 0
            processed_rows = 0
            
            # First pass to count total rows
            for csv_file in csv_files:
                df = pd.read_csv(csv_file)
                total_rows += len(df)
            
            # Process each CSV file
            for csv_file in csv_files:
                try:
                    df = pd.read_csv(csv_file, encoding='utf-8-sig')
                    print(f"\nProcessing {csv_file}")
                    print(f"Found {len(df)} rows")
                    print(f"CSV columns: {list(df.columns)}")
                    print("\nFirst few rows of UserLoginId:")
                    if 'UserLoginId' in df.columns:
                        print(df['UserLoginId'].head().to_string())
                    else:
                        print("UserLoginId column not found!")
                        print("Available columns:")
                        for col in df.columns:
                            print(f"  - {col}")
                    
                    # Read target Excel to get columns
                    try:
                        print(f"\nTrying to open Excel file: {excel_path}")
                        todo_df = pd.read_excel(excel_path, sheet_name='To Do')
                        print(f"Successfully opened Excel!")
                        print(f"Target Excel columns: {list(todo_df.columns)}")
                        
                        # Find email column
                        email_col = None
                        for col in df.columns:
                            if col == 'UserLoginId':  # Exact match for UserLoginId
                                email_col = col
                                print(f"Found UserLoginId column: {col}")
                                break
                        
                        if not email_col:
                            print(f"No UserLoginId column found in {csv_file}")
                            print("Available columns:")
                            for col in df.columns:
                                print(f"  - {col}")
                            continue
                        
                        print(f"Using column for email: {email_col}")
                        print("Sample values:")
                        print(df[email_col].head().to_string())
                        
                        # Process each row
                        total_rows = len(df)
                        print(f"\nProcessing {total_rows} rows...")
                        for idx, (_, row) in enumerate(df.iterrows(), 1):
                            if idx % 100 == 0:  # Show progress every 100 rows
                                print(f"Processing row {idx} of {total_rows}")
                                self.progress_var.set(int((idx / total_rows) * 100))
                                self.status_var.set(f"Processing row {idx} of {total_rows}")
                                self.window.update()
                            
                            # Get email and validate
                            email = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
                            stats['csv_total'] += 1  # Count total emails found
                            
                            print(f"Row data for debugging:")
                            print(f"Email column: {email_col}")
                            print(f"Raw email value: {row[email_col]}")
                            print(f"Processed email: {email}")
                            
                            if not email or '@' not in email:
                                print(f"Invalid email format: {email}")
                                stats['invalid_format'] += 1
                                continue
                            
                            # Convert email to lowercase for consistency
                            email = email.lower()
                            print(f"\nProcessing email: {email}")
                            
                            # Skip if already in current emails
                            if email in todo_df['email'].str.lower().dropna():
                                print(f"Skipping - already exists in current emails")
                                stats['already_exists'] += 1
                                continue
                            
                            # Skip if in removed emails
                            try:
                                removed_df = pd.read_excel(excel_path, sheet_name='ZRemoved')
                                if email in removed_df['email'].str.lower().dropna():
                                    print(f"Skipping - found in removed emails")
                                    stats['previously_removed'] += 1
                                    continue
                            except Exception as e:
                                print(f"Error reading ZRemoved sheet: {str(e)}")
                            
                            # Check if user is inactive
                            active_col = self.find_active_column(df)
                            if active_col:
                                status = row[active_col]
                                is_active = self.is_user_active(status)
                                print(f"Status value: {status}, Is active: {is_active}")
                                if not is_active:
                                    print(f"Skipping - user is inactive (status: {status})")
                                    stats['inactive'] += 1
                                    continue
                            
                            # Check exclusions
                            should_exclude, reason = self.find_matching_domain_record(email, todo_df)
                            if should_exclude:
                                print(f"Skipping - {reason}")
                                stats['excluded'] += 1
                                continue
                            
                            print(f"ADDING {email}")
                            
                            # Create record with email from UserLoginId
                            print("\nCreating record:")
                            print(f"UserLoginId from row: {row[email_col]}")
                            
                            record = {}
                            # First set the Email field from UserLoginId
                            record['email'] = str(row[email_col]).strip() if pd.notna(row[email_col]) else ""
                            print(f"Set email to: {record['email']}")
                            
                            # Then set other fields
                            record.update({
                                'Company': str(row['OrganizationName']).strip() if pd.notna(row.get('OrganizationName')) else "",
                                'First Name': str(row['FirstName']).strip() if pd.notna(row.get('FirstName')) else "",
                                'Last Name': str(row['LastName']).strip() if pd.notna(row.get('LastName')) else ""
                            })
                            
                            # Add empty strings for other columns
                            for col in todo_df.columns:
                                if col not in record:
                                    record[col] = ""
                            
                            print("Record created:")
                            for key, value in record.items():
                                print(f"  {key}: {value}")
                            
                            records_to_add.append(record)
                            stats['added'] += 1
                            processed_rows += 1
                        
                        # Update to 100% when done with file
                        self.progress_var.set(100)
                        
                    except Exception as e:
                        print(f"Error reading target Excel: {str(e)}")
                        print(f"Current working directory: {os.getcwd()}")
                        
                except Exception as e:
                    print(f"Error reading CSV: {str(e)}")
                    
            print(f"\nFinished processing. Records to add: {len(records_to_add)}")
            
            # Add new records to Excel
            print(f"\nTotal records collected: {len(records_to_add)}")
            if records_to_add:
                print(f"Adding {len(records_to_add)} records")
                print("First few records:")
                for i, record in enumerate(records_to_add[:3]):
                    print(f"Record {i+1}: {record}")
                
                try:
                    print("\nOpening workbook...")
                    # Load workbook with data_only=False to preserve formulas
                    wb = load_workbook(excel_path, data_only=False)
                    print("Workbook opened successfully")
                    
                    # Get To Do sheet
                    if "To Do" not in wb.sheetnames:
                        print(f"Available sheets: {wb.sheetnames}")
                        print("Could not find 'To Do' sheet!")
                        return
                    
                    sheet = wb["To Do"]
                    print("Found 'To Do' sheet")
                    
                    # Find last row with data by checking Email column
                    last_row = sheet.max_row
                    while last_row > 1:  # Start from bottom, work up, but keep header row
                        if sheet.cell(row=last_row, column=1).value:  # If we find a non-empty email
                            break
                        last_row -= 1
                    print(f"Last row with data: {last_row}")
                    
                    # Get column indices from headers
                    headers = {}
                    for col in range(1, sheet.max_column + 1):
                        header = sheet.cell(row=1, column=col).value
                        if header:
                            headers[header] = col
                    print(f"Found headers: {headers}")
                    
                    # Add each record
                    print(f"\nTotal records to add: {len(records_to_add)}")
                    print("First few records to be added:")
                    for i, record in enumerate(records_to_add[:3]):
                        print(f"\nRecord {i+1}:")
                        for key, value in record.items():
                            print(f"  {key}: {value}")
                    
                    print("\nAdding records...")
                    records_added = 0
                    start_row = last_row + 1  # Start adding after last existing row
                    current_row = start_row
                    
                    print(f"Starting to add records at row {start_row}")
                    for record in records_to_add:
                        if not record.get('email'):  # Skip records without email
                            print(f"Skipping record - no email: {record}")
                            continue
                            
                        print(f"\nWriting record at row {current_row}:")
                        for header, col_idx in headers.items():
                            value = record.get(header, '')
                            print(f"  Writing to column {col_idx} ({header}): {value}")
                            sheet.cell(row=current_row, column=col_idx, value=value)
                        records_added += 1
                        current_row += 1
                    
                    print(f"\nFinished writing {records_added} records")
                    print("First few cells after writing:")
                    for row in range(start_row, min(start_row + 3, current_row)):
                        print(f"\nRow {row}:")
                        for header, col_idx in headers.items():
                            value = sheet.cell(row=row, column=col_idx).value
                            print(f"  {header}: {value}")
                    
                    print(f"\nSaving workbook with {records_added} new records...")
                    wb.save(excel_path)
                    print("Workbook saved successfully")
                    
                except Exception as e:
                    print(f"Error writing to Excel: {str(e)}")
                    import traceback
                    print("Full error:")
                    print(traceback.format_exc())
                    return
            
            # Hide progress bar
            self.progress_bar.grid_remove()
            self.status_var.set("Processing complete")
            
            # Update tree with final statistics
            self.tree.delete(*self.tree.get_children())  # Clear existing items
            
            # Add final statistics
            self.tree.insert("", "end", values=("Total emails in CSV files", f"{stats['csv_total']:,}"))
            self.tree.insert("", "end", values=("", ""))  # Blank line
            
            self.tree.insert("", "end", values=("Already in mailing list", f"{stats['already_exists']:,}"))
            self.tree.insert("", "end", values=("Previously removed", f"{stats['previously_removed']:,}"))
            self.tree.insert("", "end", values=("Inactive users", f"{stats['inactive']:,}"))
            self.tree.insert("", "end", values=("Excluded emails", f"{stats['excluded']:,}"))
            self.tree.insert("", "end", values=("Invalid format", f"{stats['invalid_format']:,}"))
            self.tree.insert("", "end", values=("", ""))  # Blank line
            
            self.tree.insert("", "end", values=("Records added with org", f"{stats['added_with_org']:,}"))
            self.tree.insert("", "end", values=("Records added without org", f"{stats['added_without_org']:,}"))
            self.tree.insert("", "end", values=("", ""))  # Blank line
            
            # Make the total bold
            total_added = stats['added_with_org'] + stats['added_without_org']
            self.tree.insert("", "end", values=(f"Total records added: {total_added:,}", ""))
            
            # Update status
            self.status_var.set("Finished!")
            self.progress_var.set(100)
            
        except Exception as e:
            print(f"Error: {str(e)}")
            self.status_var.set("Error occurred!")
            self.progress_bar.grid_remove()
    
    def run(self):
        """Start the application."""
        self.window.mainloop()

if __name__ == "__main__":
    app = EmailListFreshener()
    app.run()
