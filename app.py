import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from docx import Document
import configparser
from datetime import datetime
import re
import math

class EmailListManager:
    def __init__(self):
        """
        Initializer for EmailListManager class.
        
        Setup main window and its size, create variables for checkboxes, 
        add trace to checkbox variables to check process button state, 
        configure style for Treeview, and call setup_gui method.
        """
        self.window = tk.Tk()
        self.window.title("EmailFusion")
        self.window.geometry("800x650")  # Made a bit taller
        self.window.iconbitmap("fusion.ico")
        
        # Initialize variables
        self.docx_path = tk.StringVar()
        self.onprem_path = tk.StringVar()
        self.hosted_path = tk.StringVar()
        self.output_path = None
        
        self.onprem_var = tk.BooleanVar()
        self.hosted_var = tk.BooleanVar()
        self.mailroom_var = tk.BooleanVar()
        self.ocp_var = tk.BooleanVar()
        
        # Load configuration
        self.config = configparser.ConfigParser()
        self.config.read('configuration.ini')
        
        # Configure styles
        self.style = ttk.Style()
        self.style.configure("Path.TLabel", foreground="gray")
        
        # Remove the focus dots around selected items
        self.style.layout("Treeview", [('Treeview.treearea', {'sticky': 'nswe'})])
        
        # Create notification window (hidden by default)
        self.notification_window = tk.Toplevel(self.window)
        self.notification_window.withdraw()  # Hide initially
        self.notification_window.overrideredirect(True)  # Remove window decorations
        self.notification_window.configure(background='#FFFDF7')
        
        # Create notification label
        self.notification_label = ttk.Label(
            self.notification_window,
            text="",
            foreground='#666666',
            background='#FFFDF7',
            padding=(10, 5)
        )
        self.notification_label.pack(expand=True)
        
        # Initialize info icons (hidden by default)
        self.info_icons = {}
        
        # Load file patterns from config
        self.file_patterns = {}
        for key in ['mailroom', 'ocp', 'hosted', 'onprem', 'both']:
            if key in self.config['FilePatterns']:
                patterns = [p.strip() for p in self.config['FilePatterns'][key].split(',')]
                self.file_patterns[key] = patterns
        
        self.setup_gui()
    
    def get_sheet_name(self, filename):
        """Get the first sheet name from an Excel file."""
        try:
            xl = pd.ExcelFile(filename)
            return xl.sheet_names[0]
        except Exception as e:
            raise Exception(f"Error reading {filename}: {str(e)}")

    def get_email_column(self, df, filename):
        """Find the email column in the dataframe."""
        # Print available columns for debugging
        self.log_message(f"\nAnalyzing columns in {filename}:")
        self.log_message("=" * 40)
        self.log_message("Found these columns:")
        for idx, col in enumerate(df.columns, 1):
            self.log_message(f"{idx}. [{col}]")
        self.log_message("=" * 40)
        
        possible_names = [
            'Email', 'email', 'Email Address', 'email address', 'EmailAddress', 'emailaddress',
            'E-mail', 'e-mail', 'Mail', 'mail', 'Email_Address', 'email_address'
        ]
        
        # Try exact matches first
        for col in df.columns:
            if str(col).lower().replace(' ', '').replace('-', '').replace('_', '') in [
                name.lower().replace(' ', '').replace('-', '').replace('_', '') 
                for name in possible_names
            ]:
                self.log_message(f"Found email column: [{col}]")
                return col
        
        # If no exact match, look for columns containing 'email' or 'mail'
        for col in df.columns:
            if 'email' in str(col).lower() or 'mail' in str(col).lower():
                self.log_message(f"Found email-like column: [{col}]")
                return col
        
        self.log_message("\nNo email column found!")
        self.log_message("Looking for columns like:")
        for name in possible_names:
            self.log_message(f"  - {name}")
        
        raise ValueError(f"No email column found in {filename}. Please check the column names above.")

    def process_excel(self, file_path, file_type):
        """Process an Excel file and return its email data and full dataframe."""
        try:
            # Read Excel file - always use first column for emails
            df = pd.read_excel(file_path)
            self.log_message(f"Reading {file_type} list...")
            
            # Get emails from first column
            emails = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
            # Filter out empty strings and 'nan'
            emails = [e for e in emails if e and e.lower() != 'nan']
            
            return emails, df
        except Exception as e:
            self.log_message(f"Error processing {file_path}: {str(e)}")
            return [], None

    def browse_docx(self):
        """Browse for DOCX template."""
        file_path = filedialog.askopenfilename(
            filetypes=[("Word Documents", "*.docx")],
            initialdir=self.config['Files']['docx_folder']
        )
        if file_path:
            self.config['Files']['docx_folder'] = os.path.dirname(file_path)
            self.save_config()
            
            self.docx_path.set(file_path)
            self.docx_label.config(text=file_path.replace('/', '\\'))
            
            # Change extension from .docx to .xlsx for output path
            docx_filename = os.path.basename(file_path)
            xlsx_filename = os.path.splitext(docx_filename)[0] + '.xlsx'
            self.output_path = os.path.join(
                self.config['Files']['output_folder'],
                xlsx_filename
            ).replace('/', '\\')
            self.output_label.config(text=self.output_path)
            self.process_button.config(state="normal")
            
            self.check_filename_patterns(os.path.basename(file_path))
    
    def browse_onprem(self):
        """Browse for OnPrem mailing list."""
        file_path = filedialog.askopenfilename(
            title="Select OnPrem Mailing List",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.config['Files']['onprem_list'] = file_path
            self.save_config()
            
            self.onprem_label.config(text=file_path.replace('/', '\\'))
    
    def browse_hosted(self):
        """Browse for Hosted mailing list."""
        file_path = filedialog.askopenfilename(
            title="Select Hosted Mailing List",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if file_path:
            self.config['Files']['hosted_list'] = file_path
            self.save_config()
            
            self.hosted_label.config(text=file_path.replace('/', '\\'))
    
    def browse_output(self):
        """Browse for output folder."""
        folder_path = filedialog.askdirectory(
            initialdir=self.config['Files']['output_folder']
        )
        if folder_path:
            self.config['Files']['output_folder'] = folder_path
            self.save_config()
            
            self.output_path = folder_path.replace('/', '\\')
            self.output_label.config(text=self.output_path)
            self.check_process_button_state()
            
            # If we have a Word doc selected, update the output filename
            if self.docx_path.get():
                docx_name = os.path.basename(self.docx_path.get())
                output_name = os.path.splitext(docx_name)[0] + '.xlsx'
                self.output_path = os.path.join(self.output_path, output_name)
                self.output_label.config(text=self.output_path)
    
    def create_tooltip(self, widget, text):
        """Create a tooltip for a given widget."""
        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 20
            
            # Create a toplevel window
            self.tooltip = tk.Toplevel()
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")
            
            label = ttk.Label(self.tooltip, text=text, justify=tk.LEFT,
                            background="#ffffe0", relief="solid", borderwidth=1)
            label.pack()
            
        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()
                
        widget.bind('<Enter>', enter)
        widget.bind('<Leave>', leave)

    def setup_gui(self):
        """Set up the GUI elements."""
        # File selection frame
        file_frame = ttk.LabelFrame(self.window, text="File Selection", padding=(10, 5))
        file_frame.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # Word Doc
        ttk.Label(file_frame, text="Word Doc:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.docx_label = ttk.Label(file_frame, text="No file selected")
        self.docx_label.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        docx_button = ttk.Button(file_frame, text="[...]", width=3, command=self.browse_docx)
        docx_button.grid(row=0, column=2, padx=5, pady=5)
        
        # OnPrem Mailing List
        ttk.Label(file_frame, text="OnPrem List:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
        self.onprem_label = ttk.Label(file_frame, text=self.config['Files']['onprem_list'].replace('/', '\\'))
        self.onprem_label.grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        onprem_button = ttk.Button(file_frame, text="[...]", width=3, command=self.browse_onprem)
        onprem_button.grid(row=1, column=2, padx=5, pady=5)
        
        # Hosted Mailing List
        ttk.Label(file_frame, text="Hosted List:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        self.hosted_label = ttk.Label(file_frame, text=self.config['Files']['hosted_list'].replace('/', '\\'))
        self.hosted_label.grid(row=2, column=1, padx=5, pady=5, sticky="ew")
        hosted_button = ttk.Button(file_frame, text="[...]", width=3, command=self.browse_hosted)
        hosted_button.grid(row=2, column=2, padx=5, pady=5)
        
        # Generated Mailing List
        ttk.Label(file_frame, text="Generated List:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
        self.output_label = ttk.Label(file_frame, text=self.config['Files']['output_folder'].replace('/', '\\'))
        self.output_label.grid(row=3, column=1, padx=5, pady=5, sticky="ew")
        output_button = ttk.Button(file_frame, text="[...]", width=3, command=self.browse_output)
        output_button.grid(row=3, column=2, padx=5, pady=5)
        
        # Checkboxes frame
        checkbox_frame = ttk.LabelFrame(self.window, text="List Selection", padding=(10, 5))
        checkbox_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # Create a frame for each checkbox and its count
        onprem_group = ttk.Frame(checkbox_frame)
        onprem_group.grid(row=0, column=0, padx=5, pady=5)
        self.onprem_cb = ttk.Checkbutton(onprem_group, text="On-Prem", variable=self.onprem_var)
        self.onprem_cb.grid(row=0, column=0)
        self.count_labels = {}
        self.count_labels['OnPrem'] = ttk.Label(onprem_group, text="-")
        self.count_labels['OnPrem'].grid(row=1, column=0)
        
        hosted_group = ttk.Frame(checkbox_frame)
        hosted_group.grid(row=0, column=1, padx=5, pady=5)
        self.hosted_cb = ttk.Checkbutton(hosted_group, text="Hosted", variable=self.hosted_var)
        self.hosted_cb.grid(row=0, column=0)
        self.count_labels['Hosted'] = ttk.Label(hosted_group, text="-")
        self.count_labels['Hosted'].grid(row=1, column=0)
        
        mailroom_group = ttk.Frame(checkbox_frame)
        mailroom_group.grid(row=0, column=2, padx=5, pady=5)
        self.mailroom_cb = ttk.Checkbutton(mailroom_group, text="MailRoom", variable=self.mailroom_var)
        self.mailroom_cb.grid(row=0, column=0)
        self.count_labels['MailRoom'] = ttk.Label(mailroom_group, text="-")
        self.count_labels['MailRoom'].grid(row=1, column=0)
        
        ocp_group = ttk.Frame(checkbox_frame)
        ocp_group.grid(row=0, column=3, padx=5, pady=5)
        self.ocp_cb = ttk.Checkbutton(ocp_group, text="OCP", variable=self.ocp_var)
        self.ocp_cb.grid(row=0, column=0)
        self.count_labels['OCP'] = ttk.Label(ocp_group, text="-")
        self.count_labels['OCP'].grid(row=1, column=0)
        
        # Process button
        self.process_button = ttk.Button(checkbox_frame, text="Process", command=self.process_data, state="disabled")
        self.process_button.grid(row=0, column=4, padx=15, pady=5)
        
        # Add tooltips
        self.create_tooltip(self.onprem_cb, "Include all emails from the On-Premises list")
        self.create_tooltip(self.hosted_cb, "Include all emails from the Hosted list")
        self.create_tooltip(self.mailroom_cb, "Include emails where Column J is not empty across both lists")
        self.create_tooltip(self.ocp_cb, "Include emails where Column K is not empty across both lists")
        
        # Status frame
        status_frame = ttk.LabelFrame(self.window, text="Status", padding=(10, 5))
        status_frame.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # Status text widget
        self.status_text = tk.Text(status_frame, height=16, width=80)  # Made twice as tall
        self.status_text.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(status_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=0, column=2, sticky="ns")
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Make text widget read-only
        self.status_text.config(state='disabled')
        
        # Buttons frame
        self.buttons_frame = ttk.Frame(status_frame)
        self.buttons_frame.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="w")
        
        # Excel button (hidden by default)
        self.excel_button = ttk.Button(
            self.buttons_frame, 
            text="Open Generated List", 
            command=self.open_excel_file
        )
        
        # Word button (hidden by default)
        self.word_button = ttk.Button(
            self.buttons_frame, 
            text="Open Word Template", 
            command=self.open_word_file
        )
    
    def clear_outputs(self):
        """Clear all output widgets"""
        self.status_text.config(state='normal')
        self.status_text.delete('1.0', tk.END)
        self.status_text.config(state='disabled')
        
        # Reset count labels
        for label in self.count_labels.values():
            label.config(text="-")
        
        # Hide buttons
        self.excel_button.grid_remove()
        self.word_button.grid_remove()
    
    def log_message(self, message):
        """Log a message to the status text widget."""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.config(state='disabled')
        self.status_text.see(tk.END)

    def clean_email(self, email):
        """Clean an email address for comparison."""
        if pd.isna(email):
            return None
        return str(email).lower().strip()

    def process_data(self):
        """Process the selected data and update the display."""
        try:
            # Clear previous outputs
            self.clear_outputs()
            
            # Initialize DataFrames
            onprem_df = None
            hosted_df = None
            combined_df = None
            self.log_message(f"Email list(s)")
            
            # Read OnPrem and Hosted lists if needed
            if self.onprem_var.get() or self.mailroom_var.get() or self.ocp_var.get():
                try:
                    onprem_df = pd.read_excel(self.config['Files']['onprem_list'])
                    self.log_message(f"{'    OnPrem list':.<50}{len(onprem_df):>5}")
                    if self.onprem_var.get():
                        # Extract and clean email column
                        onprem_emails = [self.clean_email(email) for email in onprem_df.iloc[:, 0]]
                        onprem_emails = [email for email in onprem_emails if email is not None]
                        self.count_labels["OnPrem"].config(text=str(len(onprem_emails)))
                        onprem_email_df = pd.DataFrame({'Email': onprem_emails})
                        combined_df = onprem_email_df if combined_df is None else pd.concat([combined_df, onprem_email_df], ignore_index=True)
                except Exception as e:
                    self.log_message(f"Error processing OnPrem list: {str(e)}")
            
            if self.hosted_var.get() or self.mailroom_var.get() or self.ocp_var.get():
                try:
                    hosted_df = pd.read_excel(self.config['Files']['hosted_list'])
                    self.log_message(f"{'    Hosted list':.<50}{len(hosted_df):>5}")
                    if self.hosted_var.get():
                        # Extract and clean email column
                        hosted_emails = [self.clean_email(email) for email in hosted_df.iloc[:, 0]]
                        hosted_emails = [email for email in hosted_emails if email is not None]
                        self.count_labels["Hosted"].config(text=str(len(hosted_emails)))
                        hosted_email_df = pd.DataFrame({'Email': hosted_emails})
                        combined_df = hosted_email_df if combined_df is None else pd.concat([combined_df, hosted_email_df], ignore_index=True)
                except Exception as e:
                    self.log_message(f"Error processing Hosted list: {str(e)}")
            
            # Process Mailroom (Column J)
            if self.mailroom_var.get():
                mailroom_emails = []
                if onprem_df is not None:
                    mailroom_mask = onprem_df.iloc[:, 9].notna()  # Column J
                    mailroom_emails.extend([self.clean_email(email) for email in onprem_df[mailroom_mask].iloc[:, 0]])
                if hosted_df is not None:
                    mailroom_mask = hosted_df.iloc[:, 9].notna()  # Column J
                    mailroom_emails.extend([self.clean_email(email) for email in hosted_df[mailroom_mask].iloc[:, 0]])
                
                mailroom_emails = [email for email in mailroom_emails if email is not None]
                if mailroom_emails:
                    self.count_labels["MailRoom"].config(text=str(len(mailroom_emails)))
                    mailroom_df = pd.DataFrame({'Email': mailroom_emails})
                    combined_df = mailroom_df if combined_df is None else pd.concat([combined_df, mailroom_df], ignore_index=True)
                    self.log_message(f"{'    Mailroom list':.<50}{len(mailroom_emails):>5}")
            
            # Process OCP (Column K)
            if self.ocp_var.get():
                ocp_emails = []
                if onprem_df is not None:
                    ocp_mask = onprem_df.iloc[:, 10].notna()  # Column K
                    ocp_emails.extend([self.clean_email(email) for email in onprem_df[ocp_mask].iloc[:, 0]])
                if hosted_df is not None:
                    ocp_mask = hosted_df.iloc[:, 10].notna()  # Column K
                    ocp_emails.extend([self.clean_email(email) for email in hosted_df[ocp_mask].iloc[:, 0]])
                
                ocp_emails = [email for email in ocp_emails if email is not None]
                if ocp_emails:
                    self.count_labels["OCP"].config(text=str(len(ocp_emails)))
                    ocp_df = pd.DataFrame({'Email': ocp_emails})
                    combined_df = ocp_df if combined_df is None else pd.concat([combined_df, ocp_df], ignore_index=True)
                    self.log_message(f"{'    OCP list':.<50}{len(ocp_emails):>5}")
            
            # Save combined results if we have any data
            if combined_df is not None:
                # Remove duplicates
                total_before = len(combined_df)
                combined_df.drop_duplicates(subset=['Email'], keep='first', inplace=True)
                total_after = len(combined_df)
                duplicates_removed = total_before - total_after

                # Get the last addresses from config
                last_addresses = []
                if 'LastAddressesToSend' in self.config and 'addresslist' in self.config['LastAddressesToSend']:
                    last_addresses = [
                        self.clean_email(email.strip()) 
                        for email in self.config['LastAddressesToSend']['addresslist'].split(',')
                    ]
                    last_addresses = [email for email in last_addresses if email]  # Remove empty entries
                
                if last_addresses:
                    # Remove these addresses from the combined list if they exist
                    last_addresses_already_in_list = combined_df['Email'].isin(last_addresses).sum() # How many dupes are there?
                    combined_df = combined_df[~combined_df['Email'].isin(last_addresses)]            # Example marta@here.com in the config file and already in lists
                    
                    # Add them to the end
                    last_addresses_df = pd.DataFrame({'Email': last_addresses})
                    combined_df = pd.concat([combined_df, last_addresses_df], ignore_index=True)
                    total_last_addresses_added = len(last_addresses_df) - last_addresses_already_in_list # The size of the list in config file minus those already listed

                self.log_message(f"\nReconciliation")
                self.log_message(f"{'    Count before removing dupes':.<50}{total_before:>5}")
                self.log_message(f"{'    Duplicates removed':.<50}{duplicates_removed:>5}")
                self.log_message(f"{'    Addl addresses from config file':.<50}{total_last_addresses_added:>5}")
                self.log_message(f"{'    Final list count':.<50}{len(combined_df):>5}")
                
                # Ensure output directory exists
                output_dir = os.path.dirname(self.output_path)
                os.makedirs(output_dir, exist_ok=True)
                
                # Save the combined DataFrame to Excel
                combined_df.to_excel(self.output_path, index=False)
                self.log_message(f"\nOutput files")
                self.log_message(f"    {os.path.basename(self.output_path)[:45].ljust(46, '.')}{len(combined_df):>5}")
                
                # Now also export as 500-record text files
                docx_filename = os.path.basename(self.docx_path.get()) # Start with the docx filename again
                base_name = os.path.splitext(docx_filename)[0]         # Strip .docx extension
                output_folder = self.config['Files']['output_folder']  # Same output folder as Excel output
                chunk_size = 500
                num_chunks = math.ceil(len(combined_df) / chunk_size)
                
                for i in range(num_chunks):
                    start, end = i * chunk_size, (i + 1) * chunk_size
                    chunk_list = combined_df.iloc[start:end, 0].tolist()
                    # use same base name as xlsx, append -chunkNN.txt
                    txt_filename = f"{base_name}-chunk{i+1:02d}.txt"
                    txt_path = os.path.join(output_folder, txt_filename).replace('/', '\\')
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write('; '.join(chunk_list))
                    self.log_message(f"{f'    Chunk {i+1:02d} txt':.<50}{len(chunk_list):>5}")
                # Update status and show buttons
                self.status_text.config(state='normal')
                self.status_text.config(state='disabled')
                
                # Show both buttons
                self.excel_button.grid(row=0, column=0, padx=(0, 5))
                self.word_button.grid(row=0, column=1, padx=(5, 0))
            else:
                self.log_message("No data to process!")
                
        except Exception as e:
            messagebox.showerror("Processing Error", str(e))
        
        # Update the Open Excel button state at the end
        if self.output_path and os.path.exists(self.output_path):
            self.open_excel_button.config(state="normal")
        else:
            self.open_excel_button.config(state="disabled")
    
    def open_output_file(self, event=None):
        """Open the output Excel file using the default application."""
        if self.output_path and os.path.exists(self.output_path):
            os.startfile(self.output_path)
        else:
            messagebox.showerror("Error", "Output file not found")

    def update_process_button(self, *args):
        """Enable process button if file is selected and at least one checkbox is selected."""
        any_checkbox = (self.onprem_var.get() or 
                       self.hosted_var.get() or 
                       self.mailroom_var.get() or 
                       self.ocp_var.get())
        
        if any_checkbox and self.docx_path.get():
            self.process_button.config(state="normal")
        else:
            self.process_button.config(state="disabled")

    def show_notification(self):
        """Show notification just below the file selection frame."""
        # Get file frame position and size
        file_frame = self.window.children['!labelframe']  # First LabelFrame is our file frame
        x = file_frame.winfo_rootx()
        y = file_frame.winfo_rooty() + file_frame.winfo_height()
        
        # Position notification window
        self.notification_window.geometry(f"+{x}+{y}")
        self.notification_window.deiconify()
        self.notification_window.lift()
        
        # Schedule fade out
        self.window.after(5000, self.fade_notification)
    
    def fade_notification(self):
        """Hide the notification window."""
        self.notification_window.withdraw()
    
    def add_glow(self, checkbox):
        # Add subtle glow effect to checkbox
        checkbox.configure(style='Glow.TCheckbutton')
        # Remove glow after 5 seconds
        self.window.after(5000, lambda: checkbox.configure(style='TCheckbutton'))

    def check_filename_patterns(self, filename):
        """Check filename against patterns and auto-select checkboxes."""
        filename = filename.lower()
        matched_pattern = None
        
        # Check each pattern type
        for pattern_type, patterns in self.file_patterns.items():
            for pattern in patterns:
                if pattern.lower() in filename:
                    matched_pattern = pattern
                    if pattern_type == 'mailroom':
                        self.mailroom_var.set(True)
                    elif pattern_type == 'ocp':
                        self.ocp_var.set(True)
                    elif pattern_type == 'hosted':
                        self.hosted_var.set(True)
                    elif pattern_type == 'onprem':
                        self.onprem_var.set(True)
                    elif pattern_type == 'both':
                        self.hosted_var.set(True)
                        self.onprem_var.set(True)
        
        # Show notification if any pattern matched
        if matched_pattern:
            self.notification_label.config(
                text=f" ✨ Smart selected mailing lists because Word doc contains '{matched_pattern}' - Adjust if needed"
            )
            self.show_notification()
    
    def open_excel_file(self):
        """Open the generated Excel file."""
        if os.path.exists(self.output_path):
            os.startfile(self.output_path)

    def open_word_file(self):
        """Open the Word document template."""
        if self.docx_path.get():
            os.startfile(self.docx_path.get())

    def save_config(self):
        """Save current configuration to file."""
        with open('configuration.ini', 'w') as configfile:
            self.config.write(configfile)

    def run(self):
        self.window.mainloop()

if __name__ == "__main__":
    app = EmailListManager()
    app.run()
