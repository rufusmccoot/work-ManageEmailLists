# Email List Manager

A Python application for managing and combining email lists from multiple Excel sources.

## Features

- Combine email lists from OnPrem and Hosted Excel files
- Filter emails based on MailRoom and OCP criteria
- Automatic deduplication of email addresses
- Configurable last addresses that always appear at the end of the list
- User-friendly GUI interface
- Excel output with combined results

## Requirements

- Python 3.x
- pandas
- python-docx
- tkinter (usually comes with Python)

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Configure the paths in `configuration.ini`

## Usage

1. Run the application:
   ```bash
   python app.py
   ```
   or use the `run_it.bat` file

2. Select your Word document template
3. Choose which lists to include (OnPrem, Hosted, MailRoom, OCP)
4. Click "Process" to generate the combined list

## Configuration

The `configuration.ini` file contains:
- File paths for OnPrem and Hosted lists
- Document and output folder locations
- File patterns for automatic checkbox selection
- Last addresses to always include at the end of the list
