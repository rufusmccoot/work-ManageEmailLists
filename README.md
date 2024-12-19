# Email List Manager - Tracker PRO Product Team

The Tracker PRO product team uses MS Word and MS Excel mail merge functionality to send client emails. This is a Python application for generating email lists for client notifications.

## Features

- Generate Excel listing of email addresses for targeted client notifications
- Include any combination of:
  - On-Prem
  - Hosted
  - MailRoom
  - OCP
- Automatic deduplication of email addresses
- User-friendly graphical interface
- Configurable last addresses that always appear at the end of the list
    - Allows product team members to receive final emails
    - Provide confidence Outlook processed the entire list and didn't stop in the middle
- "Smart Select" anticipates mailing list combinations based on Word template file name
    - Examples:
      - When "Closing" is part of the Word template file name, both Hosted and On-Prem lists are selected by default
      - When "Hosted" is part of the Word template file name, Hosted list is selected by default
      - Trigger words are configurable in configuration.ini
      - Smart selections can always be overridden

## Requirements

- Python 3.x
- pandas
- python-docx
- tkinter (usually comes with Python)

## Setup

1. Clone the repository:
   ```bash
   cd c:\some_folder
   git clone https://github.com/rufusmccoot/work-ManageEmailLists.git
   ```
2. Create virtual environment:
   ```bash
   python -m venv venv
   ```
3. Activate virtual environment:
   ```bash
   venv\scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
   Or
   ```bash
   pip install --trusted-host pypi.org --trusted-host pypi.python.org --trusted-host files.pythonhosted.org -r requirements.txt
   ```
5. Have a look at `configuration.ini`

## Usage

1. Run the application:
   ```bash
   python app.py
   ```
   or double click the `run_it.bat` file

2. Select your Word document template
3. Choose which mailing lists to include (OnPrem, Hosted, MailRoom, OCP)
4. Click "Process" to generate the combined list
5. Optional - Press the Open buttons at the bottom to quickly open Word doc or generated Excel list

## Configuration

The `configuration.ini` file contains:
- File paths for:
  - OnPrem and Hosted mailing lists
  - Beginning location used when browsing for Word docs
  - Output location where generated mailing list is saved
- File patterns for smart select
- Last addresses to always include at the end of the list
