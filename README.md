# Email List Manager - Tracker PRO Product Team

The Tracker PRO product team uses MS Word and MS Excel mail merge functionality to send client emails. This is a Python application for generating email lists for client notifications.

## Features

- Generate listing of email addresses for targeted client notifications without duplicates
- Output:
  - a single combined Excel list for mail merge
  - chunked text files for pasting 500 email addresses into the To line (lol mail goes BRRRRRRRRRR)
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

## Setup

1. Have a configuration.ini file in the same directory as the exe file
2. Maybe look at the config file
3. Double click the thing

## Usage

1. Double click the exe file
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
