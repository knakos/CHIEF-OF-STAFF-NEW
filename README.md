# Outlook Inbox Reader

A Python script that uses COM interface to read email subjects from Microsoft Outlook Inbox.

## Requirements

- Microsoft Outlook installed and configured
- Python for Windows (not WSL Python)
- pywin32 library

## Installation

Open **Windows Command Prompt** or **PowerShell** (not WSL terminal) and navigate to this directory:

```cmd
cd "C:\Users\knako\OneDrive\PYTHON PROJECTS\CHIEF OF STAFF NEW"
```

Install the required library:

```cmd
pip install -r requirements.txt
```

Or install directly:

```cmd
pip install pywin32
```

## Usage

Run the script using Windows Python:

```cmd
python read_outlook_inbox.py
```

## Output

The script will display a numbered list of all email subjects in your Outlook Inbox, sorted by received time (newest first).

## Notes

- This script must be run using **Windows Python**, not WSL Python, because COM interfaces are Windows-specific
- Outlook must be running or configured on your machine
- The script accesses your default Outlook profile's Inbox folder
