# Outlook Inbox Reader

A modern Python application with GUI to read and organize your Microsoft Outlook Inbox by conversations.

![Modern GUI](https://img.shields.io/badge/GUI-CustomTkinter-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-blue)
![Python](https://img.shields.io/badge/Python-3.8+-green)

## ✨ Features

- **Modern GUI**: Sleek, dark/light mode interface built with CustomTkinter
- **Conversation Grouping**: Automatically groups related emails into conversations
- **Real-time Search**: Filter conversations by subject or sender
- **Unread Indicators**: Visual markers for unread emails
- **MVC Architecture**: Clean, maintainable code structure
- **Multi-threaded**: Responsive UI with background data loading
- **One-Click Launch**: Simple batch file to run the application

## 📋 Requirements

- **OS**: Windows (COM interface requires Windows)
- **Python**: 3.8 or higher (Windows Python, not WSL)
- **Outlook**: Microsoft Outlook installed and configured
- **Dependencies**: Automatically installed by batch file
  - `pywin32` - COM interface for Outlook
  - `customtkinter` - Modern GUI framework

## 🚀 Quick Start

### Easy Way (Recommended)

1. **Download the project**
2. **Double-click** `Run_GUI_App.bat`
3. The app will:
   - Auto-install any missing dependencies
   - Launch the modern GUI
   - Connect to your Outlook
   - Load your inbox conversations

That's it! 🎉

### Manual Installation

If you prefer to install dependencies manually:

```cmd
cd "C:\Users\knako\OneDrive\PYTHON PROJECTS\CHIEF OF STAFF NEW"
pip install -r requirements.txt
python app.py
```

## 📱 Application Interface

### GUI Version (Modern)

The main application features:

- **Sidebar Controls**
  - 🔍 Real-time search
  - 🔄 Refresh button
  - 📊 Statistics display
  - 🎨 Dark/Light mode toggle

- **Main Area**
  - Conversation cards with email threads
  - Sender and timestamp info
  - Unread indicators (🔵)
  - Multi-message conversations (💬)

- **Status Bar**
  - Connection status
  - Action feedback
  - Error messages

### CLI Version (Legacy)

For command-line usage, run:

```cmd
Run_Outlook_Reader.bat
```

or

```cmd
python read_outlook_inbox.py
```

This displays a simple list of conversations in the terminal.

## 🏗️ Architecture

The application follows **MVC (Model-View-Controller)** pattern:

```
outlook-inbox-reader/
├── models/
│   ├── __init__.py
│   └── outlook_model.py      # Data layer - Outlook COM interface
├── views/
│   ├── __init__.py
│   └── main_window.py         # UI layer - CustomTkinter GUI
├── app.py                     # Controller - App logic
├── read_outlook_inbox.py      # Legacy CLI script
├── requirements.txt           # Dependencies
├── Run_GUI_App.bat           # GUI launcher
└── Run_Outlook_Reader.bat    # CLI launcher
```

### Model (`models/outlook_model.py`)
- Handles all Outlook COM interactions
- Manages connection state
- Provides conversation grouping logic
- Implements search functionality

### View (`views/main_window.py`)
- Modern CustomTkinter interface
- Dark/Light mode support
- Responsive layout
- Real-time updates

### Controller (`app.py`)
- Coordinates Model and View
- Handles user interactions
- Manages threading for responsiveness
- Error handling and status updates

## 🎯 Usage Examples

### Search for Emails
1. Type in the search box (sidebar)
2. Results filter in real-time
3. Searches both subjects and senders
4. Clear search box to see all conversations

### Refresh Inbox
1. Click "🔄 Refresh Inbox" button
2. App loads latest emails from Outlook
3. Conversations automatically grouped
4. Unread emails highlighted

### Change Theme
1. Use "Appearance" dropdown in sidebar
2. Choose: Dark, Light, or System
3. Theme applies immediately

## 🧪 Testing

Comprehensive test cases are available in `TESTING.md`. The application has been designed with:

- ✅ Comprehensive error handling at all layers (Model, View, Controller)
- ✅ Detailed logging system with full diagnostic info
- ✅ Threading for UI responsiveness
- ✅ Edge case handling (empty inbox, special characters, etc.)
- ✅ Large inbox support (500+ emails)
- ✅ Graceful degradation (skips problematic items, continues processing)
- ✅ Never crashes - all errors caught and logged

## 📝 Error Handling & Logging

The application includes **enterprise-grade error handling** with comprehensive logging:

### Log File

All operations are logged to `outlook_reader.log` with:
- Timestamps for every operation
- Success and error messages
- Full stack traces for debugging
- Connection status and data processing details

### Features

- **Model Layer**: Safe COM property access, message validation, connection error handling
- **View Layer**: Input validation, safe widget operations, type checking
- **Controller Layer**: Thread-safe operations, graceful error recovery

### Documentation

- **`ERROR_HANDLING.md`** - Complete guide to error handling system and log file
- **`TROUBLESHOOTING_CONVERSATIONS.md`** - Step-by-step guide if conversations don't show

### What Gets Logged

✅ Connection attempts and results
✅ Message processing progress
✅ Conversation building
✅ Display operations
✅ Search operations
✅ All errors with stack traces

**Check the log file if anything goes wrong - it contains detailed diagnostic information!**

## 🔧 Troubleshooting

### "Conversations not showing"
**First: Check the log file `outlook_reader.log`**

1. Look for: `"Inbox contains X items"` - How many emails?
2. Look for: `"Processed X/Y messages"` - How many processed?
3. Look for: `"Successfully built X conversations"` - Were conversations created?
4. Look for: `"Displayed X conversations"` - Were they displayed?

**See `TROUBLESHOOTING_CONVERSATIONS.md` for detailed step-by-step diagnostic guide.**

Common causes:
- Empty inbox (log will show "contains 0 items")
- Non-email items in inbox (meetings, tasks - these are skipped)
- Processing errors (check log for error messages)
- Display errors (conversations built but not displayed)

### "Failed to connect to Outlook"
- Ensure Outlook is installed
- Verify Outlook is configured with an email account
- Try running Outlook manually first
- **Check `outlook_reader.log` for detailed error**
- Try running as administrator

### "Module not found" errors
- Run `Run_GUI_App.bat` instead of direct Python
- Or manually: `pip install -r requirements.txt`

### App won't run in WSL
- This app requires **Windows Python**, not WSL Python
- COM interfaces are Windows-specific
- Run from Windows Command Prompt or PowerShell

### General Debugging
**Always check `outlook_reader.log` first!** It contains:
- Detailed error messages
- Stack traces for debugging
- Step-by-step operation log
- Connection and processing details

## 📝 Technical Notes

- **Threading**: Outlook operations run in background threads to prevent UI freezing
- **COM Interface**: Uses `win32com.client` to communicate with Outlook
- **Conversation Grouping**: Uses Outlook's `ConversationID` property
- **UI Updates**: Thread-safe updates using `view.after()` method

## 🤝 Contributing

This project follows best practices:
- Clean MVC architecture
- Type hints for clarity
- Comprehensive error handling
- Documented code

## 📄 License

Free to use and modify for personal or commercial projects.

## 🎨 Screenshots

### Main Interface
- Modern dark theme with sidebar controls
- Conversation cards showing grouped emails
- Real-time search and filtering

### Features
- Unread indicators (blue dot and border)
- Multi-message conversation grouping
- Sender info and timestamps
- Statistics dashboard

---

**Note**: This application requires Windows and Microsoft Outlook. It cannot run on macOS, Linux, or WSL without proper Outlook access.
