# Quick Start Guide

## 🚀 Running the Application

**Double-click:** `Run_GUI_App.bat`

That's it! The application will:
1. Install dependencies (first run only)
2. Launch the GUI
3. Connect to Outlook automatically
4. Load your inbox conversations

---

## ❓ Something Not Working?

### Step 1: Check the Log File

Open **`outlook_reader.log`** in the same folder

Look for:
```
Successfully connected to Outlook. Inbox contains X items.
Processed X/X messages successfully
Successfully built X conversation(s)
```

### Step 2: Common Issues

**No conversations showing?**
- Check if your Outlook inbox is empty
- Check if log shows "Inbox contains 0 items"
- See `TROUBLESHOOTING_CONVERSATIONS.md` for detailed help

**Connection error?**
- Make sure Outlook is installed
- Verify Outlook is configured with email account
- Try opening Outlook manually first
- Check `outlook_reader.log` for specific error

**Module not found?**
- Run `Run_GUI_App.bat` (not Python directly)
- It will auto-install missing dependencies

---

## 📚 Documentation

| File | Purpose |
|------|---------|
| **README.md** | Full user guide with features and usage |
| **TROUBLESHOOTING_CONVERSATIONS.md** | Step-by-step fix for "no conversations" issue |
| **ERROR_HANDLING.md** | Complete guide to error system and logs |
| **outlook_reader.log** | Diagnostic log file (check this first!) |

---

## 💡 Tips

✅ The log file (`outlook_reader.log`) is your friend - check it if anything goes wrong

✅ Click "Refresh Inbox" to reload emails

✅ Use the search box to filter by subject or sender

✅ Change theme with the Appearance dropdown

✅ The app never crashes - all errors are caught and logged

---

## 🆘 Still Need Help?

1. Check `outlook_reader.log`
2. Read `TROUBLESHOOTING_CONVERSATIONS.md`
3. Read `ERROR_HANDLING.md`
4. Collect diagnostic info from log
5. Report issue with log excerpt

---

**Most common issue:** Empty inbox or non-email items (meetings/tasks)
**Solution:** Send yourself a test email and click Refresh
