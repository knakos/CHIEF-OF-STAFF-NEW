# Error Handling & Debugging Guide

## Overview

This application includes comprehensive error handling and logging throughout all layers (Model, View, Controller). Every operation is wrapped in try-except blocks with detailed logging to help diagnose issues.

## Log File

All application activity is logged to:
```
outlook_reader.log
```

This file is created in the same directory as the application and contains:
- Timestamps for all operations
- Success and error messages
- Stack traces for exceptions
- Debug information for troubleshooting

## Error Handling Features

### Model Layer (`models/outlook_model.py`)

**Features:**
- ✅ Custom exception classes (`OutlookConnectionError`, `OutlookDataError`)
- ✅ Safe property access from COM objects
- ✅ Message class filtering (only processes email messages)
- ✅ Validates data at every step
- ✅ Detailed logging of all operations

**Error Scenarios Handled:**
- Outlook not installed
- Outlook not configured
- Permission issues
- Network/RPC errors
- Missing message properties
- Invalid message types
- Empty inboxes
- Corrupted messages

### View Layer (`views/main_window.py`)

**Features:**
- ✅ Input validation for all display data
- ✅ Safe widget destruction
- ✅ Graceful handling of missing data
- ✅ Type checking for conversations and emails
- ✅ Fallback displays for errors

**Error Scenarios Handled:**
- Invalid conversation data structures
- Missing email properties
- Date formatting errors
- Widget creation failures
- Invalid display counts
- Empty or null data

### Controller Layer (`app.py`)

**Features:**
- ✅ Thread-safe operations
- ✅ Connection error handling
- ✅ Load error recovery
- ✅ Search validation
- ✅ Graceful degradation

**Error Scenarios Handled:**
- Connection failures
- Thread exceptions
- Display errors
- Search failures
- State inconsistencies
- Cleanup errors

## Logging Levels

The application uses these log levels:

| Level | Usage |
|-------|-------|
| **DEBUG** | Detailed diagnostic information (property access, counts, progress) |
| **INFO** | General informational messages (operations starting/completing) |
| **WARNING** | Potentially problematic situations (missing properties, skipped items) |
| **ERROR** | Error events that don't crash the app (individual message failures) |

## Reading the Log File

### Typical Successful Run

```
2025-10-17 12:00:00 - __main__ - INFO - Application starting...
2025-10-17 12:00:00 - models.outlook_model - INFO - OutlookModel initialized
2025-10-17 12:00:01 - models.outlook_model - INFO - Attempting to connect to Outlook...
2025-10-17 12:00:02 - models.outlook_model - INFO - Successfully connected to Outlook. Inbox contains 150 items.
2025-10-17 12:00:03 - models.outlook_model - INFO - Found 150 messages in inbox
2025-10-17 12:00:05 - models.outlook_model - INFO - Processed 150/150 messages successfully (0 errors)
2025-10-17 12:00:05 - models.outlook_model - INFO - Successfully built 87 conversation(s)
2025-10-17 12:00:05 - views.main_window - INFO - Displaying 87 conversations
2025-10-17 12:00:05 - views.main_window - INFO - Displayed 87 conversations successfully (0 errors)
```

### Connection Error

```
2025-10-17 12:00:00 - models.outlook_model - INFO - Attempting to connect to Outlook...
2025-10-17 12:00:01 - models.outlook_model - ERROR - Failed to connect to Outlook: (-2147221005, 'Invalid class string', None, None)
2025-10-17 12:00:01 - models.outlook_model - ERROR - Connection error traceback: [full stack trace]
```

### Message Processing Errors

```
2025-10-17 12:00:03 - models.outlook_model - INFO - Processing 150 messages...
2025-10-17 12:00:03 - models.outlook_model - WARNING - Skipping non-email item (class=45)
2025-10-17 12:00:04 - models.outlook_model - ERROR - Error processing message 75: 'NoneType' object has no attribute 'Subject'
2025-10-17 12:00:05 - models.outlook_model - INFO - Processed 148/150 messages successfully (2 errors)
```

## Debugging Common Issues

### Issue: Conversations Not Showing

**Check the log for:**

1. **Connection Issues**
   ```
   Search for: "Failed to connect"
   Solution: Ensure Outlook is installed and configured
   ```

2. **Empty Inbox**
   ```
   Search for: "Inbox contains 0 items"
   Solution: Verify emails exist in Outlook inbox
   ```

3. **Processing Errors**
   ```
   Search for: "Processed 0/X messages"
   Solution: Check if messages are invalid types
   ```

4. **Display Errors**
   ```
   Search for: "Error creating conversation card"
   Solution: Check data structure in logs
   ```

### Issue: Application Crashes on Startup

**Check the log for:**

```
Search for: "Fatal application error"
```

Common causes:
- Missing dependencies (pywin32, customtkinter)
- Python version incompatibility
- Display/GUI issues

### Issue: Slow Performance

**Check the log for:**

```
Search for: "Processing message X/Y"
```

- Large message counts increase processing time
- Network issues can slow COM operations
- Check time between "Processing..." and "Processed..."

### Issue: Some Emails Missing

**Check the log for:**

```
Search for: "Skipping non-email item"
Search for: "Error processing message"
```

- Non-email items (meetings, tasks) are skipped
- Corrupted messages are skipped
- Check error count in final summary

## Error Recovery Strategies

### Connection Failures

1. **Restart Outlook** - Close and reopen Outlook
2. **Restart Application** - Close and relaunch the app
3. **Check Permissions** - Run as administrator if needed
4. **Review Log** - Check for specific COM errors

### Display Issues

1. **Refresh** - Click the Refresh button
2. **Check Log** - Look for display/widget errors
3. **Restart** - Close and relaunch if UI is corrupted

### Performance Issues

1. **Check Message Count** - Large inboxes take longer
2. **Archive Old Emails** - Reduce inbox size in Outlook
3. **Check Network** - COM operations may be slow on network drives

## Advanced Debugging

### Enable More Verbose Logging

Edit `models/outlook_model.py` line 16:

```python
# Change from INFO to DEBUG
logging.basicConfig(level=logging.DEBUG, ...)
```

This will log every property access and operation.

### Capture Full COM Error Details

Check the log file after an error - it includes:
- Full stack traces
- COM error codes
- Message being processed when error occurred

### Test Connection Separately

Run the CLI version to test Outlook connection:

```cmd
python read_outlook_inbox.py
```

This uses simpler code and can help isolate issues.

## Log File Management

### Location
```
outlook_reader.log (same directory as app.py)
```

### Size Management

The log file grows with use. To manage:

1. **Delete old logs**
   ```cmd
   del outlook_reader.log
   ```

2. **Archive logs**
   ```cmd
   move outlook_reader.log outlook_reader_backup.log
   ```

3. **View recent logs only**
   ```cmd
   powershell "Get-Content outlook_reader.log -Tail 100"
   ```

## Error Messages Reference

### Connection Errors

| Error Message | Meaning | Solution |
|---------------|---------|----------|
| "Invalid class string" | Outlook not installed | Install Microsoft Outlook |
| "Access denied" | Permission issue | Run as administrator |
| "RPC server unavailable" | Outlook not running | Start Outlook |
| "Cannot access inbox items" | Configuration issue | Configure Outlook profile |

### Processing Errors

| Error Message | Meaning | Solution |
|---------------|---------|----------|
| "Skipping non-email item" | Meeting/task/note | Normal - not an email |
| "Cannot get message class" | Corrupted item | Item is skipped automatically |
| "No ConversationID available" | Old Outlook version | Each email shown separately |
| "Invalid count" | Data corruption | Check log for details |

### Display Errors

| Error Message | Meaning | Solution |
|---------------|---------|----------|
| "Invalid conversations type" | Programming error | Report as bug |
| "Error creating card frame" | GUI issue | Restart application |
| "Error formatting timestamp" | Date parsing issue | Item still displays with "Unknown date" |

## Getting Help

If you encounter an error not covered here:

1. **Check the log file** (`outlook_reader.log`)
2. **Find the error** - Search for "ERROR" or "CRITICAL"
3. **Copy the relevant section** - Include timestamps and stack trace
4. **Check for patterns** - Is it happening for all emails or specific ones?
5. **Note your environment**:
   - Python version: `python --version`
   - Outlook version
   - Windows version
   - Number of emails in inbox

## Testing Error Handling

To verify error handling works:

### Test 1: Connection Error
1. Close Outlook completely
2. Run the application
3. Should show "Connection Error" dialog
4. Check log for connection attempt details

### Test 2: Empty Inbox
1. Archive all Outlook emails
2. Run the application
3. Should show "No conversations found"
4. Check log for "Inbox contains 0 items"

### Test 3: Large Inbox
1. Ensure 100+ emails in inbox
2. Run the application
3. Monitor log for progress messages
4. All should load successfully

### Test 4: Search Errors
1. Load conversations
2. Type in search box
3. Check log for search operations
4. Should handle invalid queries gracefully

## Summary

This application is designed to:
- **Never crash** - All errors are caught and logged
- **Show clear errors** - User gets informative dialogs
- **Log everything** - Full diagnostic info in log file
- **Gracefully degrade** - Skip bad items, continue processing
- **Recover automatically** - Most issues can be fixed by refreshing

For any persistent issues, check `outlook_reader.log` for detailed diagnostic information.
