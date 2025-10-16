# Testing Guide for Outlook Inbox Reader GUI

## Test Environment

- **Platform**: Windows (COM interface requires Windows)
- **Python**: 3.8 or higher
- **Dependencies**: pywin32, customtkinter
- **Requirements**: Microsoft Outlook installed and configured

## Pre-Test Setup

1. Install dependencies:
   ```cmd
   pip install -r requirements.txt
   ```

2. Verify Outlook is installed and configured with at least one email account

3. Ensure Outlook is running or can be started by Windows

## Test Cases

### 1. Application Launch Tests

#### Test 1.1: Initial Launch
**Steps:**
1. Double-click `Run_GUI_App.bat`
2. Observe application window opens

**Expected Results:**
- ✅ Window opens with title "Outlook Inbox Reader"
- ✅ Sidebar visible with app logo and controls
- ✅ Main area shows "Click 'Refresh Inbox' to load emails"
- ✅ Status bar shows "Connecting to Outlook..."
- ✅ Connection completes successfully

**Pass Criteria:**
- GUI appears with no errors
- All UI elements visible and properly positioned

#### Test 1.2: Dependency Auto-Installation
**Steps:**
1. Uninstall customtkinter: `pip uninstall customtkinter`
2. Run `Run_GUI_App.bat`

**Expected Results:**
- ✅ Batch file detects missing dependency
- ✅ Automatically installs requirements
- ✅ Application launches successfully

---

### 2. Connection Tests

#### Test 2.1: Successful Outlook Connection
**Steps:**
1. Launch application
2. Wait for auto-connection

**Expected Results:**
- ✅ Status shows "Connecting to Outlook..."
- ✅ Connection succeeds within 2-3 seconds
- ✅ Status updates to "Connected to Outlook..."
- ✅ Conversations auto-load

#### Test 2.2: Outlook Not Running
**Steps:**
1. Close Outlook completely
2. Launch application

**Expected Results:**
- ✅ Application attempts connection
- ✅ Either successfully starts Outlook OR shows error dialog
- ✅ Error message is clear and helpful

#### Test 2.3: Outlook Not Installed
**Steps:**
1. Test on system without Outlook

**Expected Results:**
- ✅ Error dialog appears
- ✅ Message explains Outlook is required
- ✅ Application remains stable

---

### 3. Data Loading Tests

#### Test 3.1: Load Empty Inbox
**Steps:**
1. Ensure Outlook inbox is empty
2. Click "Refresh Inbox"

**Expected Results:**
- ✅ Loading indicator shows
- ✅ Completes without errors
- ✅ Shows "No conversations found"
- ✅ Stats show "Conversations: 0, Total Emails: 0"

#### Test 3.2: Load Inbox with Single Emails
**Steps:**
1. Have 5-10 unrelated emails in inbox
2. Click "Refresh Inbox"

**Expected Results:**
- ✅ All emails load and display
- ✅ Each shows as individual conversation
- ✅ Displays sender, subject, and timestamp
- ✅ Stats are accurate

#### Test 3.3: Load Inbox with Conversations
**Steps:**
1. Have email threads (3+ related emails) in inbox
2. Click "Refresh Inbox"

**Expected Results:**
- ✅ Related emails grouped into conversations
- ✅ Shows "💬" icon and count (e.g., "5 messages")
- ✅ Displays last 3 emails in thread
- ✅ Shows "... and X more message(s)" if >3 emails

#### Test 3.4: Load Large Inbox (100+ emails)
**Steps:**
1. Test with inbox containing 100+ emails
2. Click "Refresh Inbox"

**Expected Results:**
- ✅ Loads without freezing UI
- ✅ Completes within 10 seconds
- ✅ All conversations display correctly
- ✅ Scrolling is smooth

---

### 4. UI Interaction Tests

#### Test 4.1: Refresh Button
**Steps:**
1. Load conversations
2. Click "Refresh Inbox" again
3. Send yourself a new test email
4. Click "Refresh Inbox"

**Expected Results:**
- ✅ Button disables during refresh
- ✅ Shows "Loading..." text
- ✅ Re-enables after completion
- ✅ New email appears in list

#### Test 4.2: Search Functionality - Subject
**Steps:**
1. Load conversations
2. Type partial subject in search box (e.g., "meeting")

**Expected Results:**
- ✅ Results filter in real-time
- ✅ Only matching conversations shown
- ✅ Status updates with match count
- ✅ Case-insensitive matching

#### Test 4.3: Search Functionality - Sender
**Steps:**
1. Load conversations
2. Type sender name in search box

**Expected Results:**
- ✅ Filters by sender name
- ✅ Shows all conversations from that sender
- ✅ Real-time filtering

#### Test 4.4: Clear Search
**Steps:**
1. Perform a search
2. Clear the search box

**Expected Results:**
- ✅ All conversations reappear
- ✅ Status updates to show all conversations

#### Test 4.5: Appearance Mode Toggle
**Steps:**
1. Click appearance dropdown
2. Select "Light"
3. Select "Dark"
4. Select "System"

**Expected Results:**
- ✅ UI switches to light theme
- ✅ UI switches to dark theme
- ✅ UI follows system preference
- ✅ All text remains readable

---

### 5. Visual Tests

#### Test 5.1: Unread Email Indicators
**Steps:**
1. Ensure some inbox emails are unread
2. Refresh conversations

**Expected Results:**
- ✅ Unread emails show 🔵 indicator
- ✅ Unread conversations have blue border
- ✅ Sender text is light blue for unread
- ✅ Subject is bold for conversations with unread

#### Test 5.2: Conversation Grouping Visual
**Steps:**
1. View conversation with 5+ emails

**Expected Results:**
- ✅ Shows 💬 icon
- ✅ Shows message count "(X messages)"
- ✅ Last 3 emails visible
- ✅ "... and X more" indicator shown
- ✅ Card visually distinct

#### Test 5.3: Timestamp Formatting
**Steps:**
1. Load conversations with various dates

**Expected Results:**
- ✅ Format: "MMM DD, YYYY HH:MM AM/PM"
- ✅ Example: "Oct 16, 2025 03:45 PM"
- ✅ Consistent formatting throughout

#### Test 5.4: Responsive Layout
**Steps:**
1. Resize window to various sizes
2. Test minimum and maximum sizes

**Expected Results:**
- ✅ Sidebar maintains fixed width
- ✅ Main area scales properly
- ✅ Conversation cards expand to fill width
- ✅ No text cutoff or overlap

---

### 6. Error Handling Tests

#### Test 6.1: Connection Error Dialog
**Steps:**
1. Trigger connection failure

**Expected Results:**
- ✅ Error dialog appears
- ✅ Title: "Connection Error"
- ✅ Clear error message
- ✅ Helpful troubleshooting steps
- ✅ OK button dismisses dialog

#### Test 6.2: Refresh Without Connection
**Steps:**
1. If connection fails, try to refresh

**Expected Results:**
- ✅ Shows "Not Connected" error
- ✅ Instructs to restart application

---

### 7. Performance Tests

#### Test 7.1: Threading - UI Responsiveness
**Steps:**
1. Click refresh on large inbox
2. Try to interact with UI during load

**Expected Results:**
- ✅ UI remains responsive
- ✅ Can change appearance mode during load
- ✅ Status updates show loading state
- ✅ No freezing or "Not Responding"

#### Test 7.2: Memory Usage
**Steps:**
1. Load large inbox (500+ emails)
2. Monitor Task Manager

**Expected Results:**
- ✅ Memory usage reasonable (<500MB)
- ✅ No memory leaks on multiple refreshes

---

### 8. Edge Cases

#### Test 8.1: Emails with No Subject
**Steps:**
1. Send email with blank subject
2. Refresh inbox

**Expected Results:**
- ✅ Shows as "(No Subject)"
- ✅ No errors or crashes

#### Test 8.2: Emails with Very Long Subjects
**Steps:**
1. Email with 200+ character subject

**Expected Results:**
- ✅ Subject displays without breaking layout
- ✅ Text wraps or truncates appropriately

#### Test 8.3: Emails with Special Characters
**Steps:**
1. Subjects with emojis, Unicode, special chars

**Expected Results:**
- ✅ Displays correctly
- ✅ No encoding errors
- ✅ Search works with special characters

#### Test 8.4: Empty Search Results
**Steps:**
1. Search for non-existent term

**Expected Results:**
- ✅ Shows "No conversations found"
- ✅ Stats show 0 conversations
- ✅ Can clear search to restore view

---

## Architecture Verification

### Code Structure Test
**Verify:**
- ✅ MVC pattern properly implemented
- ✅ Model (`OutlookModel`) handles all Outlook logic
- ✅ View (`MainWindow`) handles all UI
- ✅ Controller (`OutlookInboxApp`) coordinates them
- ✅ No tight coupling between layers

### Threading Test
**Verify:**
- ✅ Outlook operations run in background threads
- ✅ UI updates happen on main thread
- ✅ No race conditions
- ✅ Proper use of `view.after()` for thread safety

---

## Test Summary Template

```
Test Date: __________
Tester: __________
Python Version: __________
Windows Version: __________
Outlook Version: __________

Total Tests: 30+
Passed: ___
Failed: ___
Skipped: ___

Pass Rate: ___%

Critical Issues Found:
1.
2.

Notes:
```

---

## Automated Testing (Future Enhancement)

Consider adding:
- Unit tests for `OutlookModel` methods
- Mock COM interface for testing without Outlook
- UI automation tests with pytest + pytest-qt
- Integration tests for full workflow

## Known Limitations

1. Requires Windows OS (COM interface limitation)
2. Requires Outlook installed and configured
3. Cannot run in WSL without Windows Outlook access
