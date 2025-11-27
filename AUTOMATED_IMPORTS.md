# ⏰ Automated Daily Imports Guide

## 🎯 Overview

The Salary Ranges Calculator now supports **fully automated daily imports** from HiBob via time-based triggers. No manual intervention required!

---

## ✨ Features

### **1. Headless Import Function**
- `importBobDataHeadless()` - Runs without UI prompts
- Compatible with time-based triggers
- Logs all steps to execution log
- Email notifications on failure
- Tracks last import timestamp

### **2. One-Click Trigger Setup**
- `setupDailyImportTrigger()` - Creates/deletes daily trigger
- Easy toggle on/off from menu
- Shows existing trigger status
- No Apps Script coding required

### **3. Automatic Features**
- ✅ Imports Base Data, Bonus, Comp, Performance Ratings
- ✅ Seeds Title Mapping from legacy data
- ✅ Syncs Employees Mapped with smart suggestions
- ✅ Refines Title Mapping from approved mappings
- ✅ Updates Legacy Mappings (feedback loop)
- ✅ Email alerts on failures
- ✅ Timestamp tracking

---

## 🚀 Quick Start (3 Steps)

### **Step 1: Enable the Daily Trigger**
```
1. Open your Google Sheet
2. Menu → 💰 Salary Ranges Calculator → 🔧 Tools
3. Click "🔔 Setup Daily Import Trigger"
4. Click "Yes" to confirm
5. Done! ✅
```

### **Step 2: Verify It Worked**
```
You'll see a confirmation:
"✅ Trigger Created!
 Schedule: Every day at 6:00-7:00 AM
 Email: You'll receive notifications on failures
 Tracking: Check Base Data cell ZZ1 for last import time"
```

### **Step 3: Wait for First Import**
```
Next day at 6 AM:
- Import runs automatically
- Base Data ZZ1 shows: "Last Import: 12/1/2025 6:15 AM"
- Check Execution Log for details (Extensions → Apps Script → Executions)
```

---

## 📋 Menu Options

| Menu Item | Use Case | UI Prompts? | Trigger Compatible? |
|-----------|----------|-------------|---------------------|
| **📥 Import Bob Data** | Manual import when you're present | ✅ Yes | ❌ No |
| **⏰ Import Bob Data (Headless)** | Testing automated import | ❌ No | ✅ Yes |
| **🔔 Setup Daily Import Trigger** | Enable/disable automation | ✅ Yes (setup only) | N/A |

---

## 🔍 How It Works

### **Manual Import (with prompts):**
```javascript
User clicks: Import Bob Data
  ↓
Shows confirmation dialog
  ↓ (User clicks Yes)
Calls: importBobDataHeadless()
  ↓
Shows success dialog
```

### **Automated Import (no prompts):**
```javascript
Trigger fires at 6 AM
  ↓
Calls: importBobDataHeadless()
  ↓
Logs to execution log
  ↓
Sends email on failure
  ↓
Updates ZZ1 timestamp
```

---

## 📊 Last Import Timestamp

Check **Base Data sheet, cell ZZ1** for:
```
Last Import: 12/1/2025 6:15:32 AM
```

This updates every time the import runs (manual or automated).

---

## 📧 Email Notifications

### **Success:**
- No email sent (import is logged only)

### **Failure:**
```
Subject: ⚠️ Bob Data Import Failed
Body:
  Import failed at 12/1/2025 6:15 AM
  
  Error: HTTP 401 Unauthorized
  
  Stack: [error details]
```

**Email sent to:** Your Google account email (the one that set up the trigger)

---

## 🔧 Troubleshooting

### **Trigger Not Running?**

**Check 1: Verify trigger exists**
```
1. Extensions → Apps Script
2. Triggers (clock icon in left sidebar)
3. Should see: importBobDataHeadless | Time-driven | Day timer
```

**Check 2: Review execution log**
```
1. Extensions → Apps Script
2. Executions (in left sidebar)
3. Look for: importBobDataHeadless
4. Check status: ✅ Completed or ❌ Failed
```

**Check 3: API Credentials**
```
BOB_ID and BOB_KEY must be set in Script Properties
1. Extensions → Apps Script
2. Project Settings (gear icon)
3. Script Properties
4. Verify BOB_ID and BOB_KEY exist
```

---

### **Import Failed?**

**Common Errors:**

| Error | Cause | Fix |
|-------|-------|-----|
| `HTTP 401 Unauthorized` | Invalid API credentials | Update BOB_ID/BOB_KEY |
| `HTTP 429 Rate Limit` | Too many requests | Wait 1 hour, try again |
| `ReferenceError: SHEET_NAMES is not defined` | Code deployment issue | Redeploy with `clasp push` |
| `Base Data not found` | Fresh Build not run | Run Fresh Build first |

**View Full Error:**
```
1. Extensions → Apps Script → Executions
2. Click on failed execution
3. View full error message and stack trace
4. Copy and share with support if needed
```

---

## 🔄 Change Trigger Schedule

### **Option 1: Delete and Recreate**
```
1. Menu → Tools → Setup Daily Import Trigger
2. Click "Yes" to delete existing
3. Run again to create new trigger with different time
```

### **Option 2: Manual Edit in Apps Script**
```
1. Extensions → Apps Script
2. Triggers (clock icon)
3. Click ⋮ (three dots) → Edit
4. Change time: 6 AM → 3 AM (or any time)
5. Save
```

### **Custom Schedules:**
```javascript
// Every 12 hours
ScriptApp.newTrigger('importBobDataHeadless')
  .timeBased()
  .everyHours(12)
  .create();

// Every Monday at 9 AM
ScriptApp.newTrigger('importBobDataHeadless')
  .timeBased()
  .onWeekDay(ScriptApp.WeekDay.MONDAY)
  .atHour(9)
  .create();

// Every hour (not recommended - rate limits!)
ScriptApp.newTrigger('importBobDataHeadless')
  .timeBased()
  .everyHours(1)
  .create();
```

---

## 🎯 Best Practices

### **✅ DO:**
- Set trigger to run during off-hours (6 AM is good)
- Check execution log weekly for issues
- Test headless import manually before enabling trigger
- Keep BOB_ID and BOB_KEY credentials secure
- Monitor email for failure notifications

### **❌ DON'T:**
- Set trigger to run hourly (rate limits!)
- Share Script Properties with others
- Delete Base Data sheet (breaks tracking)
- Run multiple triggers for same function
- Ignore failure emails

---

## 📚 Related Functions

| Function | Purpose | Trigger Compatible? |
|----------|---------|---------------------|
| `importBobDataHeadless()` | Import without UI | ✅ Yes |
| `importBobData()` | Import with UI | ❌ No |
| `setupDailyImportTrigger()` | Create/delete trigger | ❌ No (setup only) |
| `buildMarketData()` | Build Full Lists | ❌ No (UI prompts) |
| `freshBuild()` | Create all sheets | ❌ No (UI prompts) |

**Note:** Only `importBobDataHeadless()` is trigger-compatible (no UI dependencies).

---

## 🔐 Security & Permissions

### **Required Authorizations:**
1. **Google Sheets API** - Read/write sheet data
2. **HiBob API** - Fetch employee data (via BOB_ID/BOB_KEY)
3. **Gmail API** - Send failure notification emails
4. **Script Service** - Create/manage triggers

### **First Run Authorization:**
```
When you first run setupDailyImportTrigger():
1. Google will show authorization prompt
2. Click "Advanced" → "Go to [Your Project]"
3. Click "Allow"
4. Trigger will be created
```

### **Revoking Access:**
```
1. Google Account → Security
2. Third-party apps with account access
3. Find "Salary Ranges Calculator"
4. Click "Remove Access"
```

---

## 📞 Support

### **Need Help?**
1. Check execution log for errors
2. Review this guide's troubleshooting section
3. Check Base Data ZZ1 for last import time
4. Contact support with:
   - Execution ID (from log)
   - Error message
   - Timestamp of failure

---

## 📝 Version History

| Version | Date | Changes |
|---------|------|---------|
| 4.0.0 | Nov 27, 2025 | Added time-based trigger support |
| 3.9.4 | Nov 27, 2025 | Added error messages for Title Mapping |
| 3.9.3 | Nov 27, 2025 | Fixed chicken-and-egg problem |

---

**Status**: ✅ Active  
**Last Updated**: November 27, 2025  
**Compatibility**: Google Apps Script, HiBob API

