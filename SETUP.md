# Salary Ranges Calculator - Quick Setup Guide

## Step-by-Step Setup

### 1. Install clasp (if not already installed)

```bash
npm install -g @google/clasp
```

### 2. Login to Google

```bash
clasp login
```

This will open a browser window. Authorize clasp to access your Google account.

### 3. Enable Apps Script API

1. Go to https://script.google.com/home/usersettings
2. Turn on "Google Apps Script API"

### 4. Create Apps Script Project

You have two options:

#### Option A: Create New Sheet + Script (Recommended)

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space/salary-ranges"
clasp create --type sheets --title "Salary Ranges Calculator"
```

This will:
- Create a new Google Sheet
- Create an Apps Script project attached to it
- Automatically update `.clasp.json` with the Script ID
- Open the sheet URL in your terminal

#### Option B: Link to Existing Google Sheet

If you already have a Google Sheet with salary data:

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Copy the **Script ID** from the URL bar:
   ```
   https://script.google.com/home/projects/YOUR_SCRIPT_ID_HERE/edit
   ```
4. Update `.clasp.json`:
   ```bash
   # Edit .clasp.json and replace YOUR_SCRIPT_ID_HERE with your actual ID
   nano .clasp.json
   ```

### 5. Push Code to Apps Script

```bash
npm run push
```

Or:
```bash
./push_to_apps_script.sh
```

### 6. Configure HiBob API Credentials

1. Open your Google Sheet
2. Go to **Extensions > Apps Script**
3. Click the **⚙️ Project Settings** icon (left sidebar)
4. Scroll to **Script Properties**
5. Add properties:
   - Click **Add script property**
   - Property: `BOB_ID`, Value: `your_bob_api_id`
   - Click **Add script property**
   - Property: `BOB_KEY`, Value: `your_bob_api_key`

### 7. Set Up the Spreadsheet

In your Google Sheet:

1. Reload the sheet (press F5 or refresh)
2. You should see a new menu: **Salary Ranges**
3. Go to **Salary Ranges > Setup > Generate Help sheet**
4. Follow the Help sheet instructions

## Quick Start After Setup

### Import Employee Data
```
Salary Ranges > Imports > Import Bob Base Data
```

### Build Salary Ranges
```
Salary Ranges > Build > Rebuild Full List tabs
```

### Use the Calculator
```
Salary Ranges > Setup > Build Calculator UI
```

Then use the "Salary Ranges" sheet to calculate ranges interactively.

## Useful Commands

```bash
# Push changes to Apps Script
npm run push

# Pull from Apps Script (download)
npm run pull

# Open project in browser
npm run open

# Watch for changes and auto-push
npm run watch

# Deploy to Apps Script + Git
npm run deploy

# View execution logs
npm run logs

# Check status
npm run status
```

## File Structure

```
salary-ranges/
├── RangeCalculator.gs          # Main calculation engine (70KB)
├── AppImports.gs               # HiBob API import functions
├── Helpers.gs                  # Utility functions
├── ExecMappingManager.html     # Web UI for mappings
├── appsscript.json             # Apps Script manifest
├── .clasp.json                 # clasp configuration (YOU NEED TO UPDATE THIS)
├── package.json                # npm scripts
├── deploy.sh                   # Deployment script
├── push_to_apps_script.sh      # Quick push script
└── README.md                   # Full documentation
```

## Next Steps

1. **Create Aon Data Tabs**: `Salary Ranges > Setup > Create Aon placeholder tabs`
2. **Paste Aon Market Data**: Into US, UK, India tabs
3. **Create Mapping Tabs**: `Salary Ranges > Setup > Create mapping placeholder tabs`
4. **Import Bob Data**: `Salary Ranges > Imports > Import Bob Base Data`
5. **Build Full List**: `Salary Ranges > Build > Rebuild Full List tabs`
6. **Use Calculator**: `Salary Ranges > Setup > Build Calculator UI`

## Troubleshooting

### Script ID Not Set

If you see:
```
❌ Please update .clasp.json with your Google Apps Script ID
```

Run:
```bash
clasp create --type sheets --title "Salary Ranges Calculator"
```

Or manually update `.clasp.json` with your Script ID.

### Not Logged In

If you see login errors:
```bash
clasp login --status  # Check status
clasp logout          # Logout
clasp login           # Login again
```

### Push Failed

```bash
# Check clasp status
clasp login --status

# Check .clasp.json
cat .clasp.json

# Try manual push
clasp push
```

### API Not Enabled

Go to: https://script.google.com/home/usersettings
Enable: "Google Apps Script API"

## Getting Your Script ID

### If you created the project with clasp:
- The Script ID is automatically in `.clasp.json`
- You can also run: `clasp open` to see the project

### If you have an existing sheet:
1. Open the sheet
2. Extensions > Apps Script
3. Copy ID from URL: `https://script.google.com/home/projects/SCRIPT_ID_HERE/edit`

### If you want to find all your projects:
1. Go to: https://script.google.com/home
2. Browse your projects
3. Open one and copy the Script ID from the URL

## Support

- **clasp Docs**: https://github.com/google/clasp
- **Apps Script Docs**: https://developers.google.com/apps-script
- **HiBob API Docs**: https://apidocs.hibob.com/

---

✅ **Setup Complete!** You're ready to calculate salary ranges.

