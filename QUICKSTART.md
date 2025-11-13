# 🚀 Quick Start - Salary Ranges Calculator

## ⚡ 3-Minute Setup

### Step 1: Install clasp (if needed)

```bash
npm install -g @google/clasp
clasp login
```

Enable API: https://script.google.com/home/usersettings

### Step 2: Create Project

```bash
cd "/Users/mujahidreza/Cursor/Cloud Agent Space/salary-ranges"
clasp create --type sheets --title "Salary Ranges Calculator"
```

✅ This creates a Google Sheet and updates `.clasp.json` automatically!

### Step 3: Push Code

```bash
npm run push
```

Pushes the consolidated script (`SalaryRangesCalculator.gs`) + HTML UI.

### Step 4: Configure HiBob

In your Google Sheet:
1. **Extensions > Apps Script > ⚙️ Project Settings**
2. **Script Properties** → Add:
   - `BOB_ID` = your_id
   - `BOB_KEY` = your_key

### Step 5: Load Aon Data

**Aon Files**: [Google Drive](https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-)

In your sheet:
1. **💰 Salary Ranges Calculator > ⚙️ Setup > 🌍 Create Aon Region Tabs**
2. Download Aon files from Drive folder
3. Paste data into:
   - `Aon US Premium - 2025`
   - `Aon UK London - 2025`
   - `Aon India - 2025`

### Step 6: Initialize

```
1. 💰 Menu > ⚙️ Setup > 🗺️ Create Mapping Tabs
2. 💰 Menu > 🏗️ Build > 🌱 Seed Exec Mappings
3. 💰 Menu > ⚙️ Setup > 📊 Build Calculator UI
```

## ✅ Done! Now Use It

### Import Employee Data
```
💰 Menu > 📥 Import Data > Import All Bob Data
```

### Generate Ranges
```
💰 Menu > 🏗️ Build > Rebuild Full List Tabs
```

### Calculate Ranges
Use the **Salary Ranges** sheet or formulas:

```javascript
=SALARY_RANGE_MIN("X0", "US", "EN.SODE", "L5 IC")
=AON_P50("UK", "EN.SODE", "L6 IC")
=INTERNAL_STATS("India", "EN.SODE", "L5 IC")
```

## 📊 Categories

- **X0**: P62.5 / P75 / P90 - *Top of market*
- **X1**: P50 / P62.5 / P75 - *Mid-market*
- **Y1**: P40 / P50 / P62.5 - *Entry-level*

## 🔄 Regular Workflow

```
1. Import Data → Import All Bob Data (monthly/quarterly)
2. Build → Rebuild Full List Tabs (after imports)
3. Use calculator or formulas (as needed)
```

## 💻 Quick Commands

```bash
npm run push          # Push changes to Apps Script
npm run open          # Open project in browser
npm run watch         # Auto-push on file save
npm run deploy        # Push + commit + git push
```

## ❓ Common Issues

**Script ID not set?**
```bash
clasp create --type sheets --title "Salary Ranges Calculator"
```

**Not logged in?**
```bash
clasp logout && clasp login
```

**Data not showing?**
```
💰 Menu > 🏗️ Build > Clear All Caches
💰 Menu > 🏗️ Build > Rebuild Full List Tabs
```

**API not enabled?**  
https://script.google.com/home/usersettings

## 📚 More Help

- Full docs: [README.md](README.md)
- Detailed setup: [SETUP.md](SETUP.md)
- Aon data: https://drive.google.com/drive/folders/1bTogiTF18CPLHLZwJbDDrZg0H3SZczs-

---

**That's it!** You're ready to calculate salary ranges. 🎉
