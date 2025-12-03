# ETL File Processor - Auto Mode

## 🎯 How It Works

1. **Drop** your Excel file in the `upload/` folder
2. **Wait** a few seconds (automatic processing)
3. **Get** your output file (same name, .xlsx extension)
4. **Check** the archive folder if needed

That's it!

## 🚀 Setup (One Time)

```bash
pip install watchdog
```

## ▶️ Start the Processor

```bash
python watch.py
```

Keep this running. It monitors the `upload/` folder automatically.

## 📁 Folder Structure

```
C:\py\Autoapp\
├── watch.py              ← Run this!
├── upload/               ← Drop your Excel files here
├── archive/              ← Processed files go here
├── logs/                 ← Processing details (if needed)
└── output .xlsx files    ← Your processed results
```

## 📝 Example

1. Drop `myquote.xlsm` into the `upload/` folder
2. Wait 2-3 seconds
3. Find `myquote.xlsx` in the main folder
4. `myquote.xlsm` automatically moved to `archive/`

## 📊 Output Format

Your output file has one sheet with:
- **Cover Details** (extracted from your Excel)
  - Quotation #, QDR #, SPR #, etc.
  - Contact name, Company, Address, etc.
- **Data from each sheet** in this format:
  ```
  ModelName 1.0 16627.25,*,*,*,16170.00
  ```

## ⚙️ Requirements

Your Excel file needs:
- ✅ Sheets with columns: `Model Number`, `Qty` (or `Quantity`), `Net Price`
- ✅ Sheets must be visible (not hidden)
- ✅ Rows with Qty > 0

## ❓ Troubleshooting

### Nothing happened?
- Check the `logs/` folder for error messages
- Make sure file is in `upload/` folder
- Wait a few seconds and check if output appeared

### Output file missing?
- Check if it's still in `upload/` (might be processing)
- Check the `logs/` folder for errors
- Verify Excel file has required columns

### Need details?
- Check `logs/etl_*.log` file - it shows everything that happened

## 🎯 That's All!

Just:
1. `python watch.py` (run once, keep it running)
2. Drop files in `upload/` folder
3. Get results in main folder

No terminal commands needed after step 1! 🎉
