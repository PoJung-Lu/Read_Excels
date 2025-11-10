# Excel Data Processor - User Guide

## Quick Start

1. **Set up your folder structure:**
   ```
   Project_A/
   ├── ExcelDataProcessor.exe
   ├── config.yaml
   └── Data/
       ├── 消防機關救災能量/
       └── 高科技產業/
   ```

2. **Double-click** `ExcelDataProcessor.exe` to launch

3. **Configure paths** in the GUI or edit `config.yaml`

4. **Click** "Run Both" to process all data

## Folder Structure Explained

### Required Structure
```
Project_A/                          # Your main folder (can be any name)
├── ExcelDataProcessor.exe          # The application executable
├── config.yaml                     # Configuration file (must be here!)
└── Data/                           # Data folder (must be named "Data")
    ├── 消防機關救災能量/            # Firefighter data
    │   ├── city_1/
    │   │   ├── division_1.xlsx
    │   │   └── division_2.xlsx
    │   └── city_2/
    │       └── division_1.xlsx
    └── 高科技產業/                  # Industry data
        ├── company_1.xlsx
        └── company_2.xlsx
```

### Output Structure
After processing, you'll see:
```
Project_A/
└── Data/
    ├── 消防機關救災能量/
    │   ├── Output/                 # Generated output files
    │   └── Distribution_by_city/   # Analysis results
    └── 高科技產業/
        └── Output/                 # Generated output files
```

## Using the Application

### Main Window

The application has four main sections:

#### 1. Working Directory
Shows where the executable is located. All paths in the config are relative to this location.

#### 2. Firefighter Training Survey
- **Enable checkbox**: Turn this analysis on/off
- **Base Directory**: Where your firefighter data is located
  - Default: `./Data/消防機關救災能量`
  - Click "Browse" to select a different folder
- **Output**: Where results will be saved (relative to base)
  - Default: `/Output`

#### 3. High-Tech Industry Rescue Equipment
- **Enable checkbox**: Turn this analysis on/off
- **Base Directory**: Where your industry data is located
  - Default: `./Data/高科技產業`
  - Click "Browse" to select a different folder
- **Output**: Where results will be saved (relative to base)
  - Default: `/Output`

#### 4. Processing Log
Shows real-time processing information:
- Files being read
- Processing steps
- Warnings or errors
- Completion status

### Control Buttons

- **Run Firefighter Analysis**: Process only firefighter data
- **Run Industry Analysis**: Process only industry data
- **Run Both**: Process both analyses sequentially
- **Save Configuration**: Save current settings to `config.yaml`
- **Reset to Defaults**: Restore original settings

## Configuration File

The `config.yaml` file stores your settings. You can edit it with any text editor:

```yaml
# Firefighter Training Survey Settings
firefighter:
  base: "./Data/消防機關救災能量"
  output: "/Output"
  enabled: true

# High-Tech Industry Rescue Equipment Settings
industry:
  base: "./Data/高科技產業"
  output: "/Output"
  enabled: true

# General Settings
general:
  auto_run: false
  show_console: true
```

### Path Rules

1. **Relative paths** start with `./`
   - Example: `./Data/消防機關救災能量`
   - Resolved relative to executable location

2. **Absolute paths** are used as-is
   - Example: `C:/Users/YourName/Documents/Data`

3. **Output paths** are relative to the base path
   - Example: If base is `./Data/消防機關救災能量` and output is `/Output`
   - Results go to `./Data/消防機關救災能量/Output/`

## Common Tasks

### Change Data Location

**Option 1: Using GUI**
1. Click "Browse" next to Base Directory
2. Navigate to your data folder
3. Click "Select Folder"
4. Click "Save Configuration"

**Option 2: Edit config.yaml**
1. Open `config.yaml` in a text editor
2. Change the `base` path:
   ```yaml
   firefighter:
     base: "./MyData/firefighter_data"
   ```
3. Save the file

### Disable an Analysis

**Option 1: Using GUI**
1. Uncheck the "Enable" checkbox
2. Click "Save Configuration"

**Option 2: Edit config.yaml**
```yaml
firefighter:
  enabled: false  # Changed from true to false
```

### Process Only One Dataset

**Method 1:** Disable the other analysis (see above)

**Method 2:** Click the specific button:
- "Run Firefighter Analysis" for firefighter only
- "Run Industry Analysis" for industry only

### View Results

After processing completes:
1. Navigate to your data folder
2. Look for the `Output` folder
3. Open the generated Excel files

Example:
- Firefighter results: `Data/消防機關救災能量/Output/`
- Industry results: `Data/高科技產業/Output/`

## Troubleshooting

### "Base directory does not exist"

**Problem:** The application can't find your data folder

**Solutions:**
1. Check that the path in config matches your folder structure
2. Make sure the `Data` folder is in the same directory as the `.exe`
3. Use the "Browse" button to select the correct folder
4. Check for typos in folder names

### "No files were found"

**Problem:** No Excel files in the specified directory

**Solutions:**
1. Verify your data folder contains `.xlsx`, `.xls`, or `.xlsm` files
2. Check that files are in the correct subdirectory structure
3. For firefighter data: files should be in `city_name/division.xlsx` format

### Application won't start

**Problem:** Double-clicking does nothing

**Solutions:**
1. Make sure `config.yaml` is in the same folder as the `.exe`
2. Check that you have permission to run the executable
3. Try right-click → "Run as Administrator"
4. Check Windows Defender or antivirus hasn't blocked it

### Processing fails with errors

**Problem:** Error messages in the log window

**Solutions:**
1. Read the error message carefully
2. Check that your Excel files are not corrupted
3. Close any Excel files that are currently open
4. Verify the data format matches expected structure
5. Check file permissions (not read-only)

### Log shows "Skipped sheets"

**Problem:** Warning about skipped sheets

**Explanation:** Some sheets don't have the required columns. This is normal if:
- Your data has summary sheets
- Some sheets use different formats
- Empty sheets exist

**Action:** Review the skipped sheets list. If important data is skipped, check the Excel file format.

## Tips

1. **Save Configuration**: Always click "Save Configuration" after making changes in the GUI

2. **Backup Data**: Keep a backup of your original data files before processing

3. **Test First**: Process a small subset of data first to verify settings

4. **Read Logs**: The processing log shows what's happening - watch for warnings

5. **Batch Processing**: Use "Run Both" to process everything at once

6. **Relative Paths**: Use relative paths (starting with `./`) for portability

7. **Close Excel**: Close Excel files before processing to avoid "file in use" errors

## Advanced Usage

### Multiple Configurations

Create multiple config files for different scenarios:

1. Copy `config.yaml` to `config_test.yaml`
2. Edit with different paths
3. Before running, rename:
   - `config.yaml` → `config_production.yaml`
   - `config_test.yaml` → `config.yaml`
4. Run the application

### Command Line (Advanced)

You can run the Python script directly for more control:

```bash
python gui_launcher.py
```

Or run analyses programmatically:
```bash
python Read_excels_as_one.py
```

### Custom Data Structures

If your data doesn't match the expected structure:
1. Review the source code in `Read_excels_as_one.py`
2. Adjust the pattern processing in `utils/patterns.py`
3. Rebuild the executable after changes

## Support

For issues, questions, or feature requests:
- Check this guide first
- Review error messages in the log panel
- Contact your system administrator
- Report bugs to the development team

## Version Information

- Application: Excel Data Processor
- Version: 1.0
- Python Version: 3.11/3.12
- Dependencies: pandas, openpyxl, pyyaml

---

**Note:** This application processes Excel files for firefighter training surveys and high-tech industry rescue equipment data. Ensure your data format matches the expected structure for best results.
