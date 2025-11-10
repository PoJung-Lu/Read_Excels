# Building the Executable

This guide explains how to create a standalone `.exe` file for the Excel Data Processor.

## Quick Start

1. **Open a terminal** in the `Read_Excels` directory
2. **Run the build script:**
   ```bash
   python build_executable.py
   ```
3. Wait for the build to complete (2-5 minutes)
4. Find your executable in `dist/ExcelDataProcessor.exe`

## Detailed Instructions

### Prerequisites

- Python 3.11 or 3.12 installed
- All project dependencies installed (`pandas`, `openpyxl`, etc.)

### Step 1: Build the Executable

```bash
cd /path/to/Read_Excels
python build_executable.py
```

The script will:
- Install PyInstaller and PyYAML if needed
- Create a PyInstaller spec file
- Build the executable
- Place it in `dist/ExcelDataProcessor.exe`

### Step 2: Set Up Project Folder

Create a folder structure like this:

```
Project_A/
├── ExcelDataProcessor.exe    # Your executable
├── config.yaml                # Configuration file
└── Data/                      # Your data folder
    ├── 消防機關救災能量/
    │   ├── city_1/
    │   ├── city_2/
    │   └── ...
    └── 高科技產業/
        ├── company_1/
        ├── company_2/
        └── ...
```

**Important:** The executable, `config.yaml`, and `Data` folder **must be in the same directory**.

### Step 3: Copy Files

```bash
# Copy the executable
cp dist/ExcelDataProcessor.exe /path/to/Project_A/

# Copy the config file
cp config.yaml /path/to/Project_A/

# Copy your data
cp -r /path/to/your/data/folders /path/to/Project_A/Data/
```

### Step 4: Configure Settings

Edit `Project_A/config.yaml` to adjust settings:

```yaml
# Configuration file for Excel Data Processing Application

firefighter:
  base: "./Data/消防機關救災能量"  # Relative path to data
  output: "/Output"                # Output subdirectory
  enabled: true                     # Enable/disable this analysis

industry:
  base: "./Data/高科技產業"        # Relative path to data
  output: "/Output"                # Output subdirectory
  enabled: true                     # Enable/disable this analysis

general:
  auto_run: false                   # Auto-run on startup
  show_console: true                # Show console output
```

### Step 5: Run the Application

Simply double-click `ExcelDataProcessor.exe`!

## Using the GUI Application

### Main Features

1. **Firefighter Training Survey Analysis**
   - Configure base directory (where your firefighter data is)
   - Set output directory (relative to base)
   - Enable/disable this analysis
   - Browse for directories with GUI

2. **High-Tech Industry Rescue Equipment Analysis**
   - Configure base directory (where your industry data is)
   - Set output directory (relative to base)
   - Enable/disable this analysis
   - Browse for directories with GUI

3. **Processing Options**
   - Run analyses individually or together
   - View real-time processing logs
   - Save configuration for future use

### GUI Controls

- **Browse Buttons**: Select directories visually
- **Enable Checkboxes**: Turn analyses on/off
- **Run Buttons**: Execute specific or all analyses
- **Save Configuration**: Save current settings to `config.yaml`
- **Reset to Defaults**: Restore original settings

### Processing Logs

The application shows detailed logs in the bottom panel:
- Processing steps
- File operations
- Warnings and errors
- Completion status

## Customization

### Change Executable Name

Edit `build_executable.py`, find:
```python
name='ExcelDataProcessor',
```
Change to your preferred name.

### Add an Icon

1. Create or obtain a `.ico` file
2. Edit `build_executable.py`, find:
   ```python
   icon=None,
   ```
3. Change to:
   ```python
   icon='path/to/your/icon.ico',
   ```
4. Rebuild

### Enable Console Window

If you want to see a console window for debugging:

Edit `build_executable.py`, find:
```python
console=False,
```
Change to:
```python
console=True,
```

## Troubleshooting

### Build Fails

**Issue:** PyInstaller errors during build

**Solutions:**
- Ensure all dependencies are installed: `pip install -r requirements.txt`
- Try: `pip install pyinstaller --upgrade`
- Clean build: `rm -rf build dist *.spec` then rebuild

### Executable Won't Run

**Issue:** Double-clicking does nothing or shows error

**Solutions:**
- Check that `config.yaml` is in the same directory as the `.exe`
- Check that `Data` folder exists and has correct structure
- Try running from command line to see error messages:
  ```bash
  cd Project_A
  ./ExcelDataProcessor.exe
  ```

### "Directory Not Found" Error

**Issue:** Application can't find data directories

**Solutions:**
- Verify directory structure matches `config.yaml` paths
- Use relative paths starting with `./` in config
- Use the Browse buttons in the GUI to select correct paths
- Check that paths use forward slashes `/` not backslashes `\`

### Missing Dependencies

**Issue:** ImportError when running executable

**Solutions:**
- Add missing modules to `hiddenimports` in `build_executable.py`:
  ```python
  hiddenimports=[
      'openpyxl',
      'openpyxl.cell._writer',
      'yaml',
      'your_missing_module',  # Add here
  ],
  ```
- Rebuild the executable

## Manual Build (Advanced)

If you prefer to build manually:

```bash
# Install PyInstaller
pip install pyinstaller pyyaml

# Create spec file (or use build_executable.py to generate it)
pyi-makespec --onefile --windowed gui_launcher.py

# Edit the spec file as needed

# Build
pyinstaller --clean ExcelDataProcessor.spec
```

## Distribution

To share your application:

1. **Package the entire folder:**
   ```
   Project_A/
   ├── ExcelDataProcessor.exe
   ├── config.yaml
   └── Data/  (optional, users can add their own)
   ```

2. **Compress to ZIP:**
   ```bash
   zip -r ExcelDataProcessor.zip Project_A/
   ```

3. **Include instructions** on:
   - Where to place data files
   - How to edit `config.yaml`
   - System requirements (Windows 10/11)

## Notes

- The executable is **Windows-only** (built on Windows)
- File size: ~50-100MB (includes Python + all dependencies)
- First launch may be slow (2-3 seconds)
- No Python installation needed on target machine
- Data files are **not** included in the executable (by design)

## Support

For issues or questions:
1. Check this guide first
2. Review error messages in the log panel
3. Check GitHub issues: https://github.com/PoJung-Lu/Read_Excels/issues
