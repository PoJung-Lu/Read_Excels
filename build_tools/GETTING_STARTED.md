# Getting Started - Build Your .exe in One Step

## Quick Start (TL;DR)

**Windows users:** Just double-click `BUILD_ALL.bat` and wait 2-5 minutes!

## What You Need

- Windows 10 or 11
- Python 3.11+ installed
- Internet connection (for downloading packages)

## Step-by-Step Instructions

### 1. Open Windows File Explorer

Navigate to:
```
C:\Users\user\Desktop\Python\Work\Read_Excels
```

### 2. Run the Build Script

**Double-click:** `BUILD_ALL.bat`

The script will automatically:
- ✅ Check Python installation
- ✅ Install all dependencies (pandas, openpyxl, pyyaml, pyinstaller)
- ✅ Build the executable
- ✅ Verify success

### 3. Wait for Completion

The build process takes 2-5 minutes. You'll see:
```
============================================================
                    BUILD SUCCESSFUL!
============================================================

Executable: dist\ExcelDataProcessor.exe
```

### 4. Deploy Your Application

After successful build:

```
# Create your deployment folder
mkdir ExcelProcessor

# Copy files
copy dist\ExcelDataProcessor.exe ExcelProcessor\
copy config.yaml ExcelProcessor\

# Add your data
mkdir ExcelProcessor\Data
# Then add your data folders inside Data\
```

### 5. Final Structure

Your deployment folder should look like:
```
ExcelProcessor\
├── ExcelDataProcessor.exe    ← Double-click to run
├── config.yaml               ← Edit to change paths
└── Data\
    ├── 消防機關救災能量\
    └── 高科技產業\
```

## Troubleshooting

### Issue: "Python is not installed"

**Solution:** Install Python from https://www.python.org/
- Download Python 3.11 or 3.12
- **Important:** Check "Add Python to PATH" during installation

### Issue: "Package installation failed"

**Solution:**
1. Check internet connection
2. Try running as Administrator:
   - Right-click `BUILD_ALL.bat`
   - Select "Run as Administrator"

### Issue: "Module not found: pandas"

**Solution:** The build script should install this automatically. If not:
```cmd
pip install pandas openpyxl pyyaml pyinstaller
```
Then run `BUILD_ALL.bat` again.

### Issue: Build succeeds but no .exe

**Solution:**
1. Check the `dist` folder manually
2. Try building with console enabled for debugging:
   - Edit `BUILD_ALL.bat` line 95
   - Change `--windowed` to `--console`
   - Run again

## Files in This Directory

### Essential Files (Don't Delete)
- `BUILD_ALL.bat` - **One-click build script** (double-click this!)
- `gui_launcher.py` - GUI application source code
- `Read_excels_as_one.py` - Main processing logic
- `config.yaml` - Configuration template
- `requirements.txt` - Python dependencies
- `utils/` - Utility modules directory

### Documentation (For Reference)
- `GETTING_STARTED.md` - This file (quick start guide)
- `USER_GUIDE.md` - End-user documentation
- `BUILD_INSTRUCTIONS.md` - Detailed build info
- `EXECUTABLE_README.md` - Technical reference
- `README.md` - Project overview

### Generated During Build (Can Be Deleted)
- `build/` - Temporary build files
- `dist/` - Contains the final .exe
- `ExcelDataProcessor.spec` - PyInstaller config
- `__pycache__/` - Python cache

## Next Steps After Building

1. **Test the executable:**
   - Copy to a test folder with sample data
   - Double-click to run
   - Verify it works correctly

2. **Share with others:**
   - Zip the deployment folder
   - Include `USER_GUIDE.md` for instructions
   - Share the zip file

3. **Update the application:**
   - Modify Python code as needed
   - Run `BUILD_ALL.bat` again
   - Replace the old .exe with the new one

## Support

If you encounter issues:
1. Read the error message carefully
2. Check the troubleshooting section above
3. Review `BUILD_INSTRUCTIONS.md` for detailed help
4. Check GitHub issues: https://github.com/PoJung-Lu/Read_Excels/issues

---

**Remember:** You only need to run `BUILD_ALL.bat` once. After that, you'll have a standalone `.exe` that works on any Windows computer without Python installed!
