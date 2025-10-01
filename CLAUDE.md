# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based Excel data processing system designed for analyzing industrial safety and emergency response data, specifically for high-tech industries and firefighting organizations in Taiwan. The system processes Excel files from various sources, aggregates data, and generates consolidated reports organized by regions and categories.

## Architecture

### Core Components

- **`Read_excels_as_one.py`**: Main entry point containing the primary processing workflows and utility functions
- **`utils/read_data.py`**: Core data reading class that handles Excel file processing with configurable patterns
- **`utils/patterns.py`**: Pattern-based data extraction logic for different data formats and structures
- **`utils/output_excel.py`**: Output handling for writing processed data to Excel files

### Key Classes

- **`read_data`**: Main data processing class that handles Excel file reading with pattern-based extraction
  - Supports multiple patterns: `default`, `苗栗縣`, `top_ten_operating_chemicals`, `sort_by_location`, `industry_rescue_equipment`, `firefighter_rescue_survey`
  - Returns generators for memory-efficient processing of large datasets

### Data Processing Patterns

The system uses different processing patterns for various data types:

1. **`top_ten_operating_chemicals`**: Processes chemical storage data from industrial facilities
2. **`industry_rescue_equipment`**: Handles rescue equipment inventory data
3. **`firefighter_rescue_survey`**: Processes firefighter training and capability surveys
4. **`sort_by_location`**: Organizes data by Taiwan regions (北部園區, 中部園區, 南部園區, 其他)

### Main Processing Workflows

1. **`high_tech_industry_chems_main()`**: Processes chemical storage data for high-tech industries
2. **`high_tech_industry_rescue_equipment_main()`**: Handles rescue equipment data analysis
3. **`firefighter_training_survey_main()`**: Processes firefighter capability surveys

## Common Development Commands

### Running the Application

```bash
# Run main processing (modify the main() function or workflow functions as needed)
python Read_excels_as_one.py

# The main entry point allows switching between different workflows by uncommenting the desired function
```

### Testing and Development

```bash
# Use Jupyter notebook for interactive development and testing
jupyter notebook Test.ipynb

# The Test.ipynb notebook contains experimental code and testing scenarios
```

## Key Configuration

### Data Paths
- Default input path: `../Data/科技廠救災能量` or `../Data/消防機關救災能量`
- Default output path: `/Output` (relative to input path)
- Excluded files: `("Output", "Distribution_by_city")` and various output file names

### Processing Parameters
The `read_data` class accepts a parameters dictionary with keys:
- `path_data`: Input data directory
- `pattern`: Processing pattern to use
- `path_output`: Output directory
- `file_name`: Output filename
- `folder_path`: Specific folder to process
- `read_all_sheets`: Whether to read all Excel sheets (default: True)

## Data Processing Flow

1. **Folder Tree Processing**: `process_folder_tree()` walks through subdirectories and processes Excel files
2. **Location Sorting**: `sort_by_location()` aggregates data by Taiwan regions
3. **Analysis**: `analyze_grouped()` performs groupby operations and generates summary reports

## Utility Functions

- **`extract_first_number()`**: Extracts numeric values from mixed text/number cells
- **`concat_list_dict()`**: Combines lists of DataFrames into single DataFrames
- **`merge_sheets_by_group()`**: Merges multiple sheets using groupby operations
- **Data cleaning functions**: `clean_chems()`, `clean_equipment()` for data type standardization

## File Structure Expectations

The system expects Excel files in a hierarchical folder structure:
```
Data/
├── 科技廠救災能量/          # High-tech industry data
│   ├── [Company folders]/
│   └── *.xlsx files
├── 消防機關救災能量/        # Firefighter data
│   ├── [City folders]/
│   │   ├── [Division folders]/
│   │   └── *.xlsx files
└── Output/                  # Generated output files
```

## Dependencies

The project uses standard Python scientific computing libraries:
- pandas: Data manipulation and analysis
- numpy: Numerical operations
- openpyxl: Excel file writing
- pathlib: Path operations
- logging: Application logging