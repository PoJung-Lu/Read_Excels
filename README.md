# Excel Data Processing System

A Python-based data processing system for analyzing industrial safety and emergency response data for high-tech industries and firefighting organizations in Taiwan.

## Features

- **Chemical Storage Analysis**: Processes chemical inventory data from industrial facilities
- **Rescue Equipment Tracking**: Aggregates emergency response equipment inventories
- **Firefighter Training Analysis**: Analyzes personnel structure and training certifications across divisions
- **Regional Aggregation**: Organizes data by Taiwan industrial park regions

## Installation

### Prerequisites

- Python 3.10+
- pip package manager

### Dependencies

Install required packages:

```bash
pip install -r requirements.txt
```

Or manually:

```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage

```python
from pathlib import Path
import utils.read_data as read_data
from utils.output_excel import output_as

# Configure parameters
parameters = {
    "path_data": "../Data/科技廠救災能量",
    "pattern": "default",
    "path_output": "/Output",
    "file_name": "Aggregated_data.xlsx",
}

# Process data
data_reader = read_data.read_data(parameters)
combined_data = data_reader.read_excel_files()
output_as(combined_data, data_reader.parameters)
```

### Processing Workflows

#### High-Tech Industry Chemical Analysis

```python
from Read_excels_as_one import high_tech_industry_chems_main

high_tech_industry_chems_main(
    base="../Data/科技廠救災能量",
    out_rel="/Output"
)
```

#### Rescue Equipment Analysis

```python
from Read_excels_as_one import high_tech_industry_rescue_equipment_main

high_tech_industry_rescue_equipment_main(
    base="../Data/科技廠救災能量",
    out_rel="/Output/Rescue_equipment"
)
```

#### Firefighter Training Survey

```python
from Read_excels_as_one import firefighter_training_survey_main

firefighter_training_survey_main(
    base="../Data/消防機關救災能量",
    out_rel="../Output"
)
```

## Data Structure

The system expects Excel files organized in a hierarchical folder structure:

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

## Processing Patterns

- `default`: Standard Excel file processing
- `top_ten_operating_chemicals`: Chemical storage data extraction
- `industry_rescue_equipment`: Equipment inventory processing
- `firefighter_rescue_survey`: Training and personnel analysis
- `sort_by_location`: Regional aggregation by Taiwan industrial parks

## Output

The system generates consolidated Excel files with:

- Aggregated data organized by region/category
- Summary statistics and groupby analyses
- Multi-sheet workbooks for different data dimensions

## Development

### Interactive Development

```bash
jupyter notebook Test.ipynb
```

### Running Main Script

```bash
python Read_excels_as_one.py
```

## Project Structure

```
Read_Excels/
├── Read_excels_as_one.py          # Main entry point and workflows
├── utils/
│   ├── read_data.py               # Core data reading class
│   ├── patterns.py                # Pattern-based extraction logic
│   ├── output_excel.py            # Excel output handling
│   ├── data_cleaners.py           # Data cleaning functions
│   ├── industry_analysis.py       # Industrial data analysis
│   └── firefighter_analysis.py    # Firefighter survey analysis
├── Test.ipynb                     # Interactive testing notebook
├── CLAUDE.md                      # AI assistant documentation
└── README.md                      # This file
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
