# HARA Automation Tool

A professional automation tool for **Hazard Analysis and Risk Assessment (HARA)** with automated **ASIL (Automotive Safety Integrity Level)** determination according to ISO 26262 standards.

[![Preview](https://github.com/jayanthmani8045/hara-tool/blob/main/preview/Success_Execution_Image.png)](https://github.com/jayanthmani8045/hara-tool/blob/main/preview/Success_Execution_Image.png)

## UML

```mermaid
classDiagram
    class HARAMainWindow {
        -user_file: str
        -result_df: DataFrame
        -advanced_settings: Dict
        -processor: ExcelProcessor
        -file_label: QLabel
        -os_combo: QComboBox
        -ra_combo: QComboBox
        -result_table: QTableWidget
        -log_text: QTextEdit
        -progress_bar: QProgressBar
        +init_ui()
        +browse_file()
        +load_sheets()
        +process_data()
        +display_results(df)
        +save_results()
        +export_results()
        +log(message)
        +update_progress(value)
        +on_process_complete(df)
        +on_process_error(msg)
        +show_advanced_settings()
    }

    class AdvancedMatchingDialog {
        -case_sensitive: QCheckBox
        -strip_whitespace: QCheckBox
        -os_weight: QSpinBox
        -hazard_weight: QSpinBox
        +init_ui()
    }

    class ExcelProcessor {
        <<QThread>>
        -user_file: str
        -os_sheet: str
        -ra_sheet: str
        -matcher: FuzzyMatcher
        -os_weight: float
        -hazard_weight: float
        +progress: Signal
        +log: Signal
        +finished: Signal
        +error: Signal
        +run()
        +process_operating_scenarios(df)
        +match_risk_assessment(df, ra_df)
        +determine_asil(df)
        +find_column_case_insensitive(df, name)
        +diagnose_s_c_values(df, s_col, c_col)
    }

    class FuzzyMatcher {
        -enabled: bool
        -threshold: int
        -algorithm: str
        -case_sensitive: bool
        -strip_whitespace: bool
        +prepare_string(text) str
        +calculate_score(str1, str2) int
        +find_best_match(target, candidates, column, secondary_column, weights) Tuple
        +match_dataframes(source_df, target_df, match_columns, os_weight, hazard_weight) DataFrame
    }

    class constants {
        <<module>>
        +ASIL_DETERMINATION_TABLE: Dict
        +ASIL_COLORS: Dict
        +DEFAULT_SETTINGS: Dict
        +COLUMN_MAPPINGS: Dict
        +PARAMETER_RANGES: Dict
    }

    class styles {
        <<module>>
        +DARK_THEME: str
        +TITLE_STYLE: str
        +SUBTITLE_STYLE: str
        +FILE_LABEL_STYLE_DEFAULT: str
        +FILE_LABEL_STYLE_SELECTED: str
        +PROCESS_BUTTON_STYLE: str
        +FUZZY_CHECKBOX_STYLE: str
        +THRESHOLD_LABEL_STYLE: str
        +STATUS_LABEL_STYLE: str
        +INFO_LABEL_STYLE: str
    }

    class main {
        <<module>>
        +main() void
    }

    class QMainWindow {
        <<PySide6>>
    }

    class QDialog {
        <<PySide6>>
    }

    class QThread {
        <<PySide6>>
    }

    class pandas {
        <<external>>
        +DataFrame
        +ExcelFile
        +ExcelWriter
    }

    class fuzzywuzzy {
        <<external>>
        +fuzz.ratio()
        +fuzz.partial_ratio()
        +fuzz.token_sort_ratio()
        +fuzz.token_set_ratio()
    }

    %% Inheritance relationships
    HARAMainWindow --|> QMainWindow : inherits
    AdvancedMatchingDialog --|> QDialog : inherits
    ExcelProcessor --|> QThread : inherits

    %% Composition/Association relationships
    HARAMainWindow "1" *-- "1" ExcelProcessor : creates/uses
    HARAMainWindow "1" ..> "0..*" AdvancedMatchingDialog : creates
    ExcelProcessor "1" *-- "1" FuzzyMatcher : contains
    
    %% Dependency relationships
    HARAMainWindow ..> styles : imports
    HARAMainWindow ..> constants : imports
    ExcelProcessor ..> constants : imports
    ExcelProcessor ..> pandas : uses
    FuzzyMatcher ..> fuzzywuzzy : uses
    FuzzyMatcher ..> pandas : uses
    main ..> HARAMainWindow : instantiates

    %% Notes
    note for HARAMainWindow "Main GUI window that manages\nthe entire application flow\nand user interactions"
    
    note for ExcelProcessor "Background thread for\nprocessing Excel files\nwithout blocking UI"
    
    note for FuzzyMatcher "Handles fuzzy string matching\nwith configurable algorithms:\n- Ratio\n- Partial Ratio\n- Token Sort\n- Token Set"
```

## Features

- **Automated ASIL Determination**: Built-in ISO 26262 ASIL determination table
- **Fuzzy Matching**: Advanced string matching algorithms for scenario matching
- **Excel Integration**: Direct processing of Excel HARA sheets
- **Professional UI**: Modern dark theme with intuitive controls
- **Real-time Processing**: Background processing with progress tracking
- **Flexible Matching**: Multiple fuzzy matching algorithms (Ratio, Partial, Token Sort, Token Set)
- **Export Options**: Save results to original file or export to new file

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Install from source

1. Clone the repository:
```bash
git clone https://github.com/jayanthmani8045/hara-tool.git
cd hara-tool
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Run the application:
```bash
python -m hara_tool.main
```

### Install as package

```bash
pip install .
```

Then run:
```bash
hara-tool
```

## Usage

### 1. File Selection
- Click **Browse** to select your Excel file containing HARA data
- Select the **Operating Scenario** sheet from the dropdown
- Select the **Risk Assessment** sheet from the dropdown

### 2. Configure Matching Settings
- **Enable Fuzzy Matching**: Toggle fuzzy matching on/off
- **Threshold**: Set minimum match score (50-100%)
- **Algorithm**: Choose matching algorithm:
  - **Ratio**: Simple string similarity
  - **Partial Ratio**: Matches substrings
  - **Token Sort Ratio**: Ignores word order
  - **Token Set Ratio**: Ignores duplicates and order
- **Advanced**: Configure case sensitivity and weights

### 3. Process HARA
- Click **Process HARA** to start processing
- Monitor progress in the progress bar
- View results in the Results tab
- Check processing details in the Log tab

### 4. Export Results
- **Save to Excel**: Add RESULT sheet to original file
- **Export New File**: Create a new Excel file with results

## Excel Sheet Format

### Operating Scenario Sheet
Required columns:
- `Operating Scenario`: Description of the operating scenario
- `E` (optional): Exposure value (1-4), defaults to 4
- `Hazard` columns (optional): Associated hazards

### Risk Assessment Sheet
Required columns:
- `Operating Scenario`: Must match scenarios from OS sheet
- `Hazard`: Hazard description
- `S`: Severity value (0-3)
- `C`: Controllability value (0-3)

Optional columns:
- `Hazardous Event`: Description of hazardous event
- `Details of Hazardous Event`: Additional details
- `People at Risk`: Affected persons
- `Δv`: Delta velocity
- `Severity Rationale`: Justification for S value
- `Controllability Rationale`: Justification for C value

## ASIL Determination Table

The tool uses the standard ISO 26262 ASIL determination table:

| S/E | E1 | E2 | E3 | E4 |
|-----|----|----|----|----|
| **S1** | | | | |
| C0 | QM | QM | QM | QM |
| C1 | QM | QM | QM | QM |
| C2 | QM | QM | QM | A |
| C3 | QM | QM | A | B |
| **S2** | | | | |
| C0 | QM | QM | QM | A |
| C1 | QM | QM | QM | A |
| C2 | QM | QM | A | B |
| C3 | QM | A | B | C |
| **S3** | | | | |
| C0 | QM | QM | A | B |
| C1 | QM | QM | A | B |
| C2 | A | B | B | C |
| C3 | B | C | C | D |

## Development

### Project Structure
```
hara_tool/
├── __init__.py          # Package initialization
├── main.py              # Entry point
├── gui.py               # Main window and dialogs
├── processor.py         # Excel processing logic
├── matcher.py           # Fuzzy matching logic
├── constants.py         # ASIL table and constants
└── styles.py            # UI styling
```

### Building Executable

To create a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=resources/icon.ico hara_tool/main.py
```

## Authors

- SafeLinkInnovations team and Jayanth Mani

## Acknowledgments

- ISO 26262 standard for ASIL determination methodology
- FuzzyWuzzy library for string matching algorithms
