# DCR Format Converter

**Version:** 1.0  
**Programmer:** Sangwoo Kim  
**Acknowledgments:** Developed with significant assistance from Opus 4.5 and Gemini.  
**Corporate Identity:** Optimized for **Sumitomo Electric Industries (SEI)** brand standards.

---

## Project Overview

The **DCR Format Converter** is a professional-grade desktop application designed for the automated conversion, processing, and statistical analysis of electronic component measurement data. Built with **Python** and **PySide6**, the application features a modern **Material Design** interface utilizing the official Sumitomo Blue corporate identity.

It streamlines complex workflows by transforming raw `.NET` and `.xlsx` files into standardized Excel reports, performing high-speed statistical calculations, and generating comprehensive measurement result files with automated formatting and judge logic.

---

## Key Features

### 1. Tab 1: DCR Format Generation (`make DCR format`)
This module handles the core transformation of circuit netlist data and vendor specifications into a unified DCR (Direct Current Resistance) format.
- **Input Processing**: Parses `.NET` netlist files and multiple vendor/part-pin specification Excel files.
- **Automated Sheet Creation**:
  - `vendor`: Advanced formatting of vendor specs with automated formula injection.
  - `DE requirement`: Continuity data extraction from part-pin files with address mapping.
  - `input check pin interm` & `input check pin`: Merging of netlist `#4W` sections with circuit requirement data.
  - `Judge(check pin)`: Dynamic generation of pass/fail logic for pin configurations.
  - `DCR`: Final consolidated DCR report with structured net names and group pins.

### 2. Tab 2: Form Measurement Reporting (`make Form Measurement Result file`)
Generates standardized measurement result files by hard-copying templates and filling them with high-precision measurement data.
- **Template Synchronization**: Hard-copies styles, merged cells, and column properties from master templates.
- **Data Integration**:
  - **TDR Data**: Automated extraction of TDR (Time-Domain Reflectometry) data from multiple DK files in a selected directory.
  - **Dimension Analysis**: Extraction of circuit width and thickness data from dimension-specific Excel reports.
  - **Spec Mapping**: Retrieves LSL, Center, and USL values from master spec files.
- **Smart Formatting**: Applies conditional formatting (Green for OK, Red for NG) and adds AutoFilters to all data columns.

### 3. Tab 3: Statistical LSL/USL Calculation (`calculate LSL USL`)
Performs advanced statistical analysis on large-scale merged measurement files.
- **High-Performance Processing**: Analyzes thousands of data points across multiple pieces and sets.
- **Statistical Suite**: Calculates Min, Max, Average, Median, Stdev, IQR, and Outlier thresholds (1stQuat-4IQR, 3rdQuat+4IQR).
- **Hybrid Calculation Engine**: Performs complex `AverageIfs` and `StdevIfs` filtering in Python for speed and accuracy, while maintaining dynamic Excel formulas for LSL/USL results.
- **Final Output**: Produces `Calculate_3Sigma_LSLUSL_final.xlsx` containing raw rearranged data, statistical summaries, and judgment columns.

### 4. Enterprise-Grade Utilities
- **Auto Execute All**: A single-click feature that sequences all tabs (Tab 1 → Tab 2 → Tab 3) into a seamless end-to-end workflow.
- **Integrated Logging**: Real-time process monitoring with a dedicated high-capacity log window and automated `.dat` file output (`log_{Operator}_{Date}.dat`).
- **Automated Cover Pages**: Every generated Excel file automatically includes a professional Cover Page detailing the program version, operator name, creation timestamp, and a full audit trail of input/output file paths.
- **Configuration Management**: Persistent storage of file paths and operator info in `files.json` for rapid subsequent runs.

---

## Technical Architecture

### Directory Structure
```
0105python/
├── main.py                 # Application entry point
├── build_exe.py            # PyInstaller build orchestration
├── pyproject.toml          # UV project & dependency management
├── README.md               # Documentation
├── files.json              # Persistent user configuration (Auto-generated)
│
├── ui/                     # User Interface Layer
│   ├── __init__.py
│   ├── main_window.py      # Material Design UI (Sumitomo Blue theme)
│   └── logo.png            # Corporate logo asset
│
├── logic/                  # Business Logic Layer
│   ├── __init__.py
│   ├── config_manager.py   # JSON configuration handler
│   ├── file_reader.py      # Multi-encoding file I/O utilities
│   ├── cover_page.py       # Automated Excel report branding
│   ├── makevendor.py       # Vendor data processing
│   ├── make_de_requirement.py    # Requirement analysis logic
│   ├── make_input_check_pin.py   # Pin configuration logic
│   ├── make_int_med.py     # Intermediate data processing
│   ├── make_judge_check_pin.py   # Automated validation logic
│   ├── make_dcr.py         # Consolidate DCR reporting
│   ├── make_form_measurement.py  # Measurement result integration
│   └── calculate_lsl_usl.py      # Advanced statistical engine
│
└── dist/                   # Deployment artifacts
    └── DCR_Converter.exe   # Standalone executable
```

### Module Specifications

| Module | Line Count | Primary Responsibility |
|:---|:---|:---|
| **main.py** | 18 | Bootstraps the PySide6 application environment. |
| **build_exe.py** | 49 | Configures PyInstaller for single-file deployment. |
| **ui/main_window.py** | ~1,470 | Implements the Material Design UI, multi-tab navigation, and logging. |
| **logic/calculate_lsl_usl.py** | 1,001 | Core statistical engine for 3-sigma analysis and data transposing. |
| **logic/make_form_measurement.py** | 613 | Integrates TDR and Dimension data into measurement reports. |
| **logic/make_judge_check_pin.py** | 367 | Implements complex validation logic for pin assignments. |
| **logic/make_input_check_pin.py** | 347 | Maps netlist configurations to physical pin inputs. |
| **logic/make_dcr.py** | 319 | Final assembly of the Direct Current Resistance report. |
| **logic/make_de_requirement.py** | 310 | Extracts and normalizes part-pin continuity requirements. |
| **logic/file_reader.py** | 200 | Handles robust reading of .NET (UTF-8/CP949) and Excel files. |
| **logic/cover_page.py** | 197 | Generates stylized cover sheets for all Excel outputs. |
| **logic/makevendor.py** | 161 | Processes and styles vendor-specific component data. |
| **logic/config_manager.py** | 111 | Manages the local settings and operator persistence. |

**Total Project Scale: Approximately 5,440 lines of Python code.**

---

## Installation & Deployment

### Development Environment Setup
This project uses **uv** for high-speed dependency management.

```bash
# Install uv (Windows)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# Clone and initialize environment
cd 0105python
uv sync

# Run the application
uv run python main.py
```

### Core Dependencies
- **Python 3.11+**
- **PySide6**: High-fidelity Qt6 GUI framework.
- **openpyxl**: Advanced Excel manipulation (styling, formulas, conditional formatting).
- **pandas & numpy**: Industrial-strength data processing and numerical analysis.
- **chardet**: Intelligent character encoding detection.

### Building Standalone Executable
To package the application into a single `.exe` file for distribution:
```bash
uv run python build_exe.py
```
Output: `dist/DCR_Converter.exe`

---

## Operational Guide

1.  **Identity Initialization**: Enter the **Operator Name** in Tab 1. This name will be embedded in file names, cover pages, and logs.
2.  **Source Selection**: Select the required input files for your specific task using the **Browse** buttons.
3.  **Execution**:
    - For individual reports, click **Execute** within the specific tab.
    - For a full end-to-end process, click **Auto Execute All** at the top.
4.  **Audit Logs**: Monitor the real-time log window. Upon completion, a detailed audit log will be saved as a `.dat` file in the application directory.
5.  **Report Verification**: Open the output Excel files to view the results. Each file will begin with a **Cover Page** for documentation compliance.

### File Naming Convention
- **Tab 1 Output**: `DCR_format_yamaha_{Operator}_{Date}.xlsx`
- **Tab 2 Output**: `Form_measurement_result_{Operator}_{Date}.xlsx`
- **Tab 3 Output**: `Calculate_3Sigma_LSLUSL_final.xlsx`
- **Log Output**: `log_{Operator}_{Date}.dat`

---

## Compliance & Notes

- **Data Integrity**: The software is designed to handle non-numeric data gracefully, converting to numeric types only where valid.
- **Standardization**: The **GND-SUS** (final NET) entry is automatically fixed to **0 (LSL)** and **50 (USL)** across all reports per corporate standard.
- **Security**: This tool is intended for internal use within the **Sumitomo Electric Industries** ecosystem.

---
*Last updated: January 2026*
