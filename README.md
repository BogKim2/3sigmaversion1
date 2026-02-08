# DCR Format Converter

<p align="center">
  <img src="https://upload.wikimedia.org/wikipedia/en/thumb/4/4e/Sumitomo_Electric_Industries_logo.svg/200px-Sumitomo_Electric_Industries_logo.svg.png" alt="Sumitomo Electric Industries" width="200"/>
</p>

<p align="center">
  <strong>DCR Format Conversion & 3-Sigma LSL/USL Calculation Tool</strong><br>
  <em>Version 1.0 | Programmed by Sangwoo Kim</em>
</p>

---

## 📋 Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [System Requirements](#system-requirements)
4. [Installation](#installation)
5. [Quick Start Guide](#quick-start-guide)
6. [Detailed User Guide](#detailed-user-guide)
7. [Program Architecture](#program-architecture)
8. [File Structure](#file-structure)
9. [Input/Output File Specifications](#inputoutput-file-specifications)
10. [Troubleshooting](#troubleshooting)
11. [FAQ](#faq)
12. [License & Credits](#license--credits)

---

## Overview

**DCR Format Converter** is a professional desktop application designed for Sumitomo Electric Industries' quality control and measurement data processing workflows. The tool automates the complex process of:

- Converting NET files and Excel data into standardized DCR format
- Processing Form Measurement Result data with TDR/Impedance analysis
- Calculating 3-Sigma LSL (Lower Specification Limit) and USL (Upper Specification Limit) statistics

This application eliminates hours of manual Excel manipulation and ensures consistent, error-free data processing.

---

## Features

### Tab 1: Make DCR Format
- ✅ Parse `.NET` files (network topology files)
- ✅ Extract vendor specifications from Excel
- ✅ Generate DE Requirement sheets
- ✅ Create Input Check Pin sheets with formulas
- ✅ Build Judge (check pin) evaluation sheets
- ✅ Produce final DCR sheets

### Tab 2: Make Form Measurement Result File
- ✅ Process DK files (DK 1.5, DK 1.6, ... DK CENTER)
- ✅ Extract TDR (Time Domain Reflectometry) data
- ✅ Fill Impedance NET resistance values
- ✅ Process Dimension data (Circuit Width/Thickness)
- ✅ Apply LSL/Center/USL specifications
- ✅ Auto-calculate Min/Max/Average with conditional formatting

### Tab 3: Calculate LSL USL
- ✅ Process merged measurement files
- ✅ Calculate comprehensive statistics:
  - Min, Max, Average, Median, Stdev
  - IQR (Interquartile Range)
  - 1st Quartile - 4×IQR, 3rd Quartile + 4×IQR
  - AverageIfs, StdevIfs (outlier-filtered)
  - LSL = A - 3B, USL = A + 3B
- ✅ Generate visualization plots (PNG files)

### Additional Features
- 🎨 Material Design UI with Sumitomo branding
- 📊 Real-time progress logging
- 💾 Automatic configuration persistence
- 📁 Organized output directory structure
- 📈 Matplotlib-based data visualization
- 🔄 Auto Execute All functionality

---

## System Requirements

| Component | Minimum | Recommended |
|-----------|---------|-------------|
| **OS** | Windows 10 (64-bit) | Windows 11 (64-bit) |
| **RAM** | 4 GB | 8 GB+ |
| **Disk Space** | 500 MB | 1 GB |
| **Python** | 3.12+ | 3.12+ |
| **Display** | 1280×720 | 1920×1080 |

### For Standalone Executable
- No Python installation required
- Download and run `DCR_Converter.exe` directly

---

## Installation

### Option A: Run Standalone Executable (Recommended for End Users)

1. Download `DCR_Converter.exe` from the `dist/` folder
2. Copy `files.json` configuration file to the same directory
3. Double-click to run

### Option B: Run from Source (For Developers)

#### Step 1: Install UV Package Manager
```bash
# Windows (PowerShell)
powershell -c "irm https://astral.sh/uv/install.ps1 | iex"

# Or using pip
pip install uv
```

#### Step 2: Clone and Setup
```bash
# Navigate to project directory
cd G:\dback\00_sumitomo_sangwoo\pythonprogram\0105python

# Create virtual environment and install dependencies
uv sync

# Activate virtual environment
.venv\Scripts\activate
```

#### Step 3: Run the Application
```bash
python main.py
```

### Option C: Build Executable from Source

```bash
# Make sure you're in the project directory with venv activated
python build_exe.py
```

The executable will be created in the `dist/` folder.

---

## Quick Start Guide

### 🚀 5-Minute Quick Start

1. **Launch** the application (`DCR_Converter.exe` or `python main.py`)

2. **Enter Operator Name** at the top of the window

3. **Tab 1 - DCR Format:**
   - Click "Browse" to select your `.NET` file
   - Click "Browse" to select vendorspec Excel file
   - Click "Browse" to select partpin Excel file
   - Click "Execute" button

4. **Tab 2 - Form Measurement:**
   - Click "Browse" to select etching directory (containing DK files)
   - Click "Browse" to select dimension file
   - Click "Browse" to select LSLUSL file
   - Click "Execute" button

5. **Tab 3 - Calculate LSL USL:**
   - Click "Browse" to select merged file
   - Click "Execute" button

6. **Or simply click "Auto Execute All"** to run all tabs sequentially!

7. **Find your outputs** in the `output/` folder

---

## Detailed User Guide

### Starting the Application

When you launch the application, you'll see a professional interface with:
- **Header**: Sumitomo Electric logo and program version
- **Operator Input**: Enter your name here (required)
- **Three Tabs**: Each for different processing stages
- **Auto Execute All Button**: Runs all tabs in sequence

### Understanding the Tabs

#### Tab 1: Make DCR Format

This tab converts raw network and specification files into the standardized DCR format.

**Required Input Files:**

| File Type | Description | Example |
|-----------|-------------|---------|
| NET File | Network topology definition | `7S3493-994110-6CRMEV.NET` |
| Vendorspec File | Vendor specifications Excel | `099-51612-03-Sumitomo.xlsx` |
| Partpin File | Part-pin mapping Excel | `6CRMEV-P2-POR_回路図 - FPC r1.xlsx` |

**Processing Steps:**
1. **Make Vendor Sheet** - Copies vendor specifications
2. **Make DE Requirement Sheet** - Extracts 4-wire pair information
3. **Make Input Check Pin Sheet** - Creates pin checking structure
4. **Create int_med.xlsx** - Intermediate 4W grouping file
5. **Create Input Check Pin Final** - Merges all pin data
6. **Create Judge (check pin)** - Adds judgment formulas
7. **Create DCR Sheet** - Final DCR format output
8. **Add Cover Page** - Adds metadata and traceability

**Output:**
- `DCR_format_yamaha_{Operator}_{Date}.xlsx`

#### Tab 2: Make Form Measurement Result File

This tab processes measurement data from etching tests.

**Required Input Files:**

| File Type | Description | Example |
|-----------|-------------|---------|
| Etching Directory | Folder containing DK*.xls files | `etching/` |
| Dimension File | Circuit width/thickness data | `7E3493-00003.xlsx` |
| LSLUSL File | LSL/USL specification values | `LSLUSL.xlsx` |

**DK Files Expected:**
- `DK1.5.xls`, `DK1.6.xls`, `DK1.7.xls`, `DK1.9.xls`
- `DK2.1.xls`, `DK2.3.xls`, `DK2.4.xls`, `DK2.5.xls`
- `DK CENTER.xls`

**Processing Steps:**
1. **Create Form Measurement File** - Copies template
2. **Fill Impedance Data** - Reads TDR values from DK files
3. **Fill Dimension Data** - Adds circuit width/thickness
4. **Fill LSL/USL Data** - Applies specification limits
5. **Add Cover Page** - Adds metadata

**Output:**
- `Form_measurement_result_{Operator}_{Date}.xlsx`

#### Tab 3: Calculate LSL USL

This tab performs comprehensive statistical analysis.

**Required Input Files:**

| File Type | Description | Example |
|-----------|-------------|---------|
| Merged File | Combined measurement data | `merged_file.xlsx` |
| DCR File | Output from Tab 1 (auto-selected) | Auto-detected |

**Processing Steps:**
1. **Read Merged Data** - Filters Method=3 data
2. **Create Cal_merged Sheet** - Reorganizes by NET count
3. **Create Sap xep Sheet** - Transposes data matrix
4. **Create tinh LCLUCL Sheet** - Calculates all statistics
5. **Create Calculate USL LSL Sheet** - Final summary with DCR data
6. **Generate Plots** - Creates visualization PNG files
7. **Add Cover Page** - Adds metadata

**Statistical Formulas Used:**
```
Min = MIN(data_range)
Max = MAX(data_range)
Average = AVERAGE(data_range)
Median = MEDIAN(data_range)
Stdev = STDEV(data_range)
IQR = QUARTILE(data, 3) - QUARTILE(data, 1)
1stQ-4IQR = MAX(0, QUARTILE(data, 1) - 4×IQR)
3rdQ+4IQR = QUARTILE(data, 3) + 4×IQR
AverageIfs = AVERAGEIFS(data, ">1stQ-4IQR", "<3rdQ+4IQR")
StdevIfs = STDEV of filtered data
LSL = AverageIfs - 3×StdevIfs
USL = AverageIfs + 3×StdevIfs
```

**Output:**
- `Calculate_3Sigma_LSLUSL_final.xlsx`
- `output/plots/` - PNG visualization files

### Using Auto Execute All

Click the **"Auto Execute All"** button to:
1. Automatically switch between tabs
2. Execute all three processing stages
3. Generate all output files
4. Create a comprehensive log file

**Note:** Make sure all input files are selected before using this feature.

### Understanding Output Files

All outputs are saved in the `output/` directory:

```
output/
├── DCR_format_yamaha_{Operator}_{Date}.xlsx
├── Form_measurement_result_{Operator}_{Date}.xlsx
├── Calculate_3Sigma_LSLUSL_final.xlsx
├── log_{Operator}_{Date}.dat
└── plots/
    ├── TDR_BoxPlot_{timestamp}.png
    ├── Dimension_BarChart_{timestamp}.png
    ├── LSLUSL_Control_{timestamp}.png
    └── LSLUSL_Histogram_NET{n}_{timestamp}.png
```

### Configuration File (files.json)

The application remembers your file selections:

```json
{
    "net_file": "path/to/file.NET",
    "xlsx_file": "path/to/vendorspec.xlsx",
    "partpin_file": "path/to/partpin.xlsx",
    "etching_directory": "path/to/etching/",
    "dimension_file": "path/to/dimension.xlsx",
    "lslusl_file": "path/to/LSLUSL.xlsx",
    "merged_file": "path/to/merged_file.xlsx",
    "operator": "Your Name"
}
```

---

## Program Architecture

### High-Level Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                      main.py (Entry Point)                  │
└─────────────────────────┬───────────────────────────────────┘
                          │
┌─────────────────────────▼───────────────────────────────────┐
│                   ui/main_window.py                         │
│  ┌─────────────┬─────────────┬─────────────┬──────────────┐ │
│  │   Tab 1     │   Tab 2     │   Tab 3     │ Auto Execute │ │
│  │ DCR Format  │Form Measure │ LSL/USL Calc│    All       │ │
│  └──────┬──────┴──────┬──────┴──────┬──────┴──────────────┘ │
└─────────┼─────────────┼─────────────┼───────────────────────┘
          │             │             │
┌─────────▼─────────────▼─────────────▼───────────────────────┐
│                    logic/ (Business Logic)                  │
│  ┌────────────────┐ ┌────────────────┐ ┌──────────────────┐ │
│  │ makevendor.py  │ │make_form_      │ │calculate_lsl_   │ │
│  │ make_de_req... │ │measurement.py  │ │usl.py           │ │
│  │ make_input_... │ │                │ │                  │ │
│  │ make_int_med.. │ │                │ │                  │ │
│  │ make_judge_... │ │                │ │                  │ │
│  │ make_dcr.py    │ │                │ │                  │ │
│  └────────────────┘ └────────────────┘ └──────────────────┘ │
│  ┌────────────────┐ ┌────────────────┐ ┌──────────────────┐ │
│  │ file_reader.py │ │ cover_page.py  │ │ visualizer.py    │ │
│  │config_manager  │ │                │ │ (matplotlib)     │ │
│  └────────────────┘ └────────────────┘ └──────────────────┘ │
└─────────────────────────────────────────────────────────────┘
          │
┌─────────▼───────────────────────────────────────────────────┐
│              External Libraries (Dependencies)              │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌────────────────┐  │
│  │ PySide6  │ │ openpyxl │ │ pandas   │ │ matplotlib     │  │
│  │ (Qt GUI) │ │ (Excel)  │ │ (Data)   │ │ (Plots)        │  │
│  └──────────┘ └──────────┘ └──────────┘ └────────────────┘  │
└─────────────────────────────────────────────────────────────┘
```

### Module Descriptions

#### Entry Point
| File | Description |
|------|-------------|
| `main.py` | Application entry point, initializes QApplication and MainWindow |

#### UI Layer (`ui/`)
| File | Description |
|------|-------------|
| `main_window.py` | Main GUI window with tabs, buttons, and event handlers |

#### Logic Layer (`logic/`)
| File | Description | Lines |
|------|-------------|-------|
| `makevendor.py` | Copies vendor specification sheet | ~100 |
| `make_de_requirement.py` | Creates DE requirement from partpin data | ~200 |
| `make_input_check_pin.py` | Builds input check pin structure | ~300 |
| `make_int_med.py` | Processes NET file for 4W groups | ~200 |
| `make_judge_check_pin.py` | Creates judgment formulas | ~150 |
| `make_dcr.py` | Generates final DCR sheet | ~250 |
| `make_form_measurement.py` | Processes form measurement data | ~400 |
| `calculate_lsl_usl.py` | Statistical calculations | ~500 |
| `file_reader.py` | Reads NET and Excel files | ~100 |
| `config_manager.py` | Saves/loads JSON configuration | ~50 |
| `cover_page.py` | Adds cover page metadata | ~100 |
| `visualizer.py` | Generates matplotlib plots | ~200 |

### Data Flow

```
Input Files                    Processing                      Output Files
────────────                   ──────────                      ────────────

┌──────────┐                                              ┌──────────────────┐
│ .NET     │──┐                                           │DCR_format_yamaha_│
│ file     │  │    ┌─────────────────────────────┐        │{Op}_{Date}.xlsx  │
└──────────┘  ├───►│                             │───────►└──────────────────┘
┌──────────┐  │    │        Tab 1 Logic          │
│vendorspec│──┤    │    (7 sequential steps)     │        ┌──────────────────┐
│ .xlsx    │  │    │                             │        │  int_med.xlsx    │
└──────────┘  │    └─────────────────────────────┘        └──────────────────┘
┌──────────┐  │
│ partpin  │──┘
│ .xlsx    │
└──────────┘

┌──────────┐                                              ┌──────────────────┐
│DK*.xls   │──┐                                           │Form_measurement_ │
│(multiple)│  │    ┌─────────────────────────────┐        │result_{Op}_{D}.  │
└──────────┘  ├───►│                             │───────►│xlsx              │
┌──────────┐  │    │        Tab 2 Logic          │        └──────────────────┘
│dimension │──┤    │    (4 sequential steps)     │
│ .xlsx    │  │    │                             │        ┌──────────────────┐
└──────────┘  │    └─────────────────────────────┘        │ plots/*.png      │
┌──────────┐  │                                           └──────────────────┘
│ LSLUSL   │──┘
│ .xlsx    │
└──────────┘

┌──────────┐                                              ┌──────────────────┐
│merged_   │──┐                                           │Calculate_3Sigma_ │
│file.xlsx │  │    ┌─────────────────────────────┐        │LSLUSL_final.xlsx │
└──────────┘  ├───►│                             │───────►└──────────────────┘
┌──────────┐  │    │        Tab 3 Logic          │
│DCR output│──┘    │  (statistical analysis)     │        ┌──────────────────┐
│(from T1) │       │                             │        │ plots/*.png      │
└──────────┘       └─────────────────────────────┘        └──────────────────┘
```

---

## File Structure

```
0105python/
│
├── main.py                    # Application entry point
├── build_exe.py               # PyInstaller build script
├── pyproject.toml             # Project dependencies
├── files.json                 # User configuration
│
├── ui/                        # User Interface
│   ├── __init__.py
│   └── main_window.py         # Main window implementation
│
├── logic/                     # Business Logic
│   ├── __init__.py
│   ├── makevendor.py          # Vendor sheet creation
│   ├── make_de_requirement.py # DE requirement processing
│   ├── make_input_check_pin.py# Input check pin creation
│   ├── make_int_med.py        # Intermediate file creation
│   ├── make_judge_check_pin.py# Judge sheet creation
│   ├── make_dcr.py            # DCR sheet creation
│   ├── make_form_measurement.py# Form measurement processing
│   ├── calculate_lsl_usl.py   # LSL/USL calculations
│   ├── file_reader.py         # File reading utilities
│   ├── config_manager.py      # Configuration management
│   ├── cover_page.py          # Cover page generation
│   └── visualizer.py          # Plot generation
│
├── output/                    # Generated output files
│   ├── *.xlsx                 # Excel outputs
│   ├── *.dat                  # Log files
│   └── plots/                 # PNG visualizations
│
├── dist/                      # Built executable
│   └── DCR_Converter.exe
│
├── 301_3sigma CAL/            # Sample input files
│   ├── etching/               # DK files directory
│   ├── *.xlsx                 # Reference Excel files
│   └── *.NET                  # Network files
│
└── Form measurement result files_form.xlsx  # Template file
```

---

## Input/Output File Specifications

### Input File Formats

#### .NET File Format
```
#CONT
...continuity data...
%END

#4W
#Gr01
EXR4W 1,2,2049,2050
EXR4W 3,4,2051,2052
...
%END
```

#### DK Excel File (Form kq sheet)
| Column | Content |
|--------|---------|
| STT | Sequence number |
| TDR | TDR measurement value |
| ... | Other measurements |

### Output File Formats

#### DCR_format_yamaha.xlsx Sheets
1. **Cover Page** - Metadata
2. **vendor** - Vendor specifications
3. **DE requirement** - Design engineering requirements
4. **input check pin interm** - Intermediate pin data
5. **input check pin** - Final pin check data
6. **Judge(check pin)** - Judgment results
7. **DCR** - Final DCR format

#### Calculate_3Sigma_LSLUSL.xlsx Sheets
1. **Cover Page** - Metadata
2. **merged_file** - Raw filtered data
3. **Cal_merged** - Transposed data
4. **Sap xep** - Reorganized data
5. **tinh LCLUCL** - Statistical calculations
6. **Calculate USL LSL** - Final summary

---

## Troubleshooting

### Common Issues

#### 1. "Template file not found" Error
**Cause:** The template file `Form measurement result files_form.xlsx` is missing.

**Solution:**
- For source: Ensure the template file is in the project root
- For executable: Place the template in the same folder as the .exe

#### 2. Application Crashes on Startup
**Cause:** Missing dependencies or corrupted installation.

**Solution:**
```bash
# Reinstall dependencies
uv sync --reinstall
```

#### 3. "Invalid column index" Error
**Cause:** Excel file has too many columns (>16384).

**Solution:** This usually indicates a data processing error. Check your merged_file.xlsx for issues.

#### 4. Empty Output Files
**Cause:** Input files may have unexpected format.

**Solution:**
- Verify input file formats match expected structure
- Check the log file for detailed error messages

#### 5. Japanese Characters Display Incorrectly
**Cause:** Encoding issues with .NET files.

**Solution:** The application uses `chardet` for auto-detection. Ensure files are saved in UTF-8 or Shift-JIS encoding.

### Log File Location
Check the log file in `output/log_{Operator}_{Date}.dat` for detailed processing information.

---

## FAQ

**Q: Can I run this on Mac or Linux?**
A: The Python source code is cross-platform, but the GUI is optimized for Windows. PySide6 works on all platforms.

**Q: How long does processing take?**
A: Tab 1 and 2 take a few seconds. Tab 3 can take 1-5 minutes depending on the size of merged_file.xlsx (can be 100,000+ rows).

**Q: Can I customize the output format?**
A: Yes, modify the files in the `logic/` directory. Each module handles specific output formatting.

**Q: What happens if I run without selecting all files?**
A: The application will skip steps with missing inputs and continue with available data.

**Q: How do I update the application?**
A: Replace the `dist/DCR_Converter.exe` with the new version, or `git pull` for source updates.

---

## License & Credits

### Credits

- **Developer:** Sangwoo Kim
- **AI Assistance:** Claude Opus 4.5 (Anthropic), Gemini (Google)
- **Company:** Sumitomo Electric Industries, Ltd.
- **Version:** 1.0

### Dependencies

| Library | Version | License |
|---------|---------|---------|
| PySide6 | 6.10.1+ | LGPL |
| openpyxl | 3.1.5+ | MIT |
| pandas | 2.3.3+ | BSD |
| matplotlib | 3.9.0+ | PSF |
| xlrd | 2.0.2+ | BSD |
| chardet | 5.2.0+ | LGPL |
| PyInstaller | 6.17.0+ | GPL |

### Contact

For bug reports or feature requests, please contact the development team.

---

<p align="center">
  <em>© 2026 Sumitomo Electric Industries, Ltd. All rights reserved.</em>
</p>
