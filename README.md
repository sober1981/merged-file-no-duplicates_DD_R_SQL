# Drilling Scorecard Data Processing Scripts

## Overview

This project provides automated tools for merging and cleaning drilling scorecard data from multiple Excel sources. The system identifies and removes duplicate run entries based on sophisticated matching criteria including job numbers, hours, drill distance, and serial numbers.

## Version

**Version 2.1** - September 2025

## Scripts

### 1. merge_excel_files_auto.py

**Purpose**: Merges 4 Excel source files into a single standardized dataset with enhanced data processing.

**Input Files Required**:
- Motor KPI* (Motor performance data - Directional reference)
- CAM Run Tracker* (CAM run tracking - Rental reference)
- POG CAM Usage* (POG CAM usage data)
- POG MM Usage* (POG MM usage data)
- FORMAT GRAL TABLE.xlsx (Column mapping template)
- LISTS_BASIN AND FORM_FAM.xlsx (Lookup tables for basin and formation family)

**Key Features**:
- Auto-detects files by pattern matching (uses * wildcard)
- Calculates MOTOR_MODEL from Serial Number (TDI motors) or MOTOR_OD (Non-TDI)
- Converts TIME_IN/TIME_OUT from string format and combines with dates
- Applies text formatting to MOTOR_MODEL, BEND, BEND_HSG columns
- Populates JOB_TYPE for Motor KPI rows (sets to "Directional")
- Adds SOURCE column to track data origin

**Output**:
- MERGED_DATA_YYYYMMDD_HHMMSS.xlsx

**Usage**:
```bash
python merge_excel_files_auto.py
```

---

### 2. detect_duplicates.py

**Purpose**: Identifies and highlights ALL duplicates (both Directional and Rental) for manual review.

**Duplicate Detection Criteria** (ALL must match):
1. **JOB_NUM**: Must match exactly
2. **Hours/Drill Tolerance**:
   - Primary: Total Hrs within ±5 hours tolerance
   - Fallback: If Total Hrs is blank/zero, uses TOTAL_DRILL within ±5 tolerance
3. **Serial Number**: Last 3 digits must match

**Special Features**:
- **Consolidated Run Detection**: If multiple reference rows exist with same JOB_NUM + SN (last 3 digits), the script sums their hours and compares against POG row
- **TOTAL_DRILL Fallback**: Handles cases where Total Hrs is blank but TOTAL_DRILL has values

**Reference Files** (Never Removed):
- Motor KPI (SOURCE='Motor_KPI'): Reference for ALL Directional runs
- CAM Run Tracker (SOURCE='CAM_Run_Tracker'): Reference for ALL Rental runs

**Processing Logic**:
- Removes rows with BOTH Total Hrs = 0/blank AND TOTAL_DRILL = 0/blank
- Highlights ALL duplicate POG rows in YELLOW for review
- Keeps duplicates in the output file

**Output**:
- DUPLICATES_DETECTED_YYYYMMDD_HHMMSS.xlsx (all duplicates highlighted in yellow)

**Usage**:
```bash
python detect_duplicates.py
```

---

### 3. clean_dd_merge.py

**Purpose**: Removes Directional duplicates but KEEPS and highlights Rental duplicates for review.

**Duplicate Detection**: Same criteria as detect_duplicates.py

**Processing Logic**:
- Removes rows with BOTH Total Hrs = 0/blank AND TOTAL_DRILL = 0/blank
- **Directional duplicates**: POG rows matching Motor KPI are REMOVED
- **Rental duplicates**: POG rows matching CAM Run Tracker are KEPT and HIGHLIGHTED in YELLOW

**Output**:
- CLEAN_DD_MERGE_YYYYMMDD_HHMMSS.xlsx
  - Contains: ALL Motor KPI + ALL CAM Run Tracker + Non-duplicate POG Directional + Highlighted POG Rental duplicates

**Usage**:
```bash
python clean_dd_merge.py
```

**Use Case**: When you want clean Directional data but need to manually review Rental duplicates before removal.

---

### 4. clean_dd_r_merge.py

**Purpose**: Removes ALL duplicates (both Directional AND Rental) for a completely clean dataset.

**Duplicate Detection**: Same criteria as detect_duplicates.py

**Processing Logic**:
- Removes rows with BOTH Total Hrs = 0/blank AND TOTAL_DRILL = 0/blank
- **Directional duplicates**: POG rows matching Motor KPI are REMOVED
- **Rental duplicates**: POG rows matching CAM Run Tracker are REMOVED

**Output**:
- CLEAN_DD_R_MERGE_YYYYMMDD_HHMMSS.xlsx
  - Contains: ALL Motor KPI + ALL CAM Run Tracker + Only unique POG rows (no duplicates)

**Usage**:
```bash
python clean_dd_r_merge.py
```

**Use Case**: When you need a completely clean dataset with no duplicates at all.

---

## Workflow

### Standard Processing Pipeline

```
Step 1: Merge Source Files
├── Run: merge_excel_files_auto.py
└── Output: MERGED_DATA_*.xlsx

Step 2: Choose Duplicate Handling Strategy

Option A: Review All Duplicates
├── Run: detect_duplicates.py
└── Output: DUPLICATES_DETECTED_*.xlsx (all duplicates highlighted)

Option B: Clean Directional, Review Rental
├── Run: clean_dd_merge.py
└── Output: CLEAN_DD_MERGE_*.xlsx (Rental duplicates highlighted)

Option C: Remove All Duplicates
├── Run: clean_dd_r_merge.py
└── Output: CLEAN_DD_R_MERGE_*.xlsx (no duplicates)
```

---

## Duplicate Detection Logic Details

### Matching Algorithm

The scripts use a two-phase approach to detect duplicates:

#### Phase 1: Consolidated Run Detection (NEW)
For cases where POG files consolidate multiple reference runs into a single row:

1. Filter reference rows by JOB_NUM
2. Group by Serial Number (last 3 digits match)
3. If multiple reference rows found with same JOB_NUM + SN:
   - **Sum Total Hrs** from all matching reference rows
   - Compare sum against POG Total Hrs (±5 hours tolerance)
   - If Total Hrs blank: **Sum TOTAL_DRILL** and compare (±5 tolerance)
4. If match found → Mark as duplicate

**Example**:
```
Motor KPI:      JOB 21467, SN ending 006, 23.46 hrs
CAM Run Tracker: JOB 21467, SN ending 006, 20.14 hrs
POG Row:        JOB 21467, SN ending 006, 43.60 hrs
Sum: 23.46 + 20.14 = 43.60 hrs → DUPLICATE DETECTED
```

#### Phase 2: Individual Row Matching
If no consolidated match found, check each reference row individually:

1. JOB_NUM must match exactly
2. Hours/Drill comparison:
   - If either Total Hrs > 0: Check Total Hrs tolerance (±5 hrs)
   - If both Total Hrs blank/zero: Fall back to TOTAL_DRILL tolerance (±5)
3. Last 3 digits of Serial Number must match
4. If all criteria match → Mark as duplicate

**Example with TOTAL_DRILL Fallback**:
```
Reference: JOB 21124, SN ending 123, Total Hrs blank, TOTAL_DRILL 150
POG Row:   JOB 21124, SN ending 123, Total Hrs blank, TOTAL_DRILL 152
Difference: |150 - 152| = 2 (within ±5 tolerance) → DUPLICATE DETECTED
```

### Configuration

```python
TOTAL_HRS_TOLERANCE = 5  # ±5 hours/drill tolerance
SN_LAST_DIGITS = 3       # Match last 3 digits of Serial Number
```

---

## Setup Instructions

### Prerequisites

```bash
pip install pandas openpyxl numpy
```

### Required Files

Each working folder must contain:

1. **Source Data Files**:
   - Motor KPI*.xlsx
   - CAM Run Tracker*.xlsx
   - POG CAM Usage*.xlsx
   - POG MM Usage*.xlsx

2. **Lookup/Template Files**:
   - FORMAT GRAL TABLE.xlsx
   - LISTS_BASIN AND FORM_FAM.xlsx

3. **Python Scripts**:
   - merge_excel_files_auto.py
   - detect_duplicates.py
   - clean_dd_merge.py
   - clean_dd_r_merge.py

### Folder Setup for New Time Period

1. Create new folder (e.g., "Scorecard Q4 2024")
2. Copy all 4 Python scripts to the new folder
3. Copy FORMAT GRAL TABLE.xlsx to the new folder
4. Copy LISTS_BASIN AND FORM_FAM.xlsx to the new folder
5. Add your 4 source data files for the time period
6. Run scripts in sequence (merge first, then choose duplicate handling)

---

## Output File Naming Convention

All output files include timestamp for version control:

- `MERGED_DATA_20250101_143022.xlsx`
- `DUPLICATES_DETECTED_20250101_143530.xlsx`
- `CLEAN_DD_MERGE_20250101_144015.xlsx`
- `CLEAN_DD_R_MERGE_20250101_144530.xlsx`

Format: `FILENAME_YYYYMMDD_HHMMSS.xlsx`

---

## Summary Reports

All scripts provide detailed console output including:

- Files processed and row counts
- Number of empty runs removed
- Duplicate detection statistics by type (Directional/Rental)
- Final row counts and file locations
- Processing time and status

Example summary:
```
======================================================================
CLEAN DD MERGE - SUMMARY
======================================================================

Original merged file rows:                  642
Rows removed (no hrs & no drill):           52
Rows after empty removal:                   590

Directional duplicates (REMOVED):           87
Rental duplicates (KEPT):                   39
Total duplicates detected:                  126

Final row count in output:                  542
Clean rows (no duplicates):                 503

Output file: CLEAN_DD_MERGE_20250101_144015.xlsx

NOTE: Directional duplicates have been REMOVED from the output.
      Rental duplicates have been KEPT and HIGHLIGHTED in YELLOW.
======================================================================
```

---

## Troubleshooting

### Common Issues

**Issue**: "No file found matching pattern"
- **Solution**: Verify file names start with exact pattern (Motor KPI, CAM Run Tracker, etc.)

**Issue**: "Column not found error"
- **Solution**: Ensure FORMAT GRAL TABLE.xlsx has correct column mappings

**Issue**: "Empty Motor KPI rows"
- **Solution**: Script auto-detects header row location; if issues persist, check file structure

**Issue**: "Unicode error with arrow character"
- **Solution**: Already fixed in current version (uses '->' instead of '→')

**Issue**: "Wrong rows highlighted"
- **Solution**: Already fixed - uses reset_index() for proper Excel row mapping

---

## Technical Details

### Data Processing Steps

1. **File Detection**: Uses glob pattern matching with wildcards
2. **Header Detection**: Auto-detects if Motor KPI has headers in first row
3. **Column Mapping**: Applies standardized column names from FORMAT GRAL TABLE
4. **Time Conversion**: Converts string times ('09:00:00') and combines with dates
5. **MOTOR_MODEL Calculation**:
   - TDI motors: Extracts from Serial Number
   - Non-TDI: Uses MOTOR_OD value
6. **Text Formatting**: Applies text format to specific columns
7. **Empty Run Removal**: Filters out rows with no hours AND no drill distance
8. **Duplicate Detection**: Multi-phase matching with summing logic
9. **Excel Formatting**: Highlights duplicates with yellow fill (PatternFill)

### Performance

- Typical processing time: 5-10 seconds for ~600 rows
- Memory efficient: Processes files in-place with minimal duplication
- Scalable: Handles datasets with thousands of rows

---

## Version History

### Version 2.1 (Current) - September 2025
- Added TOTAL_DRILL fallback criteria for duplicate detection
- Implemented consolidated run detection (summing logic)
- Fixed DataFrame index vs Excel row position bug
- Enhanced error handling and reporting
- Improved documentation

### Version 1.1
- Added three duplicate handling variants
- Implemented reference file identification by SOURCE
- Added yellow highlighting for duplicates
- Fixed Motor KPI header detection

### Version 1.0
- Initial merge functionality
- Basic duplicate detection
- MOTOR_MODEL calculation

---

## Contributing

For bug reports, feature requests, or contributions, please contact the development team.

---

## License

Internal use only - Scout Downhole Drilling Optimization Projects

---

## Contact

For questions or support, contact: Jesse Soberanes
Project: Drilling Scorecard Data Processing
Department: Drilling Optimization

---

## Notes

- Reference files (Motor KPI, CAM Run Tracker) are NEVER removed from output
- All timestamps are in local system time
- Excel files use .xlsx format (OpenPyXL engine)
- Yellow highlighting uses hex color #FFFF00
- Scripts are designed to be run sequentially in a single working directory
- Always keep backup copies of source files before processing

---

**Last Updated**: September 2025
**Script Version**: 2.1
**Author**: Jesse Soberanes (with AI assistance)
