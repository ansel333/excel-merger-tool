# Excel Merger Tool

A lightweight and efficient GUI application to merge multiple Excel files by consolidating data from the same sheet across different workbooks.

## Features

- **Multi-file merging**: Combine data from multiple Excel files into a single workbook
- **Sheet selection**: Choose which sheet to merge from the available sheets
- **Smart header handling**: Customize the number of header rows to preserve during merging
- **Empty row filtering**: Automatically removes completely empty rows from the merged result
- **Auto-numbering**: Automatically renumbers the first column if it contains sequential numbers
- **No heavy dependencies**: Built with PyQt6 and openpyxl only (no pandas/numpy overhead)
- **Pure GUI application**: Easy-to-use graphical interface for all operations
- **Column auto-fit**: Automatically adjusts column widths for better readability

## Requirements

- Python 3.8+
- openpyxl >= 3.0.0
- PyQt6 >= 6.0.0

## Installation

### From source

```bash
# Clone the repository
git clone https://github.com/yourusername/excel-merger-tool.git
cd excel-merger-tool

# Create virtual environment (recommended)
python -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Standalone executable

Download the `Excel表格合并工具.exe` from the [Releases](../../releases) page. No installation required - just run it directly.

## Usage

### Using the GUI application

1. **Run the application**:
   ```bash
   python excel_merger_gui.py
   ```
   Or on Windows with standalone executable: Double-click `Excel表格合并工具.exe`

2. **Select Excel files**:
   - Click "📁 Select Files" button
   - Choose one or more Excel files (`.xlsx` or `.xls`)

3. **Read sheets**:
   - Click "📖 Read Sheets" to scan the selected files
   - The first file's sheet list will be displayed

4. **Configure merge settings**:
   - Select the target sheet to merge from the dropdown
   - Set the number of header rows (default: 1)

5. **Merge and save**:
   - Click "✅ Start Merge" to begin the merging process
   - The merged file will be saved with a timestamp in the current directory
   - Output format: `合并结果_{SheetName}_{YYYYMMDD_HHMMSS}.xlsx`

### Python API

```python
from merger import ExcelMerger

# Scan for Excel files
files = ExcelMerger.scan_excel_files("./path/to/files")

# Get sheet names from first file
sheet_names = ExcelMerger.get_sheet_names(files[0])

# Merge sheets
header_rows, data_rows, file_order = ExcelMerger.merge_sheets(
    files, 
    target_sheet="Sheet1", 
    header_rows_count=1
)

# Create output file
output_path = ExcelMerger.create_output_file(
    header_rows, 
    data_rows, 
    "Sheet1", 
    file_order, 
    sheet_names
)
```

## How it works

1. **File scanning**: Recursively scans a directory for Excel files (`.xlsx`, `.xls`)
2. **Sheet detection**: Reads available sheets from the first Excel file
3. **Data extraction**: Extracts data from the specified sheet in all files
4. **Row filtering**: Removes completely empty rows to keep the merged file clean
5. **Smart numbering**: If the first column contains sequential numbers, auto-renumbers them
6. **File creation**: Creates a new workbook with all original sheets and merged data
7. **Auto-formatting**: Adjusts column widths automatically for better readability

## Output

The merged file will contain:
- All original sheets from the source files
- The merged data in the selected sheet
- Preserved header rows from the first file
- Automatically numbered first column (if applicable)
- Auto-adjusted column widths

Example output filename: `合并结果_明细_20251203_153000.xlsx`

## Building from source

To create a standalone executable:

```bash
# Install PyInstaller
pip install pyinstaller

# Build the executable (onefile mode)
pyinstaller --clean --onefile excel_merger.spec
```

The executable will be located at `dist/Excel表格合并工具.exe` (approximately 36 MB)

## Limitations

- Only supports `.xlsx` and `.xls` file formats
- Merges data from the same sheet name across files
- Filters out files with "合并结果" (merge result) in the filename to prevent re-merging

## Performance

- Efficient memory usage: streams rows instead of loading entire files into memory
- Fast processing: typically handles hundreds of files in seconds
- No temporary files created during the merge process

## License

MIT License - see LICENSE file for details

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Troubleshooting

### Empty rows in output
- The tool automatically filters completely empty rows
- Rows with at least one non-empty cell are preserved

### First column numbering not working
- The tool only auto-renumbers if all values in the first column are numeric
- Mixed content (text and numbers) is preserved as-is

### File not found
- Ensure Excel files are in the current working directory or specify the full path
- Check that filenames don't contain "合并结果" (which indicates a previous merge result)

## Support

For issues, questions, or suggestions, please open an [issue](../../issues) on GitHub.

---

Made with ❤️ for efficient data management
