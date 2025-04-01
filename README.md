# excel_manager

## Overview
`excel_manager.py` is a Python module designed to simplify the management of Excel files using the `openpyxl` library. It provides a convenient interface for creating, loading, modifying, and saving Excel workbooks.

## Features
- **Workbook Management**: Create new workbooks, load existing ones, save changes, and close workbooks.
- **Sheet Operations**: Count sheets, retrieve sheet names, create new sheets, access specific sheets, and delete sheets.
- **Cell Operations**: Read and write individual cells, including support for following simple cell reference formulas.
- **Range Operations**: Read and write ranges of cells efficiently.

## Usage
### Initialization
```python
from excel_manager import ExcelManager

# Initialize with an existing file
manager = ExcelManager('existing_file.xlsx')

# Or create a new workbook
manager = ExcelManager('new_file.xlsx')
```

### Workbook Operations
```python
# Save workbook
manager.save()

# Close workbook
manager.close()
```

### Sheet Operations
```python
# Create a new sheet
manager.create_sheet('Sheet1')

# Get sheet names
sheet_names = manager.get_sheet_names()

# Delete a sheet
manager.delete_sheet('Sheet1')
```

### Cell Operations
```python
# Write to a cell
manager.write_cell('Sheet1', 1, 1, 'Hello')

# Read from a cell
value = manager.read_cell('Sheet1', 1, 1)
```

### Range Operations
```python
# Read a range of cells
values = manager.read_range('Sheet1', 1, 1, 3, 3)
```

## Logging
The module includes comprehensive logging to track operations and errors, outputting logs to both the console and a file named `excel_manager.log`.

## Dependencies
- `openpyxl`

Ensure dependencies are installed using:
```bash
pip install -r requirements.txt
```

## License
Specify your project's license here.