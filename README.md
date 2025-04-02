# Excel Manager Application

A Streamlit-based application for managing Excel files with various read, write, and management operations.

## Features

The Excel Manager application provides a user-friendly interface to perform common Excel operations:

### Sheet Operations
- Create new workbooks and sheets
- Count sheets in a workbook
- Get sheet names
- Delete sheets

### Read Operations
- Read individual cell values
- Read ranges of cells
- **Read Total**: Automatically find a total value by traversing down a column from a starting point

### Write Operations
- Write values to individual cells
- Write ranges of data from CSV input

## Excel Manager Class

The `excelManager` class provides the core functionality for Excel operations:

### Initialization
```python
# Initialize with a file path
manager = excelManager("path/to/file.xlsx")

# Initialize without a file path
manager = excelManager()
```

### Basic Operations
```python
# Create a new workbook
manager.create_workbook("path/to/new_file.xlsx")

# Load an existing workbook
manager.load_workbook("path/to/file.xlsx")

# Save changes
manager.save()

# Close the workbook
manager.close()
```

### Sheet Operations
```python
# Get the number of sheets
count = manager.count_sheets()

# Get all sheet names
names = manager.get_sheet_names()

# Create a new sheet
manager.create_sheet("Sheet Name")

# Get a sheet by name
sheet = manager.get_sheet("Sheet Name")

# Delete a sheet
manager.delete_sheet("Sheet Name")
```

### Read Operations
```python
# Read a cell value (using cell reference)
value = manager.read_cell("Sheet1", "A1")

# Read a cell value (using row and column numbers)
value = manager.read_cell("Sheet1", 1, 1)

# Read a range (using range reference)
values = manager.read_range("Sheet1", "A1:C3")

# Read a range (using start and end cell references)
values = manager.read_range("Sheet1", "A1", "C3")

# Read a range (using row and column numbers)
values = manager.read_range("Sheet1", 1, 1, 3, 3)

# Find a total value by traversing down a column
total = manager.read_total("Sheet1", "A1")  # Using cell reference
total = manager.read_total("Sheet1", 1, 1)  # Using row and column numbers
```

### Write Operations
```python
# Write a cell value (using cell reference)
manager.write_cell("Sheet1", "A1", "Value")

# Write a cell value (using row and column numbers)
manager.write_cell("Sheet1", 1, 1, "Value")

# Write a range (using cell reference)
data = [["A", "B", "C"], [1, 2, 3], [4, 5, 6]]
manager.write_range("Sheet1", "A1", data)

# Write a range (using row and column numbers)
manager.write_range("Sheet1", 1, 1, data)
```

## Special Features

### Read Total Method

The `read_total` method is particularly useful for finding totals in financial documents or reports. This method:

1. Starts at a specified cell
2. Traverses down the column until it finds an empty cell
3. Returns the value from the last non-empty cell (which is typically a total)

Usage example:
```python
# Find a total at the end of a column of values starting from F25
total = manager.read_total("Sheet1", "F25")
```

## Installation

1. Clone the repository
2. Install dependencies:
```
pip install -r requirements.txt
```
3. Run the application:
```
streamlit run excel_app.py
```

## Requirements

- Python 3.6+
- streamlit
- openpyxl
- pandas