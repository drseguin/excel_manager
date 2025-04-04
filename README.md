# Excel Manager

A Python class for managing Excel files with various read, write, and sheet management operations.

## Overview

The `excelManager` class provides a powerful yet simple interface for interacting with Excel workbooks. It supports reading calculated values from formulas, maintaining formatting information, and handling various Excel operations.

## Features

- Create, load, and save Excel workbooks
- Manage sheets (create, delete, list)
- Read individual cells and ranges with calculated values
- Write to cells and ranges
- Find totals at the end of columns 
- Find totals by column title
- Extract consecutive items from columns with offset capability
- Extract multiple columns by title or cell reference
- Support for A1 notation and row/column indices
- Consistent error handling and logging
- Currency and numeric formatting support

## Installation

1. Clone the repository
2. Install the requirements:
```bash
pip install -r requirements.txt
```

## Requirements

- streamlit
- openpyxl
- pandas

## Excel Manager Class Details

The Excel Manager maintains two copies of each workbook internally - a data-only version for reading calculated values from formulas and a formula version for maintaining the original formulas and formatting when writing. This dual approach ensures that you get the calculated values when reading while preserving the original structure.

### Creating an Instance

```python
from excel_manager import excelManager

# Create a new instance with no file
excel = excelManager()

# Create a new instance with an existing file
excel = excelManager("path/to/existing_file.xlsx")

# Create a new instance and create a new file
excel = excelManager("path/to/new_file.xlsx")
```

### Workbook Methods

#### Create Workbook

```python
excel.create_workbook("path/to/file.xlsx")
```

Creates a new Excel workbook at the specified path. If no path is provided, it uses the instance's file path.

#### Load Workbook

```python
excel.load_workbook("path/to/file.xlsx")
```

Loads an existing Excel workbook from the specified path. If no path is provided, it uses the instance's file path.

#### Save Workbook

```python
excel.save()
# or
excel.save("path/to/file.xlsx")
```

Saves the workbook to disk. If no path is provided, it uses the instance's file path.

#### Close Workbook

```python
excel.close()
```

Closes the workbook and cleans up resources.

### Sheet Management Methods

#### Count Sheets

```python
sheet_count = excel.count_sheets()
```

Returns the number of sheets in the workbook.

#### Get Sheet Names

```python
sheet_names = excel.get_sheet_names()
```

Returns a list of all sheet names in the workbook.

#### Create Sheet

```python
excel.create_sheet("Sheet1")
```

Creates a new sheet with the specified name. If a sheet with that name already exists, it returns the existing sheet.

#### Get Sheet

```python
sheet = excel.get_sheet("Sheet1")
```

Returns the sheet object with the specified name.

#### Delete Sheet

```python
excel.delete_sheet("Sheet1")
```

Deletes the sheet with the specified name.

### Cell and Range Operations

#### Read Cell

```python
# Using cell reference
value = excel.read_cell("Sheet1", "A1")

# Using row and column numbers
value = excel.read_cell("Sheet1", 1, 1)

# With sheet reference in cell
value = excel.read_cell("Sheet1", "Sheet2!A1")
```

Reads a cell value with formatting preserved. Returns the calculated value if the cell contains a formula.

#### Write Cell

```python
# Using cell reference
excel.write_cell("Sheet1", "A1", "Hello World")

# Using row and column numbers
excel.write_cell("Sheet1", 1, 1, "Hello World")

# With sheet reference in cell
excel.write_cell("Sheet1", "Sheet2!A1", "Hello World")
```

Writes a value to a cell. Can handle different parameter arrangements for flexibility.

#### Read Range

```python
# Using range notation
values = excel.read_range("Sheet1", "A1:C3")

# Using start and end cell references
values = excel.read_range("Sheet1", "A1", "C3")

# Using row and column numbers
values = excel.read_range("Sheet1", 1, 1, 3, 3)
```

Reads a range of cells with formatting preserved. Returns a 2D list of values.

#### Write Range

```python
# Sample data (2D list)
data = [
    ["Name", "Age", "City"],
    ["John", 30, "New York"],
    ["Jane", 25, "Chicago"]
]

# Using cell reference
excel.write_range("Sheet1", "A1", data)

# Using row and column numbers
excel.write_range("Sheet1", 1, 1, data)
```

Writes a 2D list of values to a range of cells starting from the specified cell.

### Special Methods

#### Read Total

```python
# Using cell reference
total = excel.read_total("Sheet1", "A1")

# Using row and column numbers
total = excel.read_total("Sheet1", 1, 1)
```

Finds the last non-empty value in a column (typically a total). Starts from the specified position and traverses down the column until an empty cell is encountered, then returns the last non-empty value.

#### Read Title Total

```python
# Using cell reference
total = excel.read_title_total("Sheet1", "A1", "Revenue")

# Using row and column numbers
total = excel.read_title_total("Sheet1", 1, "Revenue", 1)
```

Finds a column with a matching title, then gets the total value from that column. The title search is case-insensitive. Starts from the specified position and traverses right to find the column with the matching title, then traverses down that column to find the total.

#### Read Items

```python
# Using cell reference
items = excel.read_items("Sheet1", "A1", offset=0)

# Using row and column numbers
items = excel.read_items("Sheet1", 1, 1, offset=2)
```

Reads consecutive non-empty values from a column. Starts from the specified position and collects values until it encounters an empty cell. The `offset` parameter allows excluding a number of items from the end (useful for excluding totals).

#### Read Columns

```python
# Using cell references
columns_data = excel.read_columns("Sheet1", "A1,C1,E1", use_titles=False)

# Using column titles
columns_data = excel.read_columns("Sheet1", "Revenue,Expenses,Profit", use_titles=True, start_row=1)

# Using a list instead of a comma-separated string
columns_data = excel.read_columns("Sheet1", ["Revenue", "Expenses", "Profit"], use_titles=True, start_row=1)
```

Reads multiple columns from a sheet and appends them side by side. This method can work in two modes:

1. **Cell Reference Mode**: Extracts columns starting from the specified cells, reading consecutive non-empty values. For example, "A1,C1,E1" will extract three columns starting from those cells.

2. **Title Mode**: Searches for columns with matching titles, then extracts the data from those columns. For example, "Revenue,Expenses,Profit" will find columns with those titles and extract their data.

The returned data is a 2D list with the first row containing the column headers and subsequent rows containing the data from each column, side by side. If columns have different lengths, shorter columns are padded with empty strings.

## Excel App (excel_app.py)

The Excel App is a Streamlit-based user interface for interacting with the Excel Manager class. It provides a visual way to test and demonstrate the capabilities of the Excel Manager without writing code.

### Running the App

To run the Excel App:

```bash
streamlit run excel_app.py
```

### File Operations

1. **Upload an Existing Excel File**:
   - Use the file uploader in the sidebar to select an existing Excel file
   - The app loads the file and displays a success message

2. **Create a New Excel File**:
   - Enter a file name in the text input field in the sidebar
   - Click "Create New File" to generate a new blank Excel workbook
   - The app creates the file and displays a success message

3. **Reset the App**:
   - Click the "Reset" button in the sidebar to clear the current workbook and start fresh
   - The app returns to its initial state

4. **Download the Modified File**:
   - After making changes, use the "Download Excel file" button at the bottom of the page to save your modified workbook

### Sheet Operations

In the "Sheets" tab, you can:

1. **Count Sheets**:
   - Click "Count Sheets" to display the number of sheets in the workbook

2. **Get Sheet Names**:
   - Click "Get Sheet Names" to display a list of all sheet names

3. **Create a New Sheet**:
   - Enter a name in the "New sheet name" field
   - Click "Create Sheet" to add it to the workbook

### Read Operations

In the "Read" tab, you can:

1. **Read a Cell Value**:
   - Select a sheet from the dropdown
   - Enter a cell reference (e.g., "A1")
   - Click "Read Cell" to display the value

2. **Read a Range of Cells**:
   - Select a sheet from the dropdown
   - Enter a range reference (e.g., "A1:C5")
   - Click "Read Range" to display the values as a table

3. **Find a Total Value**:
   - Select a sheet from the dropdown
   - Enter a starting cell reference (e.g., "A1")
   - Click "Find Total" to display the last non-empty value in that column

4. **Find a Title Total Value**:
   - Select a sheet from the dedicated dropdown
   - Enter a starting cell reference (e.g., "A1") where the header row begins
   - Enter the title text to search for (case-insensitive)
   - Click "Find Title Total" to display the total value from the column with that title

5. **Find Items in a Column**:
   - Select a sheet from the dropdown
   - Enter a starting cell reference (e.g., "A1")
   - Set an offset value (optional)
   - Click "Find Items" to display all consecutive non-empty values

6. **Read Multiple Columns**:
   - Select a sheet from the dropdown
   - Choose between "Cell References" or "Column Titles" input type
   - For Cell References:
     - Enter comma-separated cell references (e.g., "A1,C1,E1")
   - For Column Titles:
     - Enter comma-separated column titles (e.g., "Revenue,Expenses,Profit")
     - Specify the row number where titles are located
   - Click "Get Columns" to display the columns side by side as a table

### Write Operations

In the "Write" tab, you can:

1. **Write to a Cell**:
   - Select a sheet from the dropdown
   - Enter a cell reference (e.g., "A1")
   - Enter a value to write
   - Click "Write Cell" to update the cell

2. **Write to a Range**:
   - Select a sheet from the dropdown
   - Enter a starting cell reference (e.g., "A1")
   - Enter comma-separated values with one row per line in the text area
   - Click "Write Range" to update the cells

### Delete Operations

In the "Delete" tab, you can:

1. **Delete a Sheet**:
   - Select a sheet from the dropdown
   - Click "Delete Sheet" to remove it from the workbook
   - The app prevents deleting the last sheet

## How to Use the Excel Manager in Your Own Projects

To use the Excel Manager in your own Python projects:

1. Import the class:
   ```python
   from excel_manager import excelManager
   ```

2. Create an instance:
   ```python
   excel = excelManager("path/to/file.xlsx")
   ```

3. Use the methods as needed:
   ```python
   # Create a new sheet
   excel.create_sheet("Data")
   
   # Write a header row
   excel.write_range("Data", "A1", [["ID", "Name", "Value"]])
   
   # Add some data
   excel.write_cell("Data", "A2", 1)
   excel.write_cell("Data", "B2", "Product A")
   excel.write_cell("Data", "C2", 100)
   
   # Save the workbook
   excel.save()
   
   # Read the data
   range_data = excel.read_range("Data", "A1:C2")
   print(range_data)
   
   # Find a column by title and get its total
   total = excel.read_title_total("Data", "A1", "Value")
   print(f"Total value: {total}")
   
   # Extract multiple columns by title
   columns = excel.read_columns("Data", "ID,Value", use_titles=True, start_row=1)
   print(columns)
   ```

4. When you're done, close the workbook:
   ```python
   excel.close()
   ```

## Implementation Details

The class maintains two copies of each workbook:
- A data-only version for calculated values
- A formula version for maintaining formulas and formatting

This dual approach ensures that both calculated values and original formulas are accessible when reading and writing.

The class also handles various error cases, such as:
- Missing file paths
- Non-existent files
- Invalid cell references
- Non-existent sheets

All operations are logged to both a file ("excel_manager.log") and the console, making it easier to track what's happening and diagnose issues.