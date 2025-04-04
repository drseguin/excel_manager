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
- Support for A1 notation and row/column indices
- Consistent error handling and logging
- Currency and numeric formatting support

## Excel Manager Class Details

The Excel Manager maintains two copies of each workbook internally - a data-only version for reading calculated values from formulas and a formula version for maintaining the original formulas and formatting when writing. This dual approach ensures that you get the calculated values when reading while preserving the original structure.

### Constructor

```python
def __init__(self, file_path=None)
```

- **Purpose**: Initialize the Excel Manager with an optional file path
- **Parameters**:
  - `file_path` (optional): Path to an Excel file
- **Behavior**:
  - If a valid file path is provided, loads the existing workbook
  - If a non-existent file path is provided, creates a new workbook
  - If no file path is provided, initializes without loading any workbook

### Workbook Methods

#### Create Workbook

```python
def create_workbook(self, file_path=None)
```

- **Purpose**: Create a new Excel workbook
- **Parameters**:
  - `file_path` (optional): Path where the new workbook will be saved
- **Returns**: The created workbook object
- **Behavior**:
  - Creates a new workbook in memory
  - Immediately saves it to the specified path
  - Uses the instance's file path if none is provided
  - Raises an error if no file path is available

#### Load Workbook

```python
def load_workbook(self, file_path=None)
```

- **Purpose**: Load an existing Excel workbook
- **Parameters**:
  - `file_path` (optional): Path to the Excel file to load
- **Returns**: The loaded workbook object
- **Behavior**:
  - Loads the workbook from the specified path
  - Creates two copies: one with calculated values (data_only=True) and one with formulas
  - Uses the instance's file path if none is provided
  - Raises an error if the file doesn't exist or no path is provided

#### Save Workbook

```python
def save(self, file_path=None)
```

- **Purpose**: Save the workbook to disk
- **Parameters**:
  - `file_path` (optional): Path where to save the workbook
- **Behavior**:
  - Saves the formula workbook to the specified path
  - Reloads both workbooks to ensure they stay in sync
  - Uses the instance's file path if none is provided
  - Raises an error if no workbook is loaded or no path is available

#### Close Workbook

```python
def close(self)
```

- **Purpose**: Close the workbook and clean up resources
- **Behavior**:
  - Closes both workbook instances
  - Sets the workbook references to None

### Sheet Management Methods

#### Count Sheets

```python
def count_sheets(self)
```

- **Purpose**: Count the number of sheets in the workbook
- **Returns**: Integer representing the sheet count
- **Behavior**:
  - Returns the count of sheets in the formula workbook
  - Raises an error if no workbook is loaded

#### Get Sheet Names

```python
def get_sheet_names(self)
```

- **Purpose**: Get a list of all sheet names in the workbook
- **Returns**: List of sheet names as strings
- **Behavior**:
  - Returns the sheet names from the formula workbook
  - Raises an error if no workbook is loaded

#### Create Sheet

```python
def create_sheet(self, sheet_name)
```

- **Purpose**: Create a new sheet in the workbook
- **Parameters**:
  - `sheet_name`: Name for the new sheet
- **Returns**: The created sheet object
- **Behavior**:
  - Creates a sheet with the specified name in both workbook copies
  - Returns the sheet from the formula workbook
  - If the sheet already exists, returns the existing sheet
  - Raises an error if no workbook is loaded

#### Get Sheet

```python
def get_sheet(self, sheet_name)
```

- **Purpose**: Get a sheet by name
- **Parameters**:
  - `sheet_name`: Name of the sheet to retrieve
- **Returns**: The sheet object
- **Behavior**:
  - Returns the sheet from the formula workbook
  - Raises an error if the sheet doesn't exist or no workbook is loaded

#### Delete Sheet

```python
def delete_sheet(self, sheet_name)
```

- **Purpose**: Delete a sheet by name
- **Parameters**:
  - `sheet_name`: Name of the sheet to delete
- **Behavior**:
  - Deletes the sheet from both workbook copies
  - Raises an error if the sheet doesn't exist or no workbook is loaded

### Cell and Range Operations

#### Read Cell

```python
def read_cell(self, sheet_name, row_or_cell, column=None)
```

- **Purpose**: Read a cell value with formatting preserved
- **Parameters**:
  - `sheet_name`: Name of the sheet containing the cell
  - `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
  - `column` (optional): Column number (required if row_or_cell is a row number)
- **Returns**: The formatted cell value
- **Behavior**:
  - Gets the calculated value from the data-only workbook
  - Preserves currency formatting if present
  - Formats numbers with commas and two decimal places
  - Supports reading from a different sheet using notation like "Sheet2!A1"
  - Raises an error if the sheet doesn't exist or no workbook is loaded

#### Write Cell

```python
def write_cell(self, sheet_name, row_or_cell, column=None, value=None)
```

- **Purpose**: Write a value to a cell
- **Parameters**:
  - `sheet_name`: Name of the sheet containing the cell
  - `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
  - `column`: Either the value (if row_or_cell is a cell reference) or the column number
  - `value` (optional): The value to write (required if column is a column number)
- **Behavior**:
  - Writes the value to the formula workbook
  - Supports different parameter arrangements for flexibility
  - Raises an error if the sheet doesn't exist or no workbook is loaded

#### Read Range

```python
def read_range(self, sheet_name, start_cell_or_row, start_column=None, end_cell_or_row=None, end_column=None)
```

- **Purpose**: Read a range of cells with formatting preserved
- **Parameters**:
  - `sheet_name`: Name of the sheet containing the range
  - `start_cell_or_row`: Either a range reference (e.g., "A1:C3"), a start cell reference, or a start row number
  - `start_column` (optional): Either an end cell reference or a start column number
  - `end_cell_or_row` (optional): End row number (required if using row/column numbers)
  - `end_column` (optional): End column number (required if using row/column numbers)
- **Returns**: 2D list of formatted cell values
- **Behavior**:
  - Supports multiple ways of specifying the range
  - Gets calculated values from the data-only workbook
  - Preserves currency formatting
  - Formats numbers with commas and two decimal places
  - Raises an error if the sheet doesn't exist, parameters are invalid, or no workbook is loaded

#### Write Range

```python
def write_range(self, sheet_name, start_cell_or_row, start_column_or_values=None, values_or_end_row=None, end_column=None)
```

- **Purpose**: Write values to a range of cells
- **Parameters**:
  - `sheet_name`: Name of the sheet containing the range
  - `start_cell_or_row`: Either a cell reference string (e.g., "A1") or a start row number
  - `start_column_or_values`: Either the values to write (if start_cell_or_row is a cell reference) or a start column number
  - `values_or_end_row` (optional): The values to write (if using row/column numbers)
  - `end_column` (optional): Unused parameter
- **Behavior**:
  - Supports different parameter arrangements for flexibility
  - Writes the 2D list of values starting from the specified cell
  - Raises an error if the sheet doesn't exist, parameters are invalid, or no workbook is loaded

### Special Methods

#### Read Total

```python
def read_total(self, sheet_name, row_or_cell, column=None)
```

- **Purpose**: Find the last non-empty value in a column (typically a total)
- **Parameters**:
  - `sheet_name`: Name of the sheet to read from
  - `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
  - `column` (optional): Column number (required if row_or_cell is a row number)
- **Returns**: The formatted value of the last non-empty cell
- **Behavior**:
  - Starts from the specified position and traverses down the column
  - Returns the last non-empty value before an empty cell is encountered
  - If it reaches the end of the sheet, returns the last value
  - Preserves currency formatting and numeric formatting
  - Returns None if no non-empty cells are found
  - Raises an error if the sheet doesn't exist or no workbook is loaded

#### Read Title Total

```python
def read_title_total(self, sheet_name, row_or_cell, title, column=None)
```

- **Purpose**: Find a column with a matching title, then get the total value from that column
- **Parameters**:
  - `sheet_name`: Name of the sheet to read from
  - `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
  - `title`: The column title to search for (case-insensitive)
  - `column` (optional): Column number (required if row_or_cell is a row number)
- **Returns**: The formatted total value from the column with the matching title
- **Behavior**:
  - Starts from the specified position and traverses right until it finds a cell with text matching the title (case-insensitive)
  - Once the title is found, traverses down that column until an empty cell is found
  - Returns the last non-empty value in that column (typically a total)
  - Preserves currency formatting and numeric formatting
  - Returns None if the title is not found or no total value exists
  - Raises an error if the sheet doesn't exist or no workbook is loaded

#### Read Items

```python
def read_items(self, sheet_name, row_or_cell, column=None, offset=0)
```

- **Purpose**: Read consecutive non-empty values from a column
- **Parameters**:
  - `sheet_name`: Name of the sheet to read from
  - `row_or_cell`: Either a cell reference string (e.g., "A1") or a row number
  - `column` (optional): Column number (required if row_or_cell is a row number)
  - `offset`: Number of items to exclude from the end (default 0)
- **Returns**: List of formatted values
- **Behavior**:
  - Starts from the specified position and collects values until it encounters an empty cell
  - Allows excluding a specified number of items from the end (useful for excluding totals)
  - Preserves currency formatting and numeric formatting
  - Returns an empty list if no non-empty cells are found
  - Raises an error if the sheet doesn't exist or no workbook is loaded

### Helper Methods

#### Parse Cell Reference

```python
def _parse_cell_reference(self, cell_reference, current_sheet_name=None)
```

- **Purpose**: Convert a cell reference string to sheet name, row, and column
- **Parameters**:
  - `cell_reference`: Cell reference string (e.g., "A1" or "Sheet2!B3")
  - `current_sheet_name` (optional): Default sheet name to use if not specified in the reference
- **Returns**: Tuple of (sheet_name, row, column)
- **Behavior**:
  - Handles cell references with or without sheet names
  - Extracts the sheet name if present (e.g., "Sheet2!B3")
  - Converts the column letter to a column number
  - Used internally by other methods

#### Format Numeric Value

```python
def _format_numeric_value(self, value, is_currency=False)
```

- **Purpose**: Format numeric values consistently
- **Parameters**:
  - `value`: The value to format
  - `is_currency`: Whether to add a currency symbol
- **Returns**: Formatted string for numbers, original value for non-numbers
- **Behavior**:
  - Formats numbers with commas and two decimal places
  - Adds a dollar sign for currency values
  - Returns an empty string for None values
  - Returns the original value for non-numeric values
  - Used internally by other methods

## Excel App (excel_app.py)

The Excel App is a Streamlit-based user interface for interacting with the Excel Manager class. It provides a visual way to test and demonstrate the capabilities of the Excel Manager without writing code.

### Main Features

1. **File Operations**:
   - Upload existing Excel files
   - Create new Excel files
   - Download modified Excel files
   - Reset the application state

2. **Sheet Management**:
   - Count the number of sheets
   - List all sheet names
   - Create new sheets
   - Delete existing sheets

3. **Read Operations**:
   - Read individual cell values
   - Read ranges of cells with tabular display
   - Find totals in columns
   - Find totals by column title
   - Extract sequences of items from columns

4. **Write Operations**:
   - Write values to individual cells
   - Write ranges of data using CSV input

### Interface Organization

The app is organized into several tabs for different types of operations:

1. **Sheets Tab**:
   - Buttons for sheet-related operations (count, list, create)
   - Input field for creating new sheets

2. **Read Tab**:
   - Sheet selector dropdown
   - Cell reference input for reading single cells
   - Range reference input for reading ranges
   - Total finder with cell reference input
   - Title total finder with sheet selector, cell reference, and title inputs
   - Item finder with cell reference and offset inputs

3. **Write Tab**:
   - Sheet selector dropdown
   - Cell reference and value inputs for writing to cells
   - Start cell input and CSV text area for writing ranges

4. **Delete Tab**:
   - Sheet selector dropdown for deleting sheets
   - Delete button with safety check

### Usage

To run the Excel App:

```bash
streamlit run excel_app.py
```

#### File Operations

1. **Upload an Excel File**:
   - Use the file uploader in the sidebar to select an existing Excel file
   - The app loads the file and displays a success message

2. **Create a New Excel File**:
   - Enter a file name in the text input field
   - Click "Create New File" to generate a new blank Excel workbook
   - The app creates the file and displays a success message

3. **Reset the App**:
   - Click the "Reset" button to clear the current workbook and start fresh
   - The app returns to its initial state

#### Sheet Operations

1. **Count Sheets**:
   - Click "Count Sheets" to display the number of sheets in the workbook

2. **Get Sheet Names**:
   - Click "Get Sheet Names" to display a list of all sheet names

3. **Create a New Sheet**:
   - Enter a name in the "New sheet name" field
   - Click "Create Sheet" to add it to the workbook

#### Read Operations

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

#### Write Operations

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

#### Delete Operations

1. **Delete a Sheet**:
   - Select a sheet from the dropdown
   - Click "Delete Sheet" to remove it from the workbook
   - The app prevents deleting the last sheet

#### Download the Modified File

After making changes, use the "Download Excel file" button at the bottom of the page to save your modified workbook.

## Implementation Details

The class maintains two copies of each workbook:
- A data-only version for calculated values
- A formula version for maintaining formulas and formatting

This dual approach ensures that both calculated values and original formulas are accessible when reading and writing.

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