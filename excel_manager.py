import logging
import os
import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("excel_manager.log"),
        logging.StreamHandler()
    ]
)

class ExcelManager:
    def __init__(self, file_path=None):
        """
        Initialize the ExcelManager with an optional file path.
        If no file path is provided, operations will require a file path.
        """
        self.logger = logging.getLogger(__name__)
        self.file_path = file_path
        self.workbook = None
        
        if file_path and os.path.exists(file_path):
            self.load_workbook(file_path)
            self.logger.info(f"Initialized ExcelManager with existing file: {file_path}")
        elif file_path:
            self.create_workbook(file_path)
            self.logger.info(f"Initialized ExcelManager with new file: {file_path}")
        else:
            self.logger.info("Initialized ExcelManager without a file")
    
    def create_workbook(self, file_path=None):
        """
        Create a new Excel workbook.
        """
        path = file_path or self.file_path
        if not path:
            self.logger.error("No file path provided")
            raise ValueError("File path is required to create a workbook")
        
        self.workbook = Workbook()
        self.file_path = path
        self.save()
        self.logger.info(f"Created new workbook at {path}")
        return self.workbook
    
    def load_workbook(self, file_path=None):
        """
        Load an existing Excel workbook.
        """
        path = file_path or self.file_path
        if not path:
            self.logger.error("No file path provided")
            raise ValueError("File path is required to load a workbook")
        
        if not os.path.exists(path):
            self.logger.error(f"File does not exist: {path}")
            raise FileNotFoundError(f"File does not exist: {path}")
        
        self.workbook = load_workbook(path)
        self.file_path = path
        self.logger.info(f"Loaded workbook from {path}")
        return self.workbook
    
    def save(self, file_path=None):
        """
        Save the workbook to disk.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        path = file_path or self.file_path
        if not path:
            self.logger.error("No file path provided")
            raise ValueError("File path is required to save a workbook")
        
        self.workbook.save(path)
        self.file_path = path
        self.logger.info(f"Saved workbook to {path}")
    
    def close(self):
        """
        Close the workbook.
        """
        if self.workbook:
            self.workbook.close()
            self.workbook = None
            self.logger.info("Closed workbook")
    
    def count_sheets(self):
        """
        Return the number of sheets in the workbook.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        count = len(self.workbook.sheetnames)
        self.logger.info(f"Counted {count} sheets")
        return count
    
    def get_sheet_names(self):
        """
        Return the names of the sheets in the workbook.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        names = self.workbook.sheetnames
        self.logger.info(f"Retrieved sheet names: {names}")
        return names
    
    def create_sheet(self, sheet_name):
        """
        Create a new sheet in the workbook.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        if sheet_name in self.workbook.sheetnames:
            self.logger.warning(f"Sheet {sheet_name} already exists")
            return self.workbook[sheet_name]
        
        sheet = self.workbook.create_sheet(sheet_name)
        self.logger.info(f"Created new sheet: {sheet_name}")
        return sheet
    
    def get_sheet(self, sheet_name):
        """
        Get a sheet by name.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        if sheet_name not in self.workbook.sheetnames:
            self.logger.error(f"Sheet does not exist: {sheet_name}")
            raise ValueError(f"Sheet does not exist: {sheet_name}")
        
        sheet = self.workbook[sheet_name]
        self.logger.info(f"Retrieved sheet: {sheet_name}")
        return sheet
    
    def delete_sheet(self, sheet_name):
        """
        Delete a sheet by name.
        """
        if not self.workbook:
            self.logger.error("No workbook loaded")
            raise ValueError("No workbook loaded")
        
        if sheet_name not in self.workbook.sheetnames:
            self.logger.error(f"Sheet does not exist: {sheet_name}")
            raise ValueError(f"Sheet does not exist: {sheet_name}")
        
        del self.workbook[sheet_name]
        self.logger.info(f"Deleted sheet: {sheet_name}")
    
    def read_cell(self, sheet_name, row, column):
        """
        Read a cell value.
        """
        sheet = self.get_sheet(sheet_name)
        value = sheet.cell(row=row, column=column).value
        self.logger.info(f"Read value '{value}' from cell ({row}, {column}) in sheet {sheet_name}")
        return value
    
    def write_cell(self, sheet_name, row, column, value):
        """
        Write a value to a cell.
        """
        sheet = self.get_sheet(sheet_name)
        sheet.cell(row=row, column=column).value = value
        self.logger.info(f"Wrote value '{value}' to cell ({row}, {column}) in sheet {sheet_name}")
    
    def read_range(self, sheet_name, start_row, start_column, end_row, end_column):
        """
        Read a range of cells.
        """
        sheet = self.get_sheet(sheet_name)
        values = []
        for row in range(start_row, end_row + 1):
            row_values = []
            for col in range(start_column, end_column + 1):
                row_values.append(sheet.cell(row=row, column=col).value)
            values.append(row_values)
        
        self.logger.info(f"Read range from ({start_row}, {start_column}) to ({end_row}, {end_column}) in sheet {sheet_name}")
        return values
    
    def write_range(self, sheet_name, start_row, start_column, values):
        """
        Write values to a range of cells.
        """
        sheet = self.get_sheet(sheet_name)
        for i, row_values in enumerate(values):
            for j, value in enumerate(row_values):
                sheet.cell(row=start_row + i, column=start_column + j).value = value
        
        self.logger.info(f"Wrote values to range starting at ({start_row}, {start_column}) in sheet {sheet_name}")