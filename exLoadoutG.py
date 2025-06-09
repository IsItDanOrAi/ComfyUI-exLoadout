import openpyxl
import os
from folder_paths import get_filename_list, get_full_path_or_raise, get_folder_paths

# Hack: string type that is always equal in not equal comparisons
class AnyType(str):
    def __ne__(self, __value: object) -> bool:
        return False

# Our any instance wants to be a wildcard string
ANY = AnyType("*")

class exLoadoutSeg2:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": os.path.join("exLoadoutList.xlsx")}),
                "sheet_name": ("STRING", {"default": "KSAMPLER"}),
                "row_number": ("INT", {"default": 1, "min": 1, "max": 1000, "step": 1}),
                "search_string": ("STRING", {"default": ""}),
            }
        }
    
    # Return types for columns G through L
    RETURN_TYPES = (ANY, ANY, ANY, ANY, ANY, ANY, "STRING")
    RETURN_NAMES = ("Column G", "Column H", "Column I", "Column J", "Column K", "Column L", "Outputs")
    
    # Set all outputs except the summary to be lists
    OUTPUT_IS_LIST = (True, True, True, True, True, True, False)
    
    FUNCTION = "process_excel"
    CATEGORY = "exLoadout"
    DESCRIPTION = ("Reads values from columns G through L for a specified row number in an Excel spreadsheet. "
                   "Can also search for a string in Column A.")
    NAME = "exLoadoutSeg2 (List)"
    
    def process_excel(self, excel_path, sheet_name, row_number, search_string):
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_excel_path = os.path.join(current_dir, excel_path)
        if not os.path.exists(full_excel_path):
            raise FileNotFoundError(f"Excel file not found: {full_excel_path}")
        
        workbook = openpyxl.load_workbook(full_excel_path)
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")
        
        sheet = workbook[sheet_name]
        
        # Determine the actual row to read based on search_string or row_number
        actual_row = row_number
        if search_string:
            found = False
            for row_idx in range(1, sheet.max_row + 1):
                cell_value = sheet.cell(row=row_idx, column=1).value  # Column A
                if cell_value is not None and str(cell_value).strip() == search_string:
                    actual_row = row_idx
                    found = True
                    break
            
            if not found:
                workbook.close()
                raise ValueError(f"Search string '{search_string}' not found in Column A.")
        
        # Check if the actual_row is within the valid range
        if actual_row < 1 or actual_row > sheet.max_row:
            workbook.close()
            raise ValueError(f"Row number {actual_row} is out of range. The sheet has {sheet.max_row} rows.")
        
        # Read columns G through L (7 to 12 in 1-based index)
        row_data = [sheet.cell(row=actual_row, column=col_idx).value 
                    for col_idx in range(7, 13)]  # Columns G to L
        
        # Handle None values by converting to empty string
        row_data = ['' if value is None else value for value in row_data]
        
        workbook.close()
        
        # Create outputs summary with the requested format including % symbols
        # Convert to string for summary, but preserve original type in the lists
        outputs_summary = f"%G: {row_data[0]} %H: {row_data[1]} %I: {row_data[2]} %J: {row_data[3]} %K: {row_data[4]} %L: {row_data[5]} %"
        
        # Return each column as a single-item list for list compatibility
        return ([row_data[0]], [row_data[1]], [row_data[2]], [row_data[3]], [row_data[4]], [row_data[5]], outputs_summary)