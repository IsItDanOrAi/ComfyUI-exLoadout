import openpyxl
import os

class AnyType(str):
    def __ne__(self, __value: object) -> bool:
        return False

ANY = AnyType("*")

class exLoadoutReadColumn:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": "TESTLIST2.xlsx"}),  # Default Excel filename
                "sheet_name": ("STRING", {"default": "Sheet1"}),  # Default sheet name
                "column_letter": ("STRING", {"default": "A"}),  # Column selection by letter
            },
        }

    RETURN_TYPES = (ANY,)
    RETURN_NAMES = ("output_list",)
    OUTPUT_IS_LIST = (True,)
    FUNCTION = "read_excel_column"
    CATEGORY = "exLoadout"
    DESCRIPTION = "Reads all values from a specified column in an Excel spreadsheet and returns them as a comma-separated string inside a list."

    def read_excel_column(self, excel_path, sheet_name, column_letter):
        # Get the absolute path based on script location
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_excel_path = os.path.join(current_dir, excel_path)

        if not os.path.exists(full_excel_path):
            raise FileNotFoundError(f"Excel file not found: {full_excel_path}")

        # Load the Excel workbook
        workbook = openpyxl.load_workbook(full_excel_path, data_only=True)

        if sheet_name not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")

        sheet = workbook[sheet_name]

        # Convert column letter to index (A = 1, B = 2, etc.)
        column_index = openpyxl.utils.column_index_from_string(column_letter)

        # Read all non-empty values in the column, excluding the header
        values = [
            str(sheet.cell(row=row_idx, column=column_index).value)
            for row_idx in range(2, sheet.max_row + 1)  # Start from row 2 to skip header
            if sheet.cell(row=row_idx, column=column_index).value is not None
        ]

        workbook.close()
        
        # Join values into a single comma-separated string
        output_string = ", ".join(values)

        return ([output_string],)  # Output as a list

NODE_CLASS_MAPPINGS = {"exLoadoutReadColumn": exLoadoutReadColumn}
NODE_DISPLAY_NAME_MAPPINGS = {"exLoadoutReadColumn": "exLoadout Read Column"}
