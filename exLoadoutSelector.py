import openpyxl
import os
from folder_paths import get_filename_list, get_full_path_or_raise, get_folder_paths
import comfy.sd

class exLoadoutSelector:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": "exLoadoutList.xlsx"}),  # Just the filename
                "sheet_name": ("STRING", {"default": "MODELS"}),
                "Loadout": (["Default"], ),  # Static default until dynamically updated
            },
        }
    
    RETURN_TYPES = ("STRING",)
    RETURN_NAMES = ("Loadout",)
    FUNCTION = "get_selected_loadout"
    CATEGORY = "exLoadout"
    DESCRIPTION = "Dropdown populated from Column A of an Excel file. Returns the selected Loadout."
    
    @classmethod
    def NODE_NAME(cls):
        """Sets the node name to 'exLoadout Selector' instead of the class name."""
        return "exLoadout Selector"
    
    @staticmethod
    def resolve_full_path(excel_path):
        """Converts just the filename into a full path."""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, excel_path)
    
    @classmethod
    def get_excel_options(cls, excel_path, sheet_name):
        """Reads Column A from the Excel file and returns a list of options."""
        full_excel_path = cls.resolve_full_path(excel_path)
        if not os.path.exists(full_excel_path):
            print(f"Error: Excel file '{full_excel_path}' not found.")
            return ["Default", "ERROR: FILE NOT FOUND"]
        
        try:
            workbook = openpyxl.load_workbook(full_excel_path)
            if sheet_name not in workbook.sheetnames:
                workbook.close()
                print(f"Error: Sheet '{sheet_name}' not found in Excel file.")
                return ["Default", "ERROR: SHEET NOT FOUND"]
            
            sheet = workbook[sheet_name]
            # Convert cell values to strings, strip whitespace, filter out empty cells
            options = ["Default"] + [str(cell.value).strip() for cell in sheet["A"] if cell.value is not None]
            
            workbook.close()
            
            # If no options found, return Default with error
            return options if options else ["Default", "ERROR: NO DATA FOUND"]
        
        except Exception as e:
            print(f"Error reading Excel file: {e}")
            return ["Default", "ERROR: READ FAILED"]
    
    def get_selected_loadout(self, excel_path, sheet_name, Loadout):
        """
        Returns the selected Loadout value from Column A.
        If 'Default' is selected, return 'Default'.
        If an option from the Excel file is selected, return that option.
        """
        # Dynamically generate options
        dynamic_options = self.get_excel_options(excel_path, sheet_name)
        
        # If Loadout is not in the dynamic options, or if an error occurred, return 'Default'
        if Loadout not in dynamic_options or Loadout.startswith("ERROR:"):
            return ("Default",)
        
        # Otherwise, return the selected Loadout
        return (Loadout,)