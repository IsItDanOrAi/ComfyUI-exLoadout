import openpyxl
import os
import tkinter as tk
from tkinter import ttk

class AnyType(str):
    def __ne__(self, __value: object) -> bool:
        return False

ANY = AnyType("*")

class exLoadoutEditCell:
    @classmethod
    def INPUT_TYPES(cls):
        return {
            "required": {
                "excel_path": ("STRING", {"default": "TESTLIST2.xlsx"}),  # Default Excel filename
                "sheet_name": ("STRING", {"default": "Sheet1"}),  # Default sheet name
                "row_number": ("INT", {"default": 1, "min": 1, "max": 10000}),  # Row selection
                "column_letter": (
                    ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"],
                    {"default": "A"}
                ),  # Column selection limited to A-L
                "new_value": ("STRING", {"default": ""}),  # New value to be inserted
            },
        }

    RETURN_TYPES = (ANY,)
    RETURN_NAMES = ("output_row",)
    OUTPUT_IS_LIST = (True,)
    FUNCTION = "edit_excel_cell"
    CATEGORY = "exLoadout"
    DESCRIPTION = (
        "Edits a specific cell in an Excel spreadsheet and returns the entire row's values "
        "from columns A to L as a comma-separated string inside a list."
    )

    def edit_excel_cell(self, excel_path, sheet_name, row_number, column_letter, new_value):
        # Get the absolute path based on script location
        current_dir = os.path.dirname(os.path.abspath(__file__))
        full_excel_path = os.path.join(current_dir, excel_path)

        if not os.path.exists(full_excel_path):
            raise FileNotFoundError(f"Excel file not found: {full_excel_path}")

        # Load the Excel workbook
        workbook = openpyxl.load_workbook(full_excel_path)
        if sheet_name not in workbook.sheetnames:
            workbook.close()
            raise ValueError(f"Sheet '{sheet_name}' not found in the Excel file")

        sheet = workbook[sheet_name]

        # Convert column letter to index (A = 1, B = 2, etc.)
        column_index = openpyxl.utils.column_index_from_string(column_letter)

        # Ensure the specified row exists
        if row_number < 1 or row_number > sheet.max_row:
            workbook.close()
            raise ValueError(f"Row {row_number} is out of range. The sheet has {sheet.max_row} rows.")

        # Ensure the specified column is within A-L
        if column_index < 1 or column_index > 12:
            workbook.close()
            raise ValueError(f"Column '{column_letter}' is out of the allowed range A-L.")

        # Edit the specified cell
        sheet.cell(row=row_number, column=column_index).value = new_value

        # Save the workbook
        workbook.save(full_excel_path)

        # Retrieve the entire row's values from columns A to L
        row_values = []
        for col_idx in range(1, 13):  # Columns A (1) to L (12)
            cell_value = sheet.cell(row=row_number, column=col_idx).value
            column_letter = openpyxl.utils.get_column_letter(col_idx)
            row_values.append(f"{column_letter}{row_number}: {str(cell_value)}")

        workbook.close()

        # Join values into a single comma-separated string
        output_string = ", ".join(row_values)

        return ([output_string],)  # Output as a list

    def create_edit_button(self):
        # Create a simple Tkinter window with an 'Edit' button
        root = tk.Tk()
        root.title("Excel Editor")

        def on_edit():
            try:
                self.edit_excel_cell(
                    self.excel_path,
                    self.sheet_name,
                    self.row_number,
                    self.column_letter,
                    self.new_value
                )
                tk.messagebox.showinfo("Success", "Cell edited successfully.")
            except Exception as e:
                tk.messagebox.showerror("Error", str(e))

        edit_button = ttk.Button(root, text="Edit", command=on_edit)
        edit_button.pack(pady=20)

        root.mainloop()

NODE_CLASS_MAPPINGS = {"exLoadoutEditCell": exLoadoutEditCell}
NODE_DISPLAY_NAME_MAPPINGS = {"exLoadoutEditCell": "exLoadout Edit Cell"}