import os
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Protection, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime
from utils import data_folder, excel_file_path


def create_excel_file():
    """Create an Excel file in the 'datas' folder if it doesn't exist."""
    if not os.path.exists(data_folder):
        os.makedirs(data_folder)

    if not os.path.exists(excel_file_path):
        wb = Workbook()
        ws = wb.active
        ws.title = "Information"
        ws.append(["Address", "Established Year", "City", "Country"])

        # Apply header styles and data validation for the "Information" sheet
        apply_header_styles(ws)
        apply_data_validation(ws, "Information")

        wb.save(excel_file_path)
        print(f"Created new Excel file: {excel_file_path}")
    else:
        print(f"Excel file '{excel_file_path}' already exists.")

def apply_header_styles(ws):
    """Apply styles to the header row."""
    header_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.protection = Protection(locked=True)
        cell.fill = header_fill

    # ws.protection.sheet = True

def apply_data_validation(ws, sheet_type):
    """Apply data validation to the sheet based on its type."""
    if sheet_type == "Information":
        dv_int = DataValidation(type="whole", operator="between", formula1=0, formula2=9999, showErrorMessage=True)
        dv_int.error = "Please enter a valid year."
        dv_int.errorTitle = "Invalid Year"
        ws.add_data_validation(dv_int)
        dv_int.add(f"B2:B1048576")

        dv_text = DataValidation(type="textLength", operator="lessThan", formula1="255", showErrorMessage=True)
        dv_text.error = "Please enter a valid text."
        dv_text.errorTitle = "Invalid Text"
        ws.add_data_validation(dv_text)
        dv_text.add(f"A2:A100")  # Address
        dv_text.add(f"C2:C100")  # City
        dv_text.add(f"D2:D100")  # Country

    elif sheet_type == "Product":
        dv_text = DataValidation(type="textLength", operator="lessThan", formula1="255", showErrorMessage=True)
        dv_text.error = "Please enter a valid text."
        dv_text.errorTitle = "Invalid Text"
        ws.add_data_validation(dv_text)
        dv_text.add(f"B2:B1000")  # Name
        dv_text.add(f"C2:C1000")  # Description

        dv_int = DataValidation(type="whole", operator="greaterThan", formula1=0, showErrorMessage=True)
        dv_int.error = "Please enter a valid integer."
        dv_int.errorTitle = "Invalid Integer"
        ws.add_data_validation(dv_int)
        dv_int.add(f"D2:D1000")  # Stock
        dv_int.add(f"E2:E1000")  # price

        dv_date = DataValidation(type="date", formula1="1900-01-01", showErrorMessage=True)
        dv_date.error = "Please enter a valid date."
        dv_date.errorTitle = "Invalid Date"
        ws.add_data_validation(dv_date)
        dv_date.add(f"A2:A1000")


def prepare_workbook(excel_file):
    """Load the workbook or create if it doesn't exist."""
    if not os.path.exists(excel_file):
        create_excel_file()
    return load_workbook(excel_file)

def save_changes(wb, excel_file):
    try:
        wb.save(excel_file)
        print('Saved Successfully')
    except Exception as e:
        print(f"Failed to save the file: {e}")
        print("There is a chance that this file is in use.")

def add_product(file_name, name, description, stock, price):
    """Add a new product with the current date."""
    sheet_name = name
    wb = prepare_workbook(file_name)
    if sheet_name in wb.sheetnames:
        print(f"Product sheet '{sheet_name}' already exists.")
        return
    ws = wb.create_sheet(title=sheet_name, index=0)  # Add the sheet as the first sheet
    ws.append(["Transaction Date", "Name", "Description", "Stock", "Price"])

    # Apply header styles and data validation for the product sheet
    apply_header_styles(ws)
    apply_data_validation(ws, "Product")

    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), name, description, stock, price])
    save_changes(wb, file_name)
    print(f"Product sheet '{sheet_name}' added successfully.")

def edit_product(file_name, name=None, description=None, stock=None, price=None):
    """Edit an existing product by adding a new record with updated data."""
    sheet_name = name
    if not os.path.exists(file_name):
        print(f"Excel file '{file_name}' does not exist.")
        return
    wb = prepare_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"Product sheet '{sheet_name}' does not exist.")
        return
    if sheet_name not in wb.sheetnames:
        print(f"Product sheet '{sheet_name}' does not exist.")
        return
    ws = wb[sheet_name]
    
    # Retrieve the last row's values
    last_row = ws.max_row
    last_record = {ws.cell(row=1, column=col).value: ws.cell(row=last_row, column=col).value for col in range(1, ws.max_column + 1)}
    
    # Prepare the new record with existing values and update with provided arguments
    new_record = {
        "Transaction Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Name": name if name is not None else last_record["Name"],
        "Description": description if description is not None else last_record["Description"],
        "Stock": stock if stock is not None else last_record["Stock"],
        "Price": price if price is not None else last_record["Price"]
    }
    
    ws.append([new_record["Transaction Date"], new_record["Name"], new_record["Description"], new_record["Stock"], new_record['Price']])
    if name and name != last_record['Name']:
        ws.title = name
    save_changes(wb, file_name)
    print(f"Product sheet '{sheet_name}' updated successfully.")

def delete_product_sheet(file_name, sheet_name):
    """Delete a product sheet."""
    wb = prepare_workbook(file_name)
    if sheet_name not in wb.sheetnames:
        print(f"Product sheet '{sheet_name}' does not exist.")
        return
    del wb[sheet_name]
    save_changes(wb, file_name)
    print(f"Product sheet '{sheet_name}' deleted successfully.")

# Example usage
create_excel_file()
add_product(excel_file_path,  "ProductA Name", "Initial stock", 100, 2000)
edit_product(excel_file_path, "ProductA", description="Stock reduced", stock=80)
edit_product(excel_file_path, name="Updated ProductA Name", stock=60)
edit_product(excel_file_path, "ProductA", description="Final update")
add_product(excel_file_path,"ProductMilad Name", "Initial stock", 100, 200)
delete_product_sheet(excel_file_path, "ProductA")
