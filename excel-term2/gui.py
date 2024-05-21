import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
import os
from main import delete_product_sheet, add_product, edit_product
from utils import excel_file_path

EXCEL_FILE = 'datas/products.xlsx'

class ProductApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Product Management")
        self.root.configure(background="gray")

        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        # Buttons
        button_frame = tk.Frame(self.root, bg="gray")
        button_frame.grid(row=0, column=0, padx=10, pady=10, sticky=tk.N)

        add_button = tk.Button(button_frame, text="Add", command=self.add_product)
        add_button.pack(fill=tk.X, pady=5)

        update_button = tk.Button(button_frame, text="Update", command=self.update_product)
        update_button.pack(fill=tk.X, pady=5)

        remove_button = tk.Button(button_frame, text="Remove", command=self.delete_product_gui)
        remove_button.pack(fill=tk.X, pady=5)

        # Input Fields
        input_frame = tk.Frame(self.root, bg="gray")
        input_frame.grid(row=0, column=1, padx=10, pady=10, sticky=tk.N)

        tk.Label(input_frame, text="Name:", bg="gray").grid(row=0, column=0, sticky=tk.W)
        self.name_entry = tk.Entry(input_frame)
        self.name_entry.grid(row=0, column=1, pady=5)

        tk.Label(input_frame, text="Price:", bg="gray").grid(row=1, column=0, sticky=tk.W)
        self.price_entry = tk.Entry(input_frame)
        self.price_entry.grid(row=1, column=1, pady=5)

        tk.Label(input_frame, text="Description:", bg="gray").grid(row=2, column=0, sticky=tk.W)
        self.description_entry = tk.Entry(input_frame)
        self.description_entry.grid(row=2, column=1, pady=5)

        tk.Label(input_frame, text="Count:", bg="gray").grid(row=3, column=0, sticky=tk.W)
        self.count_entry = tk.Entry(input_frame)
        self.count_entry.grid(row=3, column=1, pady=5)

        # Treeview
        self.tree = ttk.Treeview(self.root, columns=("Name", "Price", "Description", "Count"), show='headings')
        self.tree.heading("Name", text="Name")
        self.tree.heading("Price", text="Price")
        self.tree.heading("Description", text="Description")
        self.tree.heading("Count", text="Count")
        self.tree.grid(row=0, column=2, padx=10, pady=10, sticky=tk.NSEW)

        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

    def handle_product( self, mode ):
        name = self.name_entry.get()
        description = self.description_entry.get()
        stock = self.count_entry.get()
        price = self.price_entry.get()

        match mode:
            case "add":
                if name and description and stock and price:
                    try:
                        add_product(excel_file_path, name, description, int(stock), int(price))
                        messagebox.showinfo("Success" , "Product operation completed successfully.")
                    except Exception as e:
                        messagebox.showerror("Error" , f"Failed to perform product operation: {e}")
            case "edit":
                try:
                    edit_product(excel_file_path , name , description , stock , price)
                except Exception as e:
                    messagebox.showerror("Error" , f"Failed to perform product operation: {e}")

            case "delete":
                if name:
                    try:
                        delete_product_sheet(excel_file_path , name)
                        messagebox.showinfo("Success" , "Product sheet deleted successfully.")
                    except Exception as e:
                        messagebox.showerror("Error" , f"Failed to delete product sheet: {e}")
                else:
                    messagebox.showerror("Warn" , "Please enter product name.")
            case _:
                messagebox.showerror("Error" , "Invalid Operation mode.")

    def load_data(self):
        if os.path.exists(EXCEL_FILE):
            wb = load_workbook(EXCEL_FILE)
            for sheet in wb.sheetnames:
                if sheet != "Information":
                    ws = wb[sheet]
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        self.tree.insert("", "end", values=row)
        else:
            messagebox.showerror("Error", f"Excel file '{EXCEL_FILE}' not found!")

    def add_product(self):
        self.handle_product(mode = "add")

    def update_product(self):
        self.handle_product(mode = "edit")

    def delete_product_gui(self):
        self.handle_product(mode = "delete")

if __name__ == "__main__":
    root = tk.Tk()
    app = ProductApp(root)
    root.mainloop()
