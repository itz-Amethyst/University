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

        self.current_sheet_index = 0  # To keep track of the current sheet index

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
        self.tree = ttk.Treeview(self.root, show='headings')
        self.tree.grid(row=0, column=2, padx=10, pady=10, sticky=tk.NSEW)

        # Navigation Buttons
        nav_frame = tk.Frame(self.root, bg="gray")
        nav_frame.grid(row=1, column=2, pady=10)

        prev_button = tk.Button(nav_frame, text="Previous", command=self.prev_sheet)
        prev_button.grid(row=0, column=0, padx=5)

        next_button = tk.Button(nav_frame, text="Next", command=self.next_sheet)
        next_button.grid(row=0, column=1, padx=5)

        self.root.grid_columnconfigure(2, weight=1)
        self.root.grid_rowconfigure(0, weight=1)

    def handle_product(self, mode):
        name = self.name_entry.get()
        description = self.description_entry.get()
        stock = self.count_entry.get()
        price = self.price_entry.get()

        try:
            match mode:
                case "add":
                    if name and description and stock and price:
                        success , msg = add_product(EXCEL_FILE , name , description , int(stock) , int(price))
                        if success:
                            messagebox.showinfo("Success" , "Product operation completed successfully.")
                        else:
                            messagebox.showerror("Error" , msg)
                    else:
                        messagebox.showerror("Error" , "Please fill in all fields for adding a product.")
                case "edit":
                    if name:
                        # Tuple
                        success, msg = edit_product(EXCEL_FILE , name , description , stock , price)
                        if success:
                            messagebox.showinfo("Success" , "Product operation completed successfully.")
                        else:
                            messagebox.showerror("Error" , msg)
                case "delete":
                    if name:
                        success, msg = delete_product_sheet(EXCEL_FILE , name)
                        if success:
                            messagebox.showinfo("Success" , "Product sheet deleted successfully.")
                        else:
                            messagebox.showinfo("Fail" , msg)
                    else:
                        messagebox.showerror("Warning" , "Please enter the product name for deleting.")
        except Exception as e:
            messagebox.showerror("Error" , f"Failed to perform product operation: {e}")

        self.load_data()

    def load_data(self):
        if os.path.exists(EXCEL_FILE):
            self.workbook = load_workbook(EXCEL_FILE)
            self.sheets = self.workbook.sheetnames
            self.display_sheet(self.sheets[self.current_sheet_index])
        else:
            messagebox.showerror("Error", f"Excel file '{EXCEL_FILE}' not found!")

    def display_sheet(self, sheet_name):
        for item in self.tree.get_children():
            self.tree.delete(item)

        ws = self.workbook[sheet_name]
        columns = [cell.value for cell in ws[1]]
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for row in ws.iter_rows(min_row=2, values_only=True):
            self.tree.insert("", "end", values=row)

    def next_sheet(self):
        if self.current_sheet_index < len(self.sheets) - 1:
            self.current_sheet_index += 1
            self.display_sheet(self.sheets[self.current_sheet_index])

    def prev_sheet(self):
        if self.current_sheet_index > 0:
            self.current_sheet_index -= 1
            self.display_sheet(self.sheets[self.current_sheet_index])

    def add_product(self):
        self.handle_product(mode="add")

    def update_product(self):
        self.handle_product(mode="edit")

    def delete_product_gui(self):
        self.handle_product(mode="delete")


if __name__ == "__main__":
    root = tk.Tk()
    app = ProductApp(root)
    root.mainloop()
