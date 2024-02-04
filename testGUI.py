import tkinter as tk
from tkinter import filedialog
import pandas as pd
from pandastable import Table, TableModel  # Install with: pip install pandastable

class ExcelLikeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel-Like App")

        # Create menu bar
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Excel", command=self.import_excel)

        # Create a PandasTable
        self.table = Table(root, showtoolbar=True, showstatusbar=True)
        self.table.grid(sticky='news')

        # Allow the table to expand with the window
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

    def import_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if file_path:
            try:
                # Read Excel file into Pandas DataFrame
                df = pd.read_excel(file_path)
                
                # Load the DataFrame into the PandasTable
                self.table.model.df = df

            except Exception as e:
                print(f"Error loading Excel file: {e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelLikeApp(root)
    root.mainloop()
