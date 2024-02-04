import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd

class BookingManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Booking Management System")

        # Set the background color to white
        self.root.configure(background='white')

        # Menu Bar
        menu_bar = tk.Menu(root)
        root.config(menu=menu_bar)

        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Data", command=self.import_data)

        # Create a DataFrame for holding booking data
        self.booking_data = pd.DataFrame(columns=['Header', 'Date', 'Booking Ref', 'Name'])

        # Create a table (Treeview) for displaying data
        columns = list(self.booking_data.columns)
        self.tree = ttk.Treeview(root, columns=columns, show='headings')
        self.tree.tag_configure("style", background="white", foreground="black")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)  # Adjust the width as needed

        self.tree.pack(pady=20)

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if file_path:
            try:
                new_data = pd.read_excel(file_path)

                # Extract and transform data
                self.booking_data['Header'] = new_data['A'].iloc[1:]
                self.booking_data['Date'] = pd.to_datetime(new_data['C'], format='%d/%m/%Y %H:%M', errors='coerce')
                self.booking_data['Booking Ref'] = new_data['D']
                self.booking_data['Name'] = new_data['H'] + ' ' + new_data['I']

                self.update_treeview()
                print("Data Imported and Transformed Successfully!")
            except Exception as e:
                print(f"Error importing and transforming data: {e}")

    def update_treeview(self):
        # Clear existing items in the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert new data into the treeview
        for index, row in self.booking_data.iterrows():
            self.tree.insert('', index, values=row.tolist(), tags=("style",))

if __name__ == "__main__":
    root = tk.Tk()
    app = BookingManagementSystem(root)
    root.mainloop()
