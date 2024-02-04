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

        # Make the window full-screen
        self.root.attributes('-fullscreen', True)

        # Menu Bar
        menu_bar = tk.Menu(root)
        root.config(menu=menu_bar)

        # File Menu
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Import Data", command=self.import_data)

        # Create a DataFrame for holding booking data
        self.booking_data = pd.DataFrame(columns=['Booking Date', 'Travel Date', 'Booking Ref', 'Name'])

        # Create a table (Treeview) for displaying data
        columns = list(self.booking_data.columns)
        self.tree = ttk.Treeview(root, columns=columns, show='headings')
        self.tree.tag_configure("style", background="white", foreground="black")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')  # Adjust the width as needed

        # Configure row and column weights for expanding
        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(0, weight=1)

        self.tree.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if file_path:
            try:
                all_data = pd.read_excel(file_path, sheet_name=None)

                # Clear existing data in the DataFrame
                self.booking_data = pd.DataFrame(columns=['Booking Date', 'Travel Date', 'Booking Ref', 'Name'])

                # Iterate through all sheets and append data to the DataFrame
                for sheet_name, data in all_data.items():
                    sheet_data = pd.DataFrame()
                    sheet_data['Booking Date'] = pd.to_datetime(data['Purchase Date (local time)'], format='%d/%m/%Y %H:%M', errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
                    sheet_data['Travel Date'] = pd.to_datetime(data['Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
                    sheet_data['Booking Ref'] = data['Booking Ref #']
                    sheet_data['Name'] = data['Traveler\'s First Name'] + ' ' + data['Traveler\'s Last Name']

                    # Append data to the main DataFrame
                    self.booking_data = pd.concat([self.booking_data, sheet_data], ignore_index=True)

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
