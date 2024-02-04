import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
import pandas as pd
import json

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
        self.booking_data = pd.DataFrame(columns=['Count', 'Booking Date', 'Travel Date', 'Booking Ref', 'Name', 'Phone No'])

        # Create a table (Treeview) for displaying data
        columns = list(self.booking_data.columns)
        self.tree = ttk.Treeview(root, columns=columns, show='headings')
        self.tree.tag_configure("style", background="white", foreground="black")

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor='center')  # Adjust the width as needed

        for col in self.tree['columns']:
            self.tree.heading(col, text=col, command=lambda c=col: self.on_column_click(c))
            self.tree.column(col, anchor='center')
            col_no_spaces = col.replace(' ', '_')  # Replace spaces with underscores
            if col != 'Count':  # Exclude 'Count' column from B1-Motion and ButtonRelease-1 bindings
                col_index = self.tree['columns'].index(col)
                self.tree.bind('<B1-Motion>', lambda event, c=col, i=col_index: self.on_column_resizing(event, c, i))
                self.tree.bind(f'<ButtonRelease-1>', self.on_column_release)

        # Add a Sizegrip widget for automatic column width adjustment
        self.sizegrip = ttk.Sizegrip(root)

        # Configure row and column weights for expanding
        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # Add a horizontal scrollbar
        x_scrollbar = ttk.Scrollbar(root, orient='horizontal', command=self.tree.xview)
        self.tree.configure(xscrollcommand=x_scrollbar.set)

        # Add a vertical scrollbar
        y_scrollbar = ttk.Scrollbar(root, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=y_scrollbar.set)

        # Place widgets in the grid
        self.tree.pack(fill=tk.BOTH, expand=True)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.sizegrip.pack(side="bottom", fill="both")

        self.load_column_configuration()

    #=========================================COLUMN RESIZING===================================================#

    def on_column_click(self, column):
        # Save the current column configuration before resizing starts
        self.save_column_configuration()

    def on_column_resizing(self, event, column):
        # Update the configuration while the user is actively resizing
        original_width = self.tree.column(column, 'width')
        delta_x = event.x - event.x_root
        new_width = max(original_width + delta_x, 1)  # Ensure the width is not negative

        # Calculate the width change relative to the original width
        delta_width = new_width - original_width

        # Update the width of the current column
        self.tree.column(column, width=new_width)

        # Spread the width change among other columns
        for col in self.tree['columns']:
            if col != column:
                current_width = self.tree.column(col, 'width')
                self.tree.column(col, width=current_width - delta_width)

    def on_column_release(self, event):
        # Save the column configuration after resizing is done
        self.save_column_configuration()

    def save_column_configuration(self):
        # Save the column configuration to a file (e.g., config.json)
        config = {col: {'width': self.tree.column(col, 'width')} for col in self.tree['columns']}
        with open('config.json', 'w') as file:
            json.dump(config, file)

    def load_column_configuration(self):
        try:
            # Load the column configuration from a file (e.g., config.json)
            with open('config.json', 'r') as file:
                config = json.load(file)

            # Apply the saved column widths
            for col, values in config.items():
                self.tree.column(col, width=values['width'], anchor='center')

        except FileNotFoundError:
            # Use default column configuration if the file is not found
            pass
        except Exception as e:
            print(f"Error loading column configuration: {e}")

    #===================================================IMPORT DATA====================================================#

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if file_path:
            try:
                all_data = pd.read_excel(file_path, sheet_name=None)

                # Clear existing data in the DataFrame
                self.booking_data = pd.DataFrame(columns=['Count', 'Booking Date', 'Travel Date', 'Booking Ref', 'Name'])

                count = 0

                # Iterate through all sheets and append data to the DataFrame
                for sheet_name, data in all_data.items():
                    sheet_data = pd.DataFrame()
                    sheet_data['Booking Date'] = pd.to_datetime(data['Purchase Date (local time)'], format='%d/%m/%Y %H:%M', errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
                    sheet_data['Travel Date'] = pd.to_datetime(data['Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
                    sheet_data['Booking Ref'] = data['Booking Ref #']
                    sheet_data['Name'] = data['Traveler\'s First Name'] + ' ' + data['Traveler\'s Last Name']
                    sheet_data['Phone No'] = data['Phone']

                    sheet_data['Count'] = range(count + 1, count + 1 + len(sheet_data))
                    count += len(sheet_data)

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

        self.save_column_configuration()

if __name__ == "__main__":
    root = tk.Tk()
    app = BookingManagementSystem(root)
    root.mainloop()
