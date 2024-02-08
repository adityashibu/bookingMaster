import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox
import pandas as pd
import json
import sqlite3

class BookingManagementSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("Booking Management System")

        # Set up a protocol handler to call save_data_to_db before closing
        root.protocol("WM_DELETE_WINDOW", self.on_close)

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
        file_menu.add_command(label="Export to Excel", command=self.export_to_excel)  # Add Export to Excel option
        file_menu.add_separator()
        file_menu.add_command(label="Close", command=self.on_close)

        # Search Menu
        search_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Search", menu=search_menu)
        search_menu.add_command(label="Filter Data", command=self.show_search_dialog)
        self.revert_filter_enabled = False  # Flag to track whether a filter is applied
        search_menu.add_command(label="Revert Filter", command=self.revert_filter, state=tk.DISABLED)

        self.search_menu = search_menu
        
        # Mail Menu
        mail_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Mail", menu=mail_menu)
        mail_menu.add_command(label="Mail to Customer", command=self.send_mail)

        # Create a DataFrame for holding booking data
        self.booking_data = pd.DataFrame(columns=['Count', 'Booking Date', 'Travel Date', 'Booking Ref', 'Name', 'Phone No', 'Adult', 'Net Price'])

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

        # SQLite Database connection
        self.db_connection = sqlite3.connect('booking_data.db')
        self.create_table_if_not_exists()

        # Load data from the SQLite database
        self.load_data_from_db()

        self.load_column_configuration()
        
    # Command to convert file to excel
    def export_to_excel(self):
        # Check if there is data to export
        if self.booking_data.empty:
            messagebox.showinfo("No Data", "There is no data to export.")
            return

        # Ask user for file save location
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # Write data to Excel file
                self.booking_data.to_excel(file_path, index=False)
                messagebox.showinfo("Export Successful", f"Data exported to {file_path} successfully.")
            except Exception as e:
                messagebox.showerror("Export Error", f"An error occurred while exporting data: {e}")
                
    # Command to handle mail requests\
    def send_mail(self):
        pass

    #======================================DB CONNECTIONS AND CONFIGURATIONS====================================#
        
    def on_close(self):
        # Save data to the SQLite database before closing
        self.save_data_to_db()
        self.root.destroy()
    
    def create_table_if_not_exists(self):
        # Create a table if it doesn't exist in the SQLite database
        query = '''
        CREATE TABLE IF NOT EXISTS bookings (
            ID INTEGER PRIMARY KEY AUTOINCREMENT,
            Booking_Date TEXT,
            Travel_Date TEXT,
            Booking_Ref TEXT,
            Name TEXT,
            Phone_No TEXT,
            Adult INTEGER,
            Net_Price REAL
        );
        '''
        self.db_connection.execute(query)
        self.db_connection.commit()

    def save_data_to_db(self):
        # Save data to the SQLite database
        self.db_connection.execute('DELETE FROM bookings;')  # Clear existing data
        self.booking_data.to_sql('bookings', self.db_connection, index=False, if_exists='replace')
        self.db_connection.commit()

    def load_data_from_db(self):
        # Load data from the SQLite database
        try:
            self.booking_data = pd.read_sql_query('SELECT * FROM bookings;', self.db_connection)
            self.update_treeview()
        except pd.io.sql.DatabaseError:
            # Use default data if the table is empty or not found
            self.booking_data = pd.DataFrame(columns=['Booking_Date', 'Travel_Date', 'Booking_Ref', 'Name', 'Phone_No', 'Adult', 'Net_Price'])
            self.update_treeview()

    def import_data(self):
        file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx;*.xls')])
        if file_path:
            try:
                all_data = pd.read_excel(file_path, sheet_name=None)

                # Clear existing data in the DataFrame and the SQLite database
                self.booking_data = pd.DataFrame(columns=['Booking_Date', 'Travel_Date', 'Booking_Ref', 'Name', 'Phone_No', 'Adult', 'Net_Price'])
                self.save_data_to_db()

                count = 0

                # Iterate through all sheets and append data to the DataFrame and SQLite database
                for sheet_name, data in all_data.items():
                    sheet_data = pd.DataFrame()
                    sheet_data['Booking_Date'] = pd.to_datetime(data['Purchase Date (local time)'], format='%d/%m/%Y %H:%M', errors='coerce').dt.strftime('%d/%m/%Y %H:%M')
                    sheet_data['Travel_Date'] = pd.to_datetime(data['Date'], format='%d/%m/%Y', errors='coerce').dt.strftime('%d/%m/%Y')
                    sheet_data['Booking_Ref'] = data['Booking Ref #']
                    sheet_data['Name'] = data['Traveler\'s First Name'] + ' ' + data['Traveler\'s Last Name']
                    sheet_data['Phone_No'] = data['Phone']
                    sheet_data['Adult'] = pd.to_numeric(data['Adult'], errors='coerce')
                    sheet_data['Net_Price'] = pd.to_numeric(data['Net Price'].str.replace(' AED', ''), errors='coerce')

                    sheet_data['Count'] = range(count + 1, count + 1 + len(sheet_data))
                    count += len(sheet_data)

                    # Append data to the main DataFrame and SQLite database
                    self.booking_data = pd.concat([self.booking_data, sheet_data], ignore_index=True)
                    self.save_data_to_db()

                print("Data Imported and Transformed Successfully!")
            except Exception as e:
                print(f"Error importing and transforming data: {e}")

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
                    sheet_data['Adult'] = pd.to_numeric(data['Adult'], errors='coerce')
                    sheet_data['Net Price'] = pd.to_numeric(data['Net Price'].str.replace(' AED', ''), errors='coerce')

                    sheet_data['Count'] = range(count + 1, count + 1 + len(sheet_data))
                    count += len(sheet_data)

                    # Append data to the main DataFrame
                    self.booking_data = pd.concat([self.booking_data, sheet_data], ignore_index=True)
                    

                self.update_treeview()
                print("Data Imported and Transformed Successfully!")
            except Exception as e:
                print(f"Error importing and transforming data: {e}")

    #================================================SEACRH DATA====================================================#
    def show_search_dialog(self):
        # Create a search dialog
        search_dialog = tk.Toplevel(self.root)
        search_dialog.title("Search Options")

        # Create a label and dropdown menu for selecting search criteria
        tk.Label(search_dialog, text="Select Search Criteria:").grid(row=0, column=0, padx=10, pady=10)
        search_criteria_var = tk.StringVar()
        search_criteria_var.set("Booking Ref")  # Set default value
        search_criteria_menu = ttk.Combobox(search_dialog, textvariable=search_criteria_var, values=['Booking Ref', 'Customer Name', 'No of Adults'])
        search_criteria_menu.grid(row=0, column=1, padx=10, pady=10)

        # Create an entry widget for entering search value
        tk.Label(search_dialog, text="Enter Search Value:").grid(row=1, column=0, padx=10, pady=10)
        search_value_var = tk.StringVar()
        search_value_entry = tk.Entry(search_dialog, textvariable=search_value_var)
        search_value_entry.grid(row=1, column=1, padx=10, pady=10)

        # Create a button to apply the search
        search_button = tk.Button(search_dialog, text="Search", command=lambda: self.apply_search(search_criteria_var.get(), search_value_var.get(), search_dialog))
        search_button.grid(row=2, column=0, columnspan=2, pady=10)

        # Enable or disable "Revert Filter" based on the filter status
        self.search_menu.entryconfig("Revert Filter", state=tk.NORMAL if self.revert_filter_enabled else tk.DISABLED)

    def apply_search(self, criteria, value, search_dialog):
        # Apply the search and update the treeview
        if criteria == 'Booking Ref':
            result = self.booking_data[self.booking_data['Booking Ref'].astype(str).str.contains(value, case=False)]
        elif criteria == 'Customer Name':
            result = self.booking_data[self.booking_data['Name'].astype(str).str.contains(value, case=False)]
        elif criteria == 'No of Adults':
            result = self.booking_data[self.booking_data['Adult'] == int(value)]

        # Update the treeview with the search result
        self.update_treeview(data=result)

        # Set the flag to indicate that a filter is applied
        self.revert_filter_enabled = True

        # Enable "Revert Filter" based on the filter status
        self.search_menu.entryconfig("Revert Filter", state=tk.NORMAL)

        # Destroy the search dialog
        search_dialog.destroy()

    def revert_filter(self):
        # Revert the filter and update the treeview
        self.update_treeview(data=self.booking_data)

        # Reset the flag to indicate that no filter is applied
        self.revert_filter_enabled = False

        # Disable "Revert Filter" since no filter is applied
        self.search_menu.entryconfig("Revert Filter", state=tk.DISABLED)

    def update_treeview(self, data=None):
        # Clear existing items in the treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Insert new data into the treeview
        if data is None:
            data = self.booking_data

        for index, row in data.iterrows():
            values = row.tolist()
            values[-1] = int(values[-1])  # Convert 'Adult' column to integer without decimal points
            values[-2] = str(values[-2]) if pd.notna(values[-2]) else ""  # Convert 'Phone No' column to string
            self.tree.insert('', index, values=row.tolist(), tags=("style",))

        self.save_column_configuration()

if __name__ == "__main__":
    root = tk.Tk()
    app = BookingManagementSystem(root)
    root.mainloop()
