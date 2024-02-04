import tkinter as tk
from tkinter import ttk
import pandas as pd
import json

class SearchFunctions:
    def __init__(self, root, booking_data, tree, revert_filter_enabled, search_menu):
        self.root = root
        self.booking_data = booking_data
        self.tree = tree
        self.revert_filter_enabled = revert_filter_enabled
        self.search_menu = search_menu

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