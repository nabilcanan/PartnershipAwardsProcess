import os
import tkinter as tk
from tkinter import messagebox, filedialog
from tkinterdnd2 import DND_FILES, TkinterDnD
import pandas as pd
from openpyxl import load_workbook


# Creation logic with drag-and-drop functionality

def creation_logic_dnd():
    # Create a new Toplevel window for Benchmark
    creation_window = tk.Toplevel()
    creation_window.title("Creation Module")
    creation_window.geometry("800x500")

    # Set up TkinterDnD functionality
    TkinterDnD.enable_drop(creation_window)

    # Label for drag and drop
    drop_label = tk.Label(creation_window, text="Drag and drop files here", font=("Verdana", 20), relief="solid",
                          height=5)
    drop_label.pack(pady=20, padx=20, expand=True, fill="both")

    # Function to handle file drop
    def handle_file_drop(event):
        files = event.data.split()
        process_files(files)

    # Bind the drop event
    drop_label.drop_target_register(DND_FILES)
    drop_label.dnd_bind('<<Drop>>', handle_file_drop)

    # Function to process dropped files
    def process_files(files):
        file_dict = {'cpn': None, 'contract': None, 'backlog': None, 'sales_history': None}

        for file in files:
            filename = os.path.basename(file).lower()
            if 'cpn' in filename:
                file_dict['cpn'] = file
            elif 'contract' in filename:
                file_dict['contract'] = file
            elif 'backlog' in filename:
                file_dict['backlog'] = file
            elif 'sales history' in filename:
                file_dict['sales_history'] = file

        # Check if all files are provided
        if all(file_dict.values()):
            df_first = pd.read_excel(file_dict['cpn'])
            if process_second_file(df_first, file_dict['contract']) and \
                    process_third_file(df_first, file_dict['backlog']) and \
                    process_fourth_file(df_first, file_dict['sales_history']):
                save_file(df_first)
        else:
            messagebox.showerror("Error", "All required files not provided.")

    def process_second_file(df_first, file_path):
        # Load the file into a DataFrame without setting a header initially
        df_second_raw = pd.read_excel(file_path, header=None)

        # Find the row index where the desired header ('IPN') is located
        header_row_idx = None
        for idx, row in df_second_raw.iterrows():
            if 'IPN' in row.values:
                header_row_idx = idx
                break

        if header_row_idx is not None and isinstance(header_row_idx, int):
            # Re-read the file with the correct header row
            df_second = pd.read_excel(file_path, header=int(header_row_idx))
            # Check if 'IPN' column is found
            if 'IPN' in df_second.columns:
                df_first['On Contract'] = df_first['CPN'].isin(df_second['IPN']).map({True: 'Y', False: ''})
                return True
            else:
                messagebox.showerror("Error", "IPN column not found in the second file.")
                return False
        else:
            messagebox.showerror("Error", "Header row with 'IPN' not found in the file.")
            return False

    def process_third_file(df_first, file_path):
        # Load the third file into a DataFrame
        df_third = pd.read_excel(file_path)
        # Check if the required columns are in the third file
        if 'Backlog CPN' in df_third.columns and 'Scheduled Arrival' in df_third.columns:
            # Group data by 'Backlog CPN' and get the latest 'Scheduled Arrival' for each CPN
            latest_arrivals = df_third.groupby('Backlog CPN')['Scheduled Arrival'].max()
            # Map the latest arrival dates to the corresponding CPNs in the first DataFrame
            df_first['Last Ship Date'] = df_first['CPN'].map(latest_arrivals)
            return True
        else:
            # Display an error message if the required columns are not found
            messagebox.showerror("Error", "Required columns not found in the third file.")
            return False

    # Process fourth file (Sales History File)
    def process_fourth_file(df_first, file_path):
        # Load the fourth file into a DataFrame
        df_fourth = pd.read_excel(file_path)
        # Check if the required columns are in the fourth file
        if 'Last Ship CPN' in df_fourth.columns and 'Last Ship Date' in df_fourth.columns:
            # Group data by 'Last Ship CPN' and get the latest 'Last Ship Date' for each CPN
            latest_ship_dates = df_fourth.groupby('Last Ship CPN')['Last Ship Date'].max()
            # Update 'Last Ship Date' in the first DataFrame with the latest ship date where it is missing
            df_first['Last Ship Date'] = df_first.apply(
                lambda row: row['Last Ship Date'] if pd.notna(row['Last Ship Date']) else latest_ship_dates.get(
                    row['CPN']), axis=1)
            return True
        else:
            # Display an error message if the required columns are not found
            messagebox.showerror("Error", "Required columns not found in the fourth file.")
            return False

    # Function to save the processed file
    def save_file(df_first):
        save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_file_path:
            # Save DataFrame to an Excel file
            df_first.to_excel(save_file_path, index=False)

            # Open the saved Excel file and adjust column widths
            workbook = load_workbook(save_file_path)
            worksheet = workbook.active

            # Auto-adjust the width of columns based on the longest entry in each column
            for col in worksheet.columns:
                max_length = 0
                column = col[0].column_letter  # Get the column letter

                for cell in col:
                    try:
                        # Check the length of the cell value
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = (max_length + 2)  # Add 2 for a little extra space
                worksheet.column_dimensions[column].width = adjusted_width

            # Save the changes
            workbook.save(save_file_path)
            messagebox.showinfo("Success", "Files processed, formatted, and saved successfully!")
        else:
            messagebox.showwarning("Cancelled", "File save cancelled.")



