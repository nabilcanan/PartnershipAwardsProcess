import tkinter as tk
from tkinter import filedialog, messagebox

import openpyxl
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment


def press_action():
    print("Hello partnership")


def sanmina_logic():
    root = tk.Tk()
    root.title("Sanmina Module")
    root.geometry("800x500")

    # Color scheme
    bg_color = '#3b5998'
    text_color = '#ffffff'
    button_color = '#8b9dc3'
    button_text_color = '#ffffff'

    root.configure(bg=bg_color)

    label = tk.Label(root, text="Welcome Partnership Member", font=("Verdana", 24), bg=bg_color, fg=text_color)
    label.pack(pady=20)

    instructions = ("Instructions:\n"
                    "1. Select the first Excel file.\n"
                    "2. Select the second Excel file.\n"
                    "3. Click the process button to process and merge the files.\n"
                    "4. Save the final merged file.")
    instructions_label = tk.Label(root, text=instructions, font=("Verdana", 16), bg=bg_color, fg=text_color)
    instructions_label.pack(pady=20)

    process_btn = tk.Button(root, text="Process and Merge Files", command=lambda: process_and_merge_files(root),
                            bg=button_color, fg=button_text_color, font=("Verdana", 16))
    process_btn.pack(pady=10)

    root.mainloop()


def process_and_merge_files(parent_window):
    file_path_1 = filedialog.askopenfilename(parent=parent_window, title="Select the first Excel file",
                                             filetypes=[("Excel files", "*.xlsx")])
    file_path_2 = filedialog.askopenfilename(parent=parent_window, title="Select the second Excel file",
                                             filetypes=[("Excel files", "*.xlsm")])
    if not file_path_1 or not file_path_2:
        messagebox.showwarning("Warning", "You need to select both files!")
        return

    try:
        # Process the first file
        workbook = load_workbook(file_path_1, data_only=False)
        sheet = workbook.active
        process_first_file(sheet)
        # Convert processed OpenPyXL worksheet to a Pandas DataFrame
        df_first = pd.DataFrame(sheet.values)
        df_first.columns = df_first.iloc[1]  # Set the second row as header
        df_first = df_first.drop([0, 1])  # Drop the first two rows which are now header and old header

        # Load the second file
        df_second = pd.read_excel(file_path_2, header=2, sheet_name='Working Copy')
        df_second.columns = [str(col).strip() for col in df_second.columns]

        # Merge data
        columns_to_merge = [
            'PSoft Part', 'Quoted Mfg', 'Quoted Part', 'Part Class', 'Last Ship Resale', 'Last Ship Date',
            'Last Ship GP', '12 Mo CPN Sales', 'Backlog Resale', 'Cust Backlog Value', 'Sager Stock', 'Stock Type',
            'On POs', 'Sager Min', 'Sager Mult', 'Factory LT', 'Avg Cost', 'Vol1 Cost', 'Vol2 Cost', 'Best Book',
            'Best Contract', 'Sager NCNR', 'Last PO Price', 'SND Cost', 'SND Quote', 'SND Exp Date', 'SND Cust ID',
            'VPC Cost', 'VPC Quote', 'VPC Exp Date', 'VPC MOQ', 'VPC TYPE', 'TIR MOQ', 'VPC Cust ID', 'Design Reg #',
            'Reg #', 'Last Ship CPN', 'Last Ship Cust ID', 'Backlog CPN', 'Backlog Entry', 'Backlog Cust ID'
        ]
        merged_data = pd.merge(df_first, df_second[columns_to_merge + ['MPN']], on='MPN', how='left')

        # Save the merged data to a new Excel file
        save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_file_path:
            merged_data.to_excel(save_file_path, index=False)
            messagebox.showinfo("Success", "Files have been processed and merged successfully!")
        else:
            messagebox.showwarning("Cancelled", "File save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", "An error occurred during processing: " + str(e))


def process_first_file(sheet):
    # Initialize dictionary to map column headers to their respective columns
    header_column_map = {}
    for column in range(1, sheet.max_column + 1):
        header_value = sheet.cell(row=2, column=column).value  # Headers are in row 2
        if header_value:
            header_column_map[header_value] = column

    # Check for required headers and map them to column letters
    required_headers = ['Awarded EAU', 'Award Price', 'Minimum Order Qty']
    missing_headers = [header for header in required_headers if header not in header_column_map]
    if missing_headers:
        messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_headers)}")
        return None  # Exit if there are missing headers

    # Determine columns for new headers and formulas
    new_column_index = sheet.max_column + 1
    ext_award_value_column = get_column_letter(new_column_index)
    sheet.cell(row=2, column=new_column_index).value = 'Ext Award Value'

    award_conf_column = get_column_letter(new_column_index + 1)
    sheet.cell(row=2, column=new_column_index + 1).value = 'Award Conf'

    eau_column = get_column_letter(new_column_index + 2)
    sheet.cell(row=2, column=new_column_index + 2).value = 'EAU'

    award_price_column_ba = get_column_letter(new_column_index + 3)
    sheet.cell(row=2, column=new_column_index + 3).value = 'Award Price'

    conf_cost_column = get_column_letter(new_column_index + 4)
    sheet.cell(row=2, column=new_column_index + 4).value = 'Conf Cost'

    ext_cost_column = get_column_letter(new_column_index + 5)
    sheet.cell(row=2, column=new_column_index + 5).value = 'Ext Cost'

    award_moq_column = get_column_letter(new_column_index + 6)
    sheet.cell(row=2, column=new_column_index + 6).value = 'Award MOQ'

    comment_cost_column = get_column_letter(new_column_index + 7)
    sheet.cell(row=2, column=new_column_index + 7).value = 'Cost Comment'

    new_business_column = get_column_letter(new_column_index + 8)
    sheet.cell(row=2, column=new_column_index + 8).value = 'New Business'

    # Applying formulas
    awarded_eau_column = get_column_letter(header_column_map['Awarded EAU'])
    award_price_column = get_column_letter(header_column_map['Award Price'])
    moq_column = get_column_letter(header_column_map['Minimum Order Qty'])

    for row in range(3, sheet.max_row + 1):  # Start processing from row 3 as row 2 contains headers
        sheet.cell(row=row, column=new_column_index).value = f'={awarded_eau_column}{row}*{award_price_column}{row}'
        sheet.cell(row=row, column=new_column_index + 2).value = f'={awarded_eau_column}{row}'
        sheet.cell(row=row, column=new_column_index + 3).value = f'={award_price_column}{row}'
        sheet.cell(row=row, column=new_column_index + 5).value = f'=({award_price_column_ba}{row}-{conf_cost_column}{row})/{award_price_column_ba}{row}'
        sheet.cell(row=row, column=new_column_index + 6).value = f'={moq_column}{row}'


