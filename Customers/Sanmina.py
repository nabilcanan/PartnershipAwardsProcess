import tkinter as tk
from copy import copy
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter


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
                    "1. Select the Award File for Sanmina.\n"
                    "2. Select the BOM file for Sanmina.\n"
                    "3. Save the final merged file.")
    instructions_label = tk.Label(root, text=instructions, font=("Verdana", 16), bg=bg_color, fg=text_color)
    instructions_label.pack(pady=20)

    process_btn = tk.Button(root, text="Process and Merge Files", command=lambda: process_and_merge_files(root),
                            bg=button_color, fg=button_text_color, font=("Verdana", 16))
    process_btn.pack(pady=10)

    root.mainloop()


def process_and_merge_files(parent_window):
    initial_dir = r"C:\Users\nabil\OneDrive\Documentos\WORKFILES\awardsprocess"

    file_path_1 = filedialog.askopenfilename(parent=parent_window, title="Select the first Excel file",
                                             filetypes=[("Excel files", "*.xlsx")], initialdir=initial_dir)
    file_path_2 = filedialog.askopenfilename(parent=parent_window, title="Select the second Excel file",
                                             filetypes=[("Excel files", "*.xlsm")], initialdir=initial_dir)
    if not file_path_1 or not file_path_2:
        messagebox.showwarning("Warning", "You need to select both files!")
        return

    try:
        # Process the first file
        workbook = load_workbook(file_path_1, data_only=False)
        sheet = workbook.active
        process_first_file(sheet)

        # Load the second file and add the 'Working Copy' sheet
        workbook2 = load_workbook(file_path_2, data_only=True)
        if 'Working Copy' in workbook2.sheetnames:
            sheet2 = workbook2['Working Copy']
            new_sheet = workbook.create_sheet('Working Copy')

            # Copying data and style from sheet2 to new_sheet
            for row in sheet2:
                for cell in row:
                    new_cell = new_sheet.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:  # Copy style if present
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = cell.number_format
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)

        else:
            messagebox.showwarning("Warning", "'Working Copy' sheet not found in the second file!")
            return

        # Ask user for save location and file name
        save_file_path = filedialog.asksaveasfilename(parent=parent_window, defaultextension=".xlsx",
                                                      filetypes=[("Excel files", "*.xlsx")],
                                                      title="Save Merged File As", initialdir=initial_dir)
        if save_file_path:
            workbook.save(save_file_path)
            messagebox.showinfo("Success", "File has been processed and saved successfully!")

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
    fill = PatternFill(start_color='FDE9D9', end_color='FDE9D9', fill_type='solid')
    wrap_text = Alignment(wrap_text=True)  # Enable text wrapping

    # List of new headers and their respective formula if applicable
    new_headers = [
        ('Ext Award Value', None),
        ('Award Conf', None),
        ('EAU', None),
        ('Award Price', None),
        ('Conf Cost', None),
        ('Ext Cost', None),
        ('Award MOQ', None),
        ('Cost Comment', None),
        ('New Business', None)
    ]

    for i, (header, formula) in enumerate(new_headers):
        col_letter = get_column_letter(new_column_index + i)
        cell = sheet.cell(row=2, column=new_column_index + i)
        cell.value = header
        cell.fill = fill  # Apply color fill
        cell.alignment = wrap_text  # Apply text wrapping to headers

    # Applying formulas
    awarded_eau_column = get_column_letter(header_column_map['Awarded EAU'])
    award_price_column = get_column_letter(header_column_map['Award Price'])
    moq_column = get_column_letter(header_column_map['Minimum Order Qty'])

    for row in range(3, sheet.max_row + 1):  # Start processing from row 3 as row 2 contains headers
        sheet.cell(row=row, column=new_column_index).value = f'={awarded_eau_column}{row}*{award_price_column}{row}'
        sheet.cell(row=row, column=new_column_index + 2).value = f'={awarded_eau_column}{row}'
        sheet.cell(row=row, column=new_column_index + 3).value = f'={award_price_column}{row}'
        sheet.cell(row=row,
                   column=new_column_index + 5).value = f'=({award_price_column}{row}-{moq_column}{row})/{award_price_column}{row}'
        sheet.cell(row=row, column=new_column_index + 6).value = f'={moq_column}{row}'

    # Apply filters to the entire header row
    sheet.auto_filter.ref = f"A2:{get_column_letter(sheet.max_column)}2"  # Set filter for all headers

