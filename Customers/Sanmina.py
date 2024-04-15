import tkinter as tk
from tkinter import filedialog, messagebox
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

    # Modern color scheme
    bg_color = '#3b5998'  # A shade of blue
    text_color = '#ffffff'  # White for readability
    button_color = '#8b9dc3'  # Lighter shade of blue for the button
    button_text_color = '#ffffff'  # White text on the button

    root.configure(bg=bg_color)

    label = tk.Label(root, text="Welcome Partnership Member", font=("Verdana", 24), bg=bg_color, fg=text_color)
    label.pack(pady=20)

    instructions = ("Instructions:\n"
                    "1. Select the file containing the CPNs.\n"
                    "2. Choose the Latest Contract File for Sanmina.\n"
                    "3. Select the Sanmina Backlog File.\n"
                    "4. Pick the Sanmina Sales History File.\n"
                    "5. Finally Select your Sanmina Agility File\n"
                    "6. Remember to SAVE your final file upon completion.")
    instructions_label = tk.Label(root, text=instructions, font=("Verdana", 20), bg=bg_color, fg=text_color)
    instructions_label.pack(pady=20)

    open_file_btn = tk.Button(root, text="Select Excel Files", command=lambda: open_and_process_files(root),
                              bg=button_color, fg=button_text_color, font=("Verdana", 16))
    open_file_btn.pack(pady=10)

    root.mainloop()


def open_and_process_files(parent_window):
    file_path_1 = filedialog.askopenfilename(parent=parent_window, title="Select the first Excel file",
                                             filetypes=[("Excel files", "*.xlsx")])
    file_path_2 = filedialog.askopenfilename(parent=parent_window, title="Select the second Excel file",
                                             filetypes=[("Excel files", "*.xlsm")])

    if not file_path_1 or not file_path_2:
        messagebox.showwarning("Warning", "You need to select both files!")
        return

    try:
        workbook, sheet = process_first_file(file_path_1)
        process_second_file(workbook, sheet, file_path_2)
    except Exception as e:
        messagebox.showerror("Error", "An error occurred during processing: " + str(e))



def process_first_file(file_path):
    workbook = load_workbook(file_path, data_only=False)
    sheet = workbook.active

    # Initialize dictionary to map column headers to their respective columns
    header_column_map = {}
    for column in range(1, sheet.max_column + 1):
        header_value = sheet.cell(row=2, column=column).value  # Check headers in row 2
        if header_value:
            header_column_map[header_value] = column

    # Check for required headers and map them to column letters
    required_headers = ['Awarded EAU', 'Award Price']
    if not all(header in header_column_map for header in required_headers):
        missing = [header for header in required_headers if header not in header_column_map]
        messagebox.showerror("Error", f"Missing required columns: {', '.join(missing)}")
        return

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
    sheet.cell(row=2, column=new_column_index + 4).value = 'Conf Cost'  # New blank column 'Conf Cost'

    ext_cost_column = get_column_letter(new_column_index + 5)
    sheet.cell(row=2, column=new_column_index + 5).value = 'Ext Cost'

    award_moq_column = get_column_letter(new_column_index + 6)
    sheet.cell(row=2, column=new_column_index + 6).value = 'Award MOQ'

    comment_cost_column = get_column_letter(new_column_index + 7)
    sheet.cell(row=2, column=new_column_index + 7).value = 'Cost Comment'  # New blank column 'Cost Comment'

    new_business_column = get_column_letter(new_column_index + 8)
    sheet.cell(row=2, column=new_column_index + 8).value = 'New Business'  # New blank column 'Conf Cost'

    # Applying formulas
    awarded_eau_column = get_column_letter(header_column_map['Awarded EAU'])
    award_price_column = get_column_letter(header_column_map['Award Price'])
    moq_column = get_column_letter(header_column_map['Minimum Order Qty'])

    for row in range(3, sheet.max_row + 1):
        sheet[ext_award_value_column + str(row)].value = f'={awarded_eau_column}{row}*{award_price_column}{row}'
        sheet[eau_column + str(row)].value = f'={awarded_eau_column}{row}'  # Copy 'Awarded EAU' into 'EAU'
        sheet[award_price_column_ba + str(
            row)].value = f'={award_price_column}{row}'  # Copy 'Award Price' into new 'Award Price'
        sheet[ext_cost_column + str(
            row)].value = f'=({award_price_column_ba}{row}-{conf_cost_column}{row})/{award_price_column_ba}{row}'
        sheet[award_moq_column + str(row)].value = f'={moq_column}{row}'

    # Adding dynamic SUBTOTAL formulas in cells AX1 and BC1
    subtotal_column_ax = 'AX'
    subtotal_column_bc = 'BC'
    sheet[
        subtotal_column_ax + '1'].value = f'=SUBTOTAL(9, {subtotal_column_ax}3:{subtotal_column_ax}{sheet.max_row})'
    sheet[
        subtotal_column_bc + '1'].value = f'=SUBTOTAL(9, {subtotal_column_bc}3:{subtotal_column_bc}{sheet.max_row})'

    # Adding the formula in cell BD1
    formula_column_bd = 'BD'
    sheet[formula_column_bd + '1'].value = f'=({subtotal_column_ax}1-{subtotal_column_bc}1)/{subtotal_column_ax}1'

    # Save the workbook
    final_save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                        filetypes=[("Excel files", "*.xlsx")])
    if final_save_file_path:
        workbook.save(final_save_file_path)
        messagebox.showinfo("Success", "File processed and saved successfully!")
    else:
        messagebox.showwarning("Cancelled", "Final file save cancelled.")

        return workbook, sheet


def process_second_file(first_file_path, second_file_path):
    try:
        # Load the data from both Excel files
        df_first = pd.read_excel(first_file_path)
        df_second = pd.read_excel(second_file_path)

        # Specify columns to merge from the second file
        columns_to_merge = [
            'PSoft Part', 'Quoted Mfg', 'Quoted Part', 'Part Class', 'Last Ship Resale', 'Last Ship Date',
            'Last Ship GP', '12 Mo CPN Sales', 'Backlog Resale', 'Cust Backlog Value', 'Sager Stock', 'Stock Type',
            'On POs', 'Sager Min', 'Sager Mult', 'Factory LT', 'Avg Cost', 'Vol1 Cost', 'Vol2 Cost', 'Best Book',
            'Best Contract', 'Sager NCNR', 'Last PO Price', 'SND Cost', 'SND Quote', 'SND Exp Date', 'SND Cust ID',
            'VPC Cost', 'VPC Quote', 'VPC Exp Date', 'VPC MOQ', 'VPC TYPE', 'TIR MOQ', 'VPC Cust ID', 'Design Reg #',
            'Reg #', 'Last Ship CPN', 'Last Ship Cust ID', 'Backlog CPN', 'Backlog Entry', 'Backlog Cust ID'
        ]

        # Check if all required columns are in the second file
        missing_columns = [col for col in columns_to_merge if col not in df_second.columns]
        if missing_columns:
            messagebox.showerror("Error", "Missing columns in the second file: " + ", ".join(missing_columns))
            return

        # Perform the merge on 'MPN'
        merged_data = pd.merge(df_first, df_second[columns_to_merge + ['MPN']], on='MPN', how='left')

        # Save the merged data to a new Excel file
        save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if save_file_path:
            merged_data.to_excel(save_file_path, index=False)
            messagebox.showinfo("Success", "Data merged and saved successfully!")
        else:
            messagebox.showwarning("Cancelled", "File save cancelled.")

    except Exception as e:
        messagebox.showerror("Error", "An error occurred: " + str(e))
