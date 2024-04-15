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

    process_first_file(file_path_1)
    process_second_file(file_path_2)


def process_first_file(file_path):
    try:
        # Load the workbook and get the active sheet
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

        # Determine the column for the new 'Ext Award Value' formula
        # This should be the next column after the last used column in row 2
        new_column_index = sheet.max_column + 1
        ext_award_value_column = get_column_letter(new_column_index)

        # Add header for new formula column in row 2
        sheet.cell(row=2, column=new_column_index).value = 'Ext Award Value'

        # Insert formula in the new column starting from row 3
        awarded_eau_column = get_column_letter(header_column_map['Awarded EAU'])
        award_price_column = get_column_letter(header_column_map['Award Price'])
        for row in range(3, sheet.max_row + 1):
            sheet[ext_award_value_column + str(row)].value = f'={awarded_eau_column}{row}*{award_price_column}{row}'

        # Save the workbook
        final_save_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if final_save_file_path:
            workbook.save(final_save_file_path)
            messagebox.showinfo("Success", "File processed and saved successfully!")
        else:
            messagebox.showwarning("Cancelled", "Final file save cancelled.")
    except Exception as e:
        messagebox.showerror("Error", "An error occurred: " + str(e))


def process_second_file(file_path):
    # Placeholder for secondary file processing
    pass
