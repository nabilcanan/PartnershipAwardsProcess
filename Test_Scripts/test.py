import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path


def compare_files():
    try:
        file1 = select_file()
        file2 = select_file()

        if file1 and file2:
            df1 = pd.read_excel(file1)
            df2 = pd.read_excel(file2)

            if df1.shape != df2.shape:
                messagebox.showinfo("Info", "Files have different shapes. Cannot compare.")
                return

            diff = ''
            for row in range(df1.shape[0]):
                for col in df1.columns:
                    if df1.at[row, col] != df2.at[row, col]:
                        diff += f'Row: {row + 1}, Column: {col}, Value1: {df1.at[row, col]}, Value2: {df2.at[row, col]}\n'

            text_area.delete('1.0', tk.END)
            text_area.insert(tk.END, diff)

    except Exception as e:
        messagebox.showerror("Error", str(e))


# Set up the GUI
root = tk.Tk()
root.title("Excel File Comparer")

# Add buttons and text area
compare_button = tk.Button(root, text="Compare Excel Files", command=compare_files)
compare_button.pack()

text_area = tk.Text(root, height=15, width=80)
text_area.pack()

# Run the application
root.mainloop()
