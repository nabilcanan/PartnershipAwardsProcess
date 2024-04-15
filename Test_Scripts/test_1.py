import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


def read_file(file_path):
    if file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    elif file_path.endswith(('.xls', '.xlsx')):
        return pd.read_excel(file_path)
    else:
        print("Unsupported file format.")
        return None


def select_file(file_label):
    file_path = filedialog.askopenfilename(title=f"Select your {file_label} File")
    return file_path


def compare_files():
    cpn_file_path = select_file("CPN")
    backlog_file_path = select_file("Backlog")
    sales_history_file_path = select_file("Sales History")

    cpn_df = read_file(cpn_file_path)
    backlog_df = read_file(backlog_file_path)
    sales_history_df = read_file(sales_history_file_path)

    if cpn_df is None or backlog_df is None or sales_history_df is None:
        messagebox.showerror("Error", "One or more files could not be read.")
        return

    cpn_values = cpn_df.iloc[:, 0]
    cpn_df['In Backlog'] = cpn_values.isin(backlog_df['Backlog CPN'])
    cpn_df['In Sales History'] = cpn_values.isin(sales_history_df['Last Ship CPN'])

    display_results(cpn_df)


# displaying results to be filtered accordingly on backlog file and the difference between the sales
# history file this way we can analyze the CPN accordingly to bring in the columns and check for errors


def display_results(df):
    result_window = tk.Toplevel()
    result_window.title("CPN Comparison Results")

    tree = ttk.Treeview(result_window)
    tree["columns"] = list(df.columns)
    tree["show"] = "headings"

    for column in tree["columns"]:
        tree.heading(column, text=column)
        tree.column(column, anchor="center")

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)

    tree.pack(expand=True, fill='both')

    # Filter buttons
    filter_backlog_button = tk.Button(result_window, text="Filter by In Backlog",
                                      command=lambda: filter_data(tree, df, 'In Backlog'))
    filter_backlog_button.pack()

    filter_sales_history_button = tk.Button(result_window, text="Filter by In Sales History",
                                            command=lambda: filter_data(tree, df, 'In Sales History'))
    filter_sales_history_button.pack()


def filter_data(tree, df, column):
    tree.delete(*tree.get_children())
    sorted_df = df.sort_values(by=column, ascending=False)
    for row in sorted_df.to_numpy().tolist():
        tree.insert("", "end", values=row)


# Setting up the GUI
root = tk.Tk()

root.title("CPN File Comparison Tool")
root.geometry("300x300")

compare_button = tk.Button(root, text="Select and Compare Files", command=compare_files)
compare_button.pack()

root.mainloop()
