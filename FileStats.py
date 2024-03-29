import os
import pandas as pd
import xlsxwriter as xw
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
#
file_path = None
delimiter = None


def open_file():
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return


def get_delimiter():
    global delimiter
    delimiter = simpledialog.askstring(title="Select Delimiter",
                                       prompt="Enter the delimiter for the selected CSV file:")
    if not delimiter:
        messagebox.showerror("Error", "No delimiter entered.")
        return


def run_program():
    global file_path, delimiter
    try:
        df = pd.read_csv(file_path, sep=delimiter)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")
        return
    input_file_name = os.path.basename(file_path)
    input_file_name_without_ext = os.path.splitext(input_file_name)[0]
    output_file_name = f"{input_file_name_without_ext}_fillrate.xlsx"
    if os.path.exists(output_file_name):
        os.remove(output_file_name)
    workbook = xw.Workbook(output_file_name)
    worksheet = workbook.add_worksheet()

    worksheet.write_row("A1", ["column name", "total count", "fill count", "percent_fill_rate", "blank count",
                               "percent blank count", "zero count", "percent zero count"])

    for i, col in enumerate(df.columns):
        filled_count = df[col].count()
        total_count = len(df)
        fill_rate = filled_count / total_count
        blank_count = df[col].isna().sum()
        zero_count = (df[col] == 0).sum()
        blank_rate = blank_count / total_count
        zero_rate = zero_count / total_count

        worksheet.write(i + 1, 0, col)
        worksheet.write(i + 1, 1, total_count)
        worksheet.write(i + 1, 2, filled_count)
        worksheet.write(i + 1, 3, fill_rate)
        worksheet.write(i + 1, 4, blank_count)
        worksheet.write(i + 1, 5, blank_rate)
        worksheet.write(i + 1, 6, zero_count)
        worksheet.write(i + 1, 7, zero_rate)

    workbook.close()
    messagebox.showinfo("Success", "The fill rate chart has been created successfully.")
    root.destroy()

root = tk.Tk()
root.title("Fill Rate Calculator")
file_label = tk.Label(root, text="Select a CSV file:")
file_label.grid(row=0, column=0, padx=10, pady=10)
file_button = tk.Button(root, text="Select File", command=open_file)
file_button.grid(row=0, column=1, padx=10, pady=10)
delimiter_label = tk.Label(root, text="Select a Delimiter:")
delimiter_label.grid(row=1, column=0, padx=10, pady=10)
delimiter_button = tk.Button(root, text="Select Delimiter", command=get_delimiter)
delimiter_button.grid(row=1, column=1, padx=10, pady=10)
run_button = tk.Button(root, text="Run", command=run_program)
run_button.grid(row=2, column=1, padx=10, pady=10)
root.mainloop()

