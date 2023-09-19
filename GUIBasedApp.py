import os
import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
from tkinter import messagebox
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException
from tqdm import tqdm
import csv
import datetime

final_list = []
combined_df = pd.DataFrame()

def getFileList(input_folder):
    res = []
    for path in os.listdir(input_folder):
        if path.endswith('xlsx'):
            res.append(path)
    return res

def getCount(input_folder):
    count = 0
    for path in os.listdir(input_folder):
        if path.endswith('xlsx'):
            count += 1
    return count

def getIndex(df):
    col_name = df.keys()[3]
    series = df[col_name]
    idx = series[series == "TOTAL NO PCS."].index[0]
    return idx

def buildCurrList(bill, df, index):
    col_name = df.keys()[3]  # name
    amount_col_name = df.keys()[8]  # All aggregations and total
    piece_count_col_name = df.keys()[6]  # no. of pcs.

    res_list = [
        bill,
        df[piece_count_col_name][index],
        df[amount_col_name][index + 1],
        df[amount_col_name][index + 2],
        df[amount_col_name][index + 3],
        df[amount_col_name][index + 4],
        df[amount_col_name][index + 5],
        df[amount_col_name][index + 6],
        df[amount_col_name][index + 7],
        df[amount_col_name][index + 8],
        df[amount_col_name][index + 9],
        df[amount_col_name][index + 10],
    ]
    return res_list

def generate_report(input_folder, report_name):
    print("Generating Report ...")
    with open(os.path.join(input_folder, report_name + "_SaleReport.csv"), 'w', newline='', encoding="utf-8-sig") as f:
        fields = ['Creation Date', 'Party', 'TOTAL NO PCS.', 'TOTAL AMOUNT', 'Murti Dukan', 'GL  Dukan',
                  'LOADING CHARGE', 'TRANSPORTATION', 'PACKING CHARGES', 'Dues', 'GRAND TOTAL', 'ADVANCE', 'PAYABLE']
        write = csv.writer(f)
        write.writerow(fields)
        write.writerows(final_list)
    print("Report Generated Successfully ...")
    messagebox.showinfo("Report Generated", "Report Generated Successfully!")

def prepare_aggregations(res, input_folder):
    for bill in tqdm(res, desc="Calculating the Aggregations ..."):
        df = pd.read_excel(os.path.join(input_folder, bill), engine='openpyxl')
        curr_list = buildCurrList(bill, df, getIndex(df))
        final_list.append(curr_list)

def generate(input_folder, output_file_name):
    global combined_df
    print("Total bills found: ", getCount(input_folder))
    res = getFileList(input_folder)
    print(len(res))
    for file in res:
        excel_file_path = os.path.join(input_folder, file)
        df = pd.read_excel(excel_file_path, sheet_name='Summary Sheet')
        curr_list = df.iloc[0].tolist()

        creation_date = datetime.datetime.fromtimestamp(os.stat(excel_file_path).st_ctime)
        print(creation_date)
        curr_list = [creation_date] + [file] + curr_list
        print(curr_list)
        final_list.append(curr_list)
    generate_report(input_folder, output_file_name)
    print("Aggregations done, calculating and storing data ...")

def select_input_folder():
    input_folder = filedialog.askdirectory()
    if input_folder:
        input_folder_label.config(text=input_folder)

def execute_script():
    input_folder = input_folder_label.cget("text")
    if not input_folder:
        messagebox.showerror("Error", "Please select an input folder.")
        return

    report_name = simpledialog.askstring("Report Name", "Enter the Report Name:")
    if not report_name:
        return

    output_file_name = report_name + '.xlsx'
    generate(input_folder, output_file_name)

def open_generated_csv():
    global generated_report_filename
    if generated_report_filename:
        os.system("start excel " + generated_report_filename)
    else:
        messagebox.showerror("Error", "No report has been generated yet.")

# Create the main window
root = tk.Tk()
root.title("Report Generator")

# Create and configure widgets
select_folder_button = tk.Button(root, text="Select Input Folder", command=select_input_folder)
execute_button = tk.Button(root, text="Execute Script", command=execute_script)
open_csv_button = tk.Button(root, text="Open Generated CSV", command=open_generated_csv)
input_folder_label = tk.Label(root, text="")

# Layout widgets
select_folder_button.pack(pady=10)
input_folder_label.pack()
execute_button.pack(pady=10)
open_csv_button.pack(pady=10)

# Start the GUI event loop
root.mainloop()
