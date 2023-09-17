import os
import pandas as pd
from openpyxl.utils.exceptions import InvalidFileException
from tqdm import tqdm 
import csv 

final_list = []
combined_df = pd.DataFrame()

def getFileList(Input_folder):
        res = []
        for path in os.listdir(Input_folder):
            if path[-4:] == 'xlsx':
                res.append(path)
        return res

def getCount():
    count = 0
    for path in os.listdir(os.path.dirname(os.path.realpath(__file__))):
        if path[-4:] == 'xlsx':
            count+= 1
    return count

def getIndex(df,):
        col_name = df.keys()[3]
        series = df[col_name]
        idx = series[series == "TOTAL NO PCS."].index[0]
        return idx

def buildCurrList(bill, df, index):
    col_name = df.keys()[3]               #name
    amount_col_name = df.keys()[8]        #All aggregations and total
    piece_count_col_name = df.keys()[6]   #no. of pcs.
    
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

def GenerateReport(Input_folder, report_name):
    print("Generating Report ...")
    with open( os.path.join(Input_folder, report_name + "_SaleReport.csv"), 'w', newline='',encoding="utf-8-sig") as f:
        fields = ['Party', 'TOTAL NO PCS.', 'TOTAL AMOUNT', 'Murti Dukan', 'GL  Dukan', 'LOADING CHARGE', 'TRANSPORTATION', 'PACKING CHARGES', 'Dues', 'GRAND TOTAL', 'ADVANCE', 'PAYABLE']
        write = csv.writer(f)
        write.writerow(fields)
        write.writerows(final_list)
    print("Report Generated Successfully ...")

def prepareAggregations(res, Input_folder, ):
    for bill in tqdm(res, desc = "Calculating the Aggregations ..."):
        df = pd.read_excel(os.path.join(Input_folder, bill),engine='openpyxl')
        curr_list = buildCurrList(bill, df, getIndex(df))
        final_list.append(curr_list)

def generate(input_folder, output_file_name):
    global combined_df 
    print("Total bills found: ", getCount())
    res = getFileList(input_folder)
    print(len(res))
    for file in res:
        excel_file_path = os.path.join(input_folder, file)
        df = pd.read_excel(excel_file_path, sheet_name='Summary Sheet')
        curr_list =  df.iloc[0].tolist()
        curr_list = [file] + curr_list
        print(curr_list)
        final_list.append(curr_list)
    GenerateReport(input_folder, output_file_name)
    print("Aggregrations done, calculating and storing data ...")

def main():
    folder_path = os.path.dirname(os.path.realpath(__file__))
    print("The folder path is::", folder_path)

    report_name = input("Enter the Report Name: ") 

    Input_folder = folder_path
    output_file_name = report_name + '.xlsx'
    generate(Input_folder, output_file_name)


if __name__ == "__main__":
    main()
    


