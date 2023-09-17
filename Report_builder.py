import pandas as pd
import numpy as np
from tqdm import tqdm
import os
import csv



def main():
    print("########################## File Aggregation System. #############################")
    print("")
    print("")

    folder_path = os.path.dirname(os.path.realpath(__file__))

    report_name = input("Enter the Report Name: ")    
    # list to store file names
    res = []

    def getFileList():
        for path in os.listdir(os.path.dirname(os.path.realpath(__file__))):
            if path[-4:] == 'xlsx':
                res.append(path)


    def getIndex(df,):
        col_name = df.keys()[3]
        series = df[col_name]
        idx = series[series == "TOTAL NO PCS."].index[0]
        return idx

    final_list = []

    def buildList(bill, df, index):
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

    def prepareAggregations():
        for bill in tqdm(res, desc = "Calculating the Aggregations ..."):
            df = pd.read_excel(os.path.join(folder_path, bill),engine='openpyxl')
            curr_list = buildList(bill, df, getIndex(df))
            final_list.append(curr_list)
        
        print("Aggregrations done, calculating and storing data ...")
        
        def GenerateReport():
            print("Generating Report ...")
            with open( os.path.join(folder_path, report_name + "Report.csv"), 'w',encoding="utf-8") as f:
                fields = ['Party', 'TOTAL NO PCS.', 'TOTAL AMOUNT', '(मूर्ति)  दुकान', '(गणेश लक्ष्मी)  दुकान', 'LOADING CHARGE (लेबर खर्चा)', 'TRANSPORTATION (गाड़ी भाड़ा)', 'PACKING CHARGES (कार्टून, रस्सी, नेवारी)', 'पहेले का बकाया', 'GRAND TOTAL/ कुल', 'ADVANCE/ अग्रिम', 'PAYABLE/ देय']
                write = csv.writer(f)
                write.writerow(fields)
                write.writerows(final_list)
            print("Report Generated Successfully ...")
        GenerateReport()
    
    getFileList()
    prepareAggregations()
    print("")
    print("")
    print("Please find your reports at:", folder_path, "by the name ->", report_name)
    
    
if __name__== "__main__" :
        main()
## -> next step -> https://towardsdatascience.com/how-to-easily-convert-a-python-script-to-an-executable-file-exe-4966e253c7e9
## -> next Step -> https://www.youtube.com/watch?v=svcv8uub0D0&ab_channel=CodingIsFun