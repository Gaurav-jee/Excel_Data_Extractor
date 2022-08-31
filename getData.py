import pandas as pd
import numpy as np
from tqdm import tqdm
import os
import csv

# folder path
dir_path = r'C:\Bill DashBoard Testing\new bills'

# list to store files
res = []

# Iterate directory
for path in os.listdir(dir_path):
    # check if current path is a file
    if os.path.isfile(os.path.join(dir_path, path)):
        print(path)
    res.append(path)
# print(res)


# file_ = pd.read_excel(r'testing.xlsx')
# file_


APP_PATH = 'C:\Bill DashBoard Testing'

# df1 = pd.read_excel(
#      os.path.join(APP_PATH, "new bills", "new (5).xlsx"),
#      engine='openpyxl',
# )

# print(df1.head())

# ls = list(df1.keys())

# print(ls)
# print(df1['Unnamed: 8'][1991])

final_output = {}
sum_total = 0

# from tqdm import tqdm
  
# for i in tqdm (range (100), desc="Loading..."):
#     pass


for bills in tqdm(res, desc = "calculating sum of bills"):
    df1 = pd.read_excel(
     os.path.join(APP_PATH, "new bills", bills),
     engine='openpyxl')
    curr_val = df1['Unnamed: 8'][1991] 
    if pd.isna(curr_val) == False:
        sum_total += curr_val
    curr_dict = {bills : curr_val}
    final_output.update(curr_dict)
        
print("************************" + str(sum_total) + "********************************")


with open('Output.csv', 'w', encoding="utf-8") as f:
    for key in final_output:
        f.write("%s,%s\n"%(key,final_output[key]))

    # final_output.update(bills = df1['Unnamed: 8'][1991])
    # sum_total += df1['Unnamed: 8'][1991]

# print(final_output)
# print("TOTAL is", sum_total)

