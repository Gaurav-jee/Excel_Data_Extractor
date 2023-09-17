import pandas as pd

def ReadExcel(filePath):
    ... 

def main():
    excel_file_path = r"C:\EKNazar_HPWhite\2022\Finished Bills\s k das,     पूर्णिया ,   9631049491.xlsx"
    # Read a specific sheet from the Excel file
    # usecols=['Column6', 'Column7', 'Column8']
    df = pd.read_excel(excel_file_path, sheet_name='Summary Sheet')
    print(df)
    pass
    

if __name__ == "__main__":
    main()