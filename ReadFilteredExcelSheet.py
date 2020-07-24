# Author: Apurv Srivastav
# email-id: apurv11419@gmail.com

# read filtered excel sheet into your dataframe

import pandas as pd
from openpyxl import load_workbook

def read_filtered_excel_sheet_to_Df(self, excel_file: str, sheet_name: str) -> pd.DataFrame:
    
    book = load_workbook(excel_file)
    sheet = book[sheet_name]

    r = 1
    df = pd.DataFrame()
    for row in sheet:
        if sheet.row_dimensions[row[0].row].hidden == False:
            cell_array = [cell.value for cell in row]
            if r == 1:
                df = pd.DataFrame(columns=cell_array)
                r += 1
            else:
                df.loc[ len(df) ] = cell_array

    return df
