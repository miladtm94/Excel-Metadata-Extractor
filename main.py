################################################
### App name:   Excel MetaData Extractor     ###
### Scripted by: Milad Tatar Mamaghani       ###
### Date:       11/03/2022  (version 1)      ###
################################################

import pandas as pd

filename = ''


def extractExcel(filename):
    input_excel = pd.ExcelFile(filename)
    dfs = pd.read_excel(filename, sheet_name=None)
    cols = [df.columns for i, df in dfs.items()]
    rows = [i for i, df in dfs.items()]
    res = dict(zip(rows, cols))
    df = pd.DataFrame.from_dict(data=res, orient='index').T
    df.to_excel("Result.xlsx", index=False)
    print(df)


def print_hi(name):
    # Use a breakpoint in the code line below to debug your script.
    print(f'Hi, Welcome to my Application, {name},  Enjoy!')  # Press Ctrl+F8 to toggle the breakpoint.


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    print_hi('Excel MetaData-Extractor')
    filename = input('Enter your Excel file name: ')
    filename = filename + '.xlsx'
    extractExcel(filename)
