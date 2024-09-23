import xml.etree.ElementTree as ET
import glob
import csv
import numpy as np
import re
import sys
import pandas as pd
print("version",pd.__version__)

#
#   python excelgrep_by_Pandas.py 'SP8/*.xlsm' 'dcct'
#   python excelgrep_by_Pandas.py 'SP8/2024_06_SP8.xlsm' 'dcct'
#   
#   python excelgrep_by_Pandas.py 'SP8/2024_06_SP8.xlsm' 'S.S'      # SESを検索
#   python excelgrep_by_Pandas.py 'SP8/2024_06_SP8.xlsm' 'CB\d{0,1}'    # CB
#   python excelgrep_by_Pandas.py 'SACLA/2024_05.xlsm' 'CB\d{0,1}' 
#
#   問題：SACLA/2024_05.xlsmが読み込めない
#

args = sys.argv
print("検索対象:\t",args[1])
print("検索文字列:\t",args[2])
 
pd.set_option('display.max_rows', 10000)

files = glob.glob(args[1], recursive=True)

for file in files:

    if file.__contains__("~"):
        continue

    print('\n\nexcel.exe ',file,'& ----------------------------------------------------------------')
    df = pd.read_excel(file, sheet_name=0,header=None,engine="openpyxl")
    print(df.index)
    print(df.info)
    print(df.head)
    print(df[df[2].str.contains(args[2], na=False, case=False)])
    
    del df





