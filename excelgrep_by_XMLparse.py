import xml.etree.ElementTree as ET
import glob
import csv
import sys
import io
import codecs
import pandas as pd

from datetime import datetime, time
from datetime import timedelta

#
# python excelgrep_by_XMLparse.py sharedStrings.xml sheet1.xml
#


print("============ ここから excelgrep_by_XMLparse.py ============")

#print("TEST",sDateTime)
#sys.exit()


print("version",pd.__version__)
#pd.set_option('display.max_rows', 70)
pd.set_option('display.max_rows', None)
pd.options.display.colheader_justify = 'left' #列名表示の右寄せ

print('sys.stdout.encoding:', sys.stdout.encoding)
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')  # default でutf-8なのに、これをしないと文字化けする。なぜ？？
print('sys.stdout.encoding:', sys.stdout.encoding)
#sys.exit()




args = sys.argv
print("Arg[sharedStrings.xml]:\t",args[1])
print("Arg[sheet1.xml]:\t",args[2])


    
    
    

#   sharedStrings.xml のsiタグの部分だけ配列に格納
sslist = []
xmls = glob.glob(args[1], recursive=True)
for xml in xmls:
#    print("xml file=",xml)
    tree = ET.parse(xml)
    root = tree.getroot()
    for ssl in root:
        for child in ssl.iter():
                   
            if child.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si':   # 特定要素(si)の抽出
#                print("child.tag = ", child.tag)
                
                for child2 in child.iter():
                     if child2.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t':
                        #print("Hit child2.text= ",child2.text)
                        sslist.append(child2.text)
                        break


maxsslit = len(sslist)
print("lenght of maxsslit = ",maxsslit)
#!for index, item in enumerate(sslist):
#!    print("Index:",str(index)," value:",item)
#!    print("Index:",str(index)," value:",item.encode('cp932', 'replace').decode("cp932", errors="replace"))    #   s-jisにバイト型にエンコードして、s-jisでstr型にデコードにしてprint    CP932に存在しない文字は、'?'に置き換わるとともにエラーを回避できます。

# $ ./excelgrep_by_XMLparse.sh 'dcct' SP8/2024_06_SP8.xlsm だと問題ないのに、
# $ ./excelgrep_by_XMLparse.sh 'dcct' SP8/2024_06_SP8.xlsm  | xargs -I{} grep --color -iE 'dcct'  {}
# すると以下のようになる。   コマンドプロンプトの文字コート変えても同じ。
# xargsで渡すのが問題なのか？？？？
#     
#    print("インデックス：" + str(index) + ", 値：" + item)
#       UnicodeEncodeError: 'cp932(shift_jis)' codec can't encode character '\xa0' in position 14: illegal multibyte sequence
#   print("インデックス：" + str(index) + ", 値：" + item.encode('cp932', "ignore"))
#    TypeError: can only concatenate str (not "bytes") to str      参考：grep は-aオプションでバイトっぽいのをごり押しできるが
#    print("インデックス：" + str(index) + ", 値：" + item.replace('\xa0','',regex=True))
#   TypeError: str.replace() takes no keyword arguments




    
#=====================================================================================================    
#   sheet1.xml のA,B,C列をピックアップ
xmls = glob.glob(args[2], recursive=True)

columns = ['A', 'B', 'C', 'DT'] # DTはA(日付)とB(時間)を日時にしたものを入れる
df = pd.DataFrame(columns=columns) 
df.style.set_properties(**{'text-align': 'left'})   # pip install Jinja2  左寄せ　うまくいかず、、、
df.style.background_gradient(cmap='viridis', low=.5, high=0) # 連続値のグラデーション背景 Matplotlib colormapのviridisにして、0.0 - 5.0のレンジでグラデーション
df.style.set_properties(**{'background-color': 'black', # 背景
                           'color': 'lawngreen', # 文字色
                           'border-color': 'white', # 枠の色っぽいが、変わってない？
                           'align':'left'}) # 文字の揃える位置っぽいが、変わってない？

df_tmp = pd.DataFrame(index=[1],columns=columns)

for xml in xmls:
#    print("xml file=",xml)
    tree = ET.parse(xml)
    root = tree.getroot()
    for sheetData in root:
        for child in sheetData.iter():                        
            if child.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row':   # 特定要素(row)の抽出
                CELL = "-"
                for child2 in child.iter():
                    if child2.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c':    # 特定要素(c)の抽出
                        if not child2.attrib["r"].find('A') or not child2.attrib["r"].find('B') or not child2.attrib["r"].find('C'):
                            for child3 in child2.iter():
                                CELL = child2.attrib["r"]
                                #print("\t",CELL,end='')
                                VAL = "-"
                                if child3.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v':
                                    #print("Hit child3.tag =   ",child3.text)
                                    if not child2.attrib["r"].find('A'):
                                        VAL = int(child3.text)
#                                        print("\t",CELL,"\t",VAL)
                                        df_tmp.iloc[0, 0] = VAL
                                    if not child2.attrib["r"].find('B'):
                                        VAL = float(child3.text)
#                                        print("\t",CELL,"\t",VAL,end='')
#                                        df_tmp.iloc[0, 1] = VAL*24 #時間に変換
#                                        df_tmp.iloc[0, 1] = (datetime(1899,12,30) + timedelta(VAL)).strftime('%H:%M')
                                        df_tmp.iloc[0, 1] = VAL
                                    if not child2.attrib["r"].find('C'):
#                                        print("\t",CELL,"\t",VAL)
                                        try:
                                            if int(child3.text) < maxsslit:
                                                VAL = sslist[int(child3.text)]
                                            else:
                                                VAL = child3.text
                                        except:
                                            VAL = child3.text
                                        #print("\t",CELL,"\t",VAL,end='')
                                        df_tmp.iloc[0, 2] = VAL.ljust(500)  #左寄せ
                            
#                            print("TYPE =\t",type(df_tmp.iloc[0, 0]),"\t",type(df_tmp.iloc[0, 1]),end='\n')
#                            print("VALUE =\t",df_tmp.iloc[0, 0],"\t",df_tmp.iloc[0, 1],end='\n')
                            try:
                                df_tmp.iloc[0, 3] = datetime(1899,12,30) + timedelta(df_tmp.iloc[0, 0]+df_tmp.iloc[0, 1]) #　B列(時間)がない場合、例外が発生するので、その時は00:00にするしかない
                            except:
                                df_tmp.iloc[0, 3] = datetime(1899,12,30) + timedelta(df_tmp.iloc[0, 0]) 

                df = pd.concat([df, df_tmp], ignore_index=True, axis=0)  # 行の結合 concat　　axis=0は縦方向に追加する　1だと横
                df_tmp.iloc[0, 2] = "-" # 次の行への準備。C列(内容部分)だけクリア、A、B列は日時なのでクリアしたくない

#!                if not df_tmp['C'].hasnans:     # C列(内容部分)に値があるときだけ
#!                    df = pd.concat([df, df_tmp], ignore_index=True, axis=0)  # 行の結合 concat　　axis=0は縦方向に追加する　1だと横

    try:
        df['A'] = pd.to_timedelta(df['A'],unit='D',errors="coerce")+pd.to_datetime("1899/12/30")    #  errors=‘coerce’, then invalid parsing will be set as NaT.  ‘ignore’, then invalid parsing will return the input.
    except:
        print('Error')


#    df = df.replace('\uff5e', '-',regex=True).replace('\uff0d', '-',regex=True).replace('\xa0', '',regex=True)         #shift-jisにない文字を置換
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
#    print(df)

#df.drop(df[(df['A'] <= 3) & (df['B'] == 'orange')].index, inplace=True)
    df.drop(df[(df['C'] == "-") ].index, inplace=True)
    df.drop(df[df['C'].str.contains('>本シフトの運転状況<',case=False,na=False)].index, inplace=True) 
    df.drop(df[df['C'].str.contains('シフト交替',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('シフトリーダー:',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('オペレーター:',case=False,na=False)].index, inplace=True)
#    df.drop(df[df['C'].str.contains('',case=False,na=False)].index, inplace=True)

#大文字小文字を無視したい場合は、case=False,NaNを無視するには、na=False


    print(df.loc[:,['DT', 'C']])
    

    print(f"type: {type(df)}")
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

#for index, row in df.iterrows():
#    print(f"Index: {index}", f"A: {row['A']}, B: {row['B']}, C: {row['C']}, DT: {row['DT']}")
    
    

    




