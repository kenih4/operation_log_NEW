import xml.etree.ElementTree as ET
import glob
import csv
import sys
import io
import codecs
import pandas as pd
from pandas.errors import EmptyDataError
from pathlib import Path

from datetime import datetime, time
from datetime import timedelta

import tkinter as tk #ポップアップメッセージを出したい
from tkinter import messagebox
root = tk.Tk()# Tkinterのルートウィンドウを作成するが、非表示にする
root.withdraw()

#
# python excelgrep_by_XMLparse_for_Untenshyukei.py sharedStrings.xml sheet1.xml
#
# TEST  2024/10
# python excelgrep_by_XMLparse_for_Untenshyukei.py C:/Users/kenichi/AppData/Local/Temp/tmp.XwS6GHBs35/xl/sharedStrings.xml C:/Users/kenichi/AppData/Local/Temp/tmp.XwS6GHBs35/xl/worksheets/sheet1.xml
#
# # Formatter     Shift+Alt+F
#
print("============ ここから excelgrep_by_XMLparse.py ============")

#print("TEST",sDateTime)
#sys.exit()


# 行番号と色を引数にする関数
def highlight_rows(x, rows_to_highlight, color="yellow"):
    style = f'background-color: {color}'
    return [style if x.name in rows_to_highlight else '' for _ in x]

# 列全体に色を付ける関数
def highlight_column_BL2(val):
    return 'background-color: gold'
def highlight_column_BL3(val):
    return 'background-color: dodgerblue'

# 特定の文字列が含まれる行に色を付ける関数 Not use
#def highlight_syuryo(row):
#    return ['background-color: red' if '終了' in str(row['C']) else '' for _ in row]

# 指定されたExcelファイルとシート名を読み込み、pandasのDataFrameとして返します　==============================================================================================
def load_excel_to_dataframe(file_path_str: str, sheet_name: str) -> pd.DataFrame | None:
    """
    Args:
        file_path (str): 読み込むExcelファイル名 (例: 'test.xlsm')。
        sheet_name (str): 読み込むシート名 (例: 'sheet1')。

    Returns:
        pd.DataFrame | None: 読み込まれたDataFrame、またはエラーが発生した場合はNone。
    """
    try:
        # Excelファイルを読み込み、DataFrameに格納
        file_path = Path(file_path_str) 
        print(f"file:'{file_path_str}', sheet:'{sheet_name}'")
        df = pd.read_excel(
            file_path,        # ファイルパス
            sheet_name=sheet_name, # 読み込むシート名
            engine='openpyxl'  # .xlsm/.xlsxファイルに対応
        )
#        print("Complete reading Excel file into DataFrame.")
        return df

    except FileNotFoundError:
        print(f"エラー: ファイル '{file_path}' が見つかりません。ファイルパスを確認してください。")
        return None
    except ValueError as e:
        # シート名が存在しない場合などに発生
        print(f"エラー: 指定されたシート '{sheet_name}' が見つからないか、その他の読み込みエラーが発生しました。詳細: {e}")
        return None
    except EmptyDataError:
        print(f"警告: ファイル '{file_path}' のシート '{sheet_name}' にデータが含まれていません。")
        # 空のDataFrameを返すこともできます
        # return pd.DataFrame()
        return None
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        return None

#Pandas DataFrameの指定された列に、特定の日時データ(±許容時間内)が存在するかを確認します。==============================================================================================
from typing import Union, List
def check_datetime_existence_with_tolerance(
    df: pd.DataFrame, 
    column_name: str, 
    date_to_check: datetime, 
    tolerance_minutes: int = 1
) -> Union[int, List[int]]:
    """
    Pandas DataFrameの指定された列に、特定の日時データ(±許容時間内)が存在するかを確認し、
    存在する場合は対応するインデックスを返します。
    
    Args:
        df (pd.DataFrame): 検索対象のDataFrame。
        column_name (str): 検索対象の日時列名。
        date_to_check (datetime): 存在を確認したい基準日時。
        tolerance_minutes (int): 許容時間 (分)。デフォルトは1分。
        
    Returns:
        Union[int, List[int]]: 条件を満たす行のインデックス。
                              該当する行が複数ある場合はインデックスのリスト。
                              存在しない場合は -1 を返します。
    """
    
    # 1. 前処理とエラーハンドリング
    try:
        # 列のデータ型をdatetimeに揃える (元の関数と同様の処理を推奨)
        if not pd.api.types.is_datetime64_any_dtype(df[column_name]):
            # utc=Trueで一度UTCに変換し、tz_localize(None)でタイムゾーン情報を除去（比較をしやすくするため）
            # ただし、date_to_checkもタイムゾーン情報がない（naive）前提での処理
            df[column_name] = pd.to_datetime(df[column_name], errors='coerce', utc=True).dt.tz_localize(None)
    except KeyError:
        print(f"エラー: 指定された列名 '{column_name}' がDataFrameに存在しません。")
        return -1
    except Exception as e:
        print(f"日時変換中にエラーが発生しました: {e}")
        return -1

    # 2. 許容範囲の計算
    tolerance = timedelta(minutes=tolerance_minutes) 
    start_time = date_to_check - tolerance
    end_time = date_to_check + tolerance 
    
    # 3. 条件を満たす行のインデックスを取得
    
    # 日時が許容範囲内にあるかどうかの真偽値シリーズを作成
    mask = df[column_name].between(start_time, end_time)
    
    # 条件を満たす行のインデックスを取得
    matching_indices = df[mask].index.tolist()
    
    # 4. 結果の返却
    if matching_indices:
        # 該当するインデックスが存在する場合
        # 1つだけ見つかった場合はそのインデックス (int) を返し、
        # 複数見つかった場合はリスト (List[int]) を返す
        if len(matching_indices) == 1:
            return matching_indices[0]
        else:
            return matching_indices
    else:
        # 該当するインデックスが存在しない場合は -1 を返す
        return -1

# ***********************************
# 💡 注意: 戻り値の型を Union[int, List[int]] としています。
# 該当行が複数見つかる可能性があるためです。
# 常に最初のインデックスだけが欲しい場合は、df[mask].index[0] のように変更可能です。
# ***********************************

#ical用　始め　=============================================================================================
import requests
from requests.exceptions import Timeout
import re
import pandas as pd
import sys
from icalendar import Calendar, Event
#Japanese
import locale
dt = datetime(2018, 1, 1)
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))
locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
print(locale.getlocale(locale.LC_TIME))
#print(dt.strftime('%A, %a, %B, %b'))
config_file_sig = "ical_SACLA.xlsx"
df_sig = pd.read_excel(config_file_sig, sheet_name="sig")
# print(df_sig)

def get_ical(url):
    #print(url)
    try:
        res = requests.get(url, timeout=(30.0, 30.0))
    except Exception as e:
        #print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!@get_ical	' + url)
        print(e.args)
        return ''
    else:
        res.raise_for_status()
        return res.text

class SigInfo:
    def __init__(self):
        self.srv = ''
        self.url = ''
        self.sname = ''
        self.sid = 0
        self.sta = ''
        self.sto = ''
        self.time = ''
        self.val = ''
        self.sortedval = []
        self.rave = []
        self.rave_sigma = []
        self.d = {}
        self.t = {}
        self.mu = 0
        self.icaldata = ''
        self.sigma = 0

sig = [SigInfo() for _ in range(len(df_sig))]

from datetime import datetime, timedelta, timezone
JST = timezone(timedelta(hours=+9), 'JST')

def get_schedule_from_ical(df_lognote):
#    print(df_lognote)
    for n, s in enumerate(sig, 0):
        print("label: ",str(df_sig.loc[n]['label']))
        s.icaldata = get_ical(str(df_sig.loc[n]['url']))
        #print(s.icaldata)
        cal = Calendar.from_ical(s.icaldata)

        for index,item in df_lognote.iterrows():
#            print("index : ", index, "  item['DT'] = ", item['DT'])
            for ev in cal.walk():

                if ev.name == 'VEVENT':
                    start_dt = ev.decoded("dtstart")
                    end_dt = ev.decoded("dtend")
                    try:
                        summary = ev['summary']
                    except Exception as e:
                        print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                    else:
                        try:
                            #print('type[item[DT]] = ',type(item['DT']))
                            if (item['DT'].astimezone(JST) - start_dt).total_seconds() > 0 and (item['DT'].astimezone(JST) - end_dt).total_seconds() < 0:
                                
                                tmp_summary = str(summary).replace(' ', '')
                                tmp_summary = re.sub("（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す
                                tmp_summary = tmp_summary.rstrip('<br>')
                                tmp_summary = tmp_summary.replace("/30Hz", "")
                                tmp_summary = tmp_summary.replace("/60Hz", "")
                                
                                if(df_sig.loc[n]['label']=="BL2"):
                                    df_lognote.loc[index, 'BL2ical'] = tmp_summary
                                elif(df_sig.loc[n]['label']=="BL3"):
                                    df_lognote.loc[index, 'BL3ical'] = tmp_summary
                                #print("index = ", index, "  item['DT']= ",item['DT'],"   :    ", tmp_summary)                                
                                continue
                        except:
                            pass    #print('Exception!!!!!!!!!!!!!!')
#    print(df_lognote.loc[:,['DT','BL3ical', 'C']])
#                        else:
#                            print('type(item[DT]) = ', type(item['DT']), '   item[DT] = ' , item['DT'])
#ical用　終わり=============================================================================================




print("version",pd.__version__)
#pd.set_option('display.max_rows', 70)
pd.set_option('display.max_rows', None)

#pd.set_option("display.max_colwidth", 2000) # #カラム内の文字数
pd.options.display.max_colwidth = 2000

pd.set_option('display.width', 1000) # 少ないと改行されてしまうので増やす
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
if maxsslit == 0:
    print("sharedStrings.xml のsiタグの数がlenght of maxsslit = ",maxsslit,"  です。終了します")
    sys.exit()

    
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

columns = ['A', 'B', 'C', 'DT', 'formatted_DT','BL1ical', 'BL2ical', 'BL3ical'] # DTはA(日付)とB(時間)を日時にしたものを入れる
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
                                try:
                                    df_tmp.iloc[0, 3] = datetime(1899,12,30) + timedelta(df_tmp.iloc[0, 0]) 
                                except:
                                    df_tmp.iloc[0, 3] = 0
                            
                            #print(">>>df_tmp =\t",type(df_tmp.iloc[0, 3]),"\t",df_tmp.iloc[0, 3],end='\n')
                            
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


#消すな将来用    df['C'] = df['C'].replace({'終了時': '終了 時', '終了後': '終了 後'},regex=True) # とりあえずテスト的。今のところ意味ない

    df.drop(df[(df['C'] == "-") ].index, inplace=True)
    df.drop(df[df['C'].str.contains('>本シフトの運転状況<',case=False,na=False)].index, inplace=True) 
    df.drop(df[df['C'].str.contains('シフト交替',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('シフトリーダー:',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('オペレーター:',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('プロファイル定時確認',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('プロファイル確認',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('BL2: ',case=False,na=False)].index, inplace=True)
    df.drop(df[df['C'].str.contains('BL3: ',case=False,na=False)].index, inplace=True)
    #大文字小文字を無視したい場合は、case=False,NaNを無視するには、na=False
    


    #ログノートA列の日付が00:00を過ぎても日付はそのままなので対処   
    bf_itemDT = datetime(year=2000, month=1, day=1, hour=0, minute=0, second=0)
    for index,item in df.iterrows():
        if(type(item['DT']) is not datetime):
            print('DEBUG ~~~~~~~  type item[DT]=', type(item['DT']))
            continue
        try:
            if((item['DT'] - bf_itemDT).total_seconds() >= 0):
#                print('TIME OK:     ',  item['DT'],"    - ", bf_itemDT, "   =   ", (item['DT'] - bf_itemDT).total_seconds())
                pass
            else:
                if(abs(item['DT'] - bf_itemDT).total_seconds() > 28800): # 2直17:00には絶対時刻があるので、28,800sec=8時間以上開いてる時だけ0時を堺に日付を+1日する                    
                    df.loc[index, 'DT'] = item['DT'] + timedelta(days = 1)
#                    print(index,' TIME INVERT: ',  item['DT']," - ", bf_itemDT, " = ", (item['DT'] - bf_itemDT).total_seconds(), " NEW df.loc[index, 'DT'] = ",  df.loc[index, 'DT']  )
                else:
                    print(index,' TIME INVERT: ログノートの時刻記載が間違ってる可能性があります。',  item['DT']," - ", bf_itemDT, " = ", (item['DT'] - bf_itemDT).total_seconds(), " NEW df.loc[index, 'DT'] = ",  df.loc[index, 'DT']  )
            bf_itemDT = df.loc[index, 'DT']
            df.loc[index, 'formatted_DT'] = df.loc[index, 'DT'].strftime('%Y/%#m/%#d %#H:%#M')
        except Exception as e:
            print(dir(e))
            print("message:{0}".format(e.message))
            pass

    print("この処理には時間が掛かる~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    get_schedule_from_ical(df)
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
#    print(df.loc[1:10,['DT','BL3ical', 'C']])
#    print(df.loc[:,['DT','BL3ical', 'C']])
    
    styler = df.loc[:,['formatted_DT','BL2ical', 'BL3ical', 'C']].style.map(lambda x: 'background-color: skyblue' if ('引渡' or '引き渡') in str(x) else '')
#    styler = styler.map(lambda x: 'color: yellow' if ('変更依頼' or 'ユニット' or '切替') in str(x) else '')
    styler = styler.map(lambda x: 'color: yellow' if ('切替') in str(x) else '')
    styler = styler.map(lambda x: 'color: pink' if ("加速器調整" in str(x) or "BL-study" in str(x) or "BL調整" in str(x)) else '') #　なぜか or　が効かない
    styler = styler.map(lambda x: 'color: red' if ('終了') in str(x) else '')
    styler = styler.applymap(highlight_column_BL2, subset=['BL2ical'])# BL2/BL3 ical列に色付け
    styler = styler.applymap(highlight_column_BL3, subset=['BL3ical'])# BL2/BL3 ical列に色付け
#    styler = styler.apply(highlight_syuryo, axis=1) #　終了のワードがある行に色付け
    styler = styler.set_properties(**{'text-align': 'left'}) #左寄せ    

    for index,item in df.iterrows():
        try:
            if not ("加速器調整" in str(item['BL2ical']) or "BL-study" in str(item['BL2ical']) or "BL調整" in str(item['BL2ical'])) and not ("加速器調整" in str(item['BL3ical']) or "BL-study" in str(item['BL3ical']) or "BL調整" in str(item['BL3ical'])): # ユーザー運転
                if "終了" in str(item['C']):
                    #print('両方ユーザー運転中に終了したぞ！  item[BL2ical]=', str(item['C']))
                    styler = styler.apply(highlight_rows, rows_to_highlight=[index], color="lime", axis=1)
        except Exception as e:
            print(dir(e))
            print("message:{0}".format(e.message))
            pass
    
#    styler.to_excel('output1.xlsx')
#    styler.to_html('hoge.html',index=False)
#   セル内での改行をしたくないのでcssを噛ます    
    css = '''
    <style>
    table {
        white-space: nowrap;
    }
    </style>
    '''
    html_output = css + styler.to_html(index=False)
    with open('output.html', 'w', encoding='utf-8') as f:
        f.write(html_output)
    import webbrowser
#    webbrowser.open_new_tab('hoge.html')
    webbrowser.open_new_tab('output.html')    
#    display(styler)
    
    


    #/=======SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間がログノートに存在するか確認
    #  なぜか、get_schedule_from_ical(df)の前でこれをすると、icalからとってきたスケジュールがうまくdfに入らない。なぜ？？？？
    print("SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間がログノートに存在するか確認~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    for index,item in df.iterrows():#ログノートの最初の日時を取得(ログノートが何月のなのかを確認するため)
        if(type(item['DT']) is datetime):
            first_dt = datetime(year=item['DT'].year, month=item['DT'].month, day=item['DT'].day, hour=0, minute=0, second=0)
            break
    print(first_dt)

    with open(r"C:\me\unten\OperationSummary\dt_beg.txt", mode='r', encoding="UTF-8") as f:
        buff_dt_beg = f.read()
    with open(r"C:\me\unten\OperationSummary\dt_end.txt", mode='r', encoding="UTF-8") as f:
        buff_dt_end = f.read()
    dt_beg = datetime.strptime(buff_dt_beg, "%Y/%m/%d %H:%M")
    dt_end = datetime.strptime(buff_dt_end, "%Y/%m/%d %H:%M")
    print("dt_beg=",dt_beg)
    print("dt_end=",dt_end)            
    df_kiroku = load_excel_to_dataframe(r"\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転集計記録.xlsm", "調整時間")
    start_row_index = 1 # 2行目以降
    column_index = 2 # 'end'列目  調整時間のstartにはチョッパーOFF時間になってる事があるので、ログノートの記載時間と合わないことがあるのでendで確認する。
    ans_line=-1
    for index, value in df_kiroku.iloc[start_row_index:, column_index].items():
        if value.month == first_dt.month and value >= dt_beg and value <= dt_end: #指定された月のログノートで、かつ、運転集計する期間内だけ確認
            result = check_datetime_existence_with_tolerance(df, 'DT', value, tolerance_minutes=0.5)            
            print(f"💡index: {index}, value: {value}, type(value): {type(value)}     検索結果のインデックス: {result}")
            if result == -1:
                print("💡 SACLA運転集計記録test.xlsmのシート[調整時間]に記載されている調整「終了」時間がログノートに存在しません！    " + str(value))
                messagebox.showwarning("要確認", "SACLA運転集計記録test.xlsmの\nシート[調整時間]に\n記載されている調整「終了」時間がログノートに存在しません！\n" + str(value))
                ans_line=result
            elif isinstance(result, int):# 1つのインデックス（int型）が返された場合                
                ans_line=result
            else: # 複数のインデックス（リスト型）が返された場合
                ans_line=result[0]
            
            if ans_line != -1:
                matching_row = df.loc[ans_line,['C']]#matching_row = df.loc[result,['formatted_DT','BL2ical', 'BL3ical', 'C']]
                print(matching_row.to_string(header=False, index=False).replace('\n', ' ').strip())
                if not "引" in matching_row.to_string(header=False, index=False).replace('\n', ' ').strip():
                    messagebox.showwarning("要確認", "SACLA運転集計記録test.xlsmの\nシート[調整時間]に\n記載されている調整「終了」時間がログノートに存在しますが、「引渡」と書かれていませんよ！！\n" + str(value))                
        else:
            print(f"index: {index}, value: {value} is out of range.")
    print("SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間がログノートに存在するか確認　が終了しました。~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")    
    sys.exit()
    #========================================================================/