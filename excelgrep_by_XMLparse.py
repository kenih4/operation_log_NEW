import xml.etree.ElementTree as ET
import glob
import argparse
import sys
import io
import pandas as pd

from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Union, List

#
# 使い方:
#   python excelgrep_by_XMLparse.py [--mode search|summary] sharedStrings.xml sheet1.xml
#
#   --mode search  (既定) ログノート検索モード。DT,C列をターミナルに出力する(呼び出し側のgrepで検索)
#   --mode summary        運転集計モード。icalカレンダーを付与してHTML出力し、
#                         SACLA運転集計記録.xlsmの調整時間がログノートに記載されているか確認する
#
# TEST
# python excelgrep_by_XMLparse.py --mode search C:/Users/kenichi/AppData/Local/Temp/tmp.jdpng8Hbvj/xl/sharedStrings.xml C:/Users/kenichi/AppData/Local/Temp/tmp.jdpng8Hbvj/xl/worksheets/sheet1.xml
# python excelgrep_by_XMLparse.py --mode summary C:/Users/kenichi/AppData/Local/Temp/tmp.XwS6GHBs35/xl/sharedStrings.xml C:/Users/kenichi/AppData/Local/Temp/tmp.XwS6GHBs35/xl/worksheets/sheet1.xml
#
# 通常は excelgrep_by_XMLparse.sh から呼ばれる(-k=検索ワード があれば search、なければ summary)
#
# Formatter     Shift+Alt+F

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

COLUMNS = ['A', 'B', 'C', 'DT', 'formatted_DT', 'BL1ical',
           'BL2ical', 'BL3ical']  # DTはA(日付)とB(時間)を日時にしたものを入れる

# 不要行(C列にこの文字列を含む行)。モードごとに異なる。
DROP_WORDS_SEARCH = [
    '>本シフトの運転状況<', 'シフト交替', 'シフトリーダー:', 'オペレーター:', 'プロファイル定時確認',
    # SR LOG特有
    'シフト交代', '運転員:', 'パラメータセーブ', 'バンチ純度測定結果', 'クレーン',
]
DROP_WORDS_SUMMARY = [
    '>本シフトの運転状況<', 'シフト交替', 'シフトリーダー', 'オペレーター', 'プロファイル定時確認',
    'プロファイル確認', 'BL2: ', 'BL3: ',
    # SR LOG特有
    'シフト交代', '運転員', 'パラメータセーブ', 'バンチ純度測定結果', 'クレーン',
]


# ============================================================================================
# 共通処理
# ============================================================================================

def load_shared_strings(path_pattern: str) -> list:
    """sharedStrings.xml のsiタグの部分(最初のtだけ)を配列に格納"""
    sslist = []
    for xml in glob.glob(path_pattern, recursive=True):
        root = ET.parse(xml).getroot()
        for ssl in root:
            for child in ssl.iter():
                # 特定要素(si)の抽出
                if child.tag == NS + 'si':
                    for child2 in child.iter():
                        if child2.tag == NS + 't':
                            sslist.append(child2.text)
                            break
    return sslist


def load_sheet(path_pattern: str, sslist: list) -> pd.DataFrame:
    """sheet1.xml のA,B,C列をピックアップしてDataFrameにする"""
    maxsslit = len(sslist)
    df = pd.DataFrame(columns=COLUMNS)
    df_tmp = pd.DataFrame(index=[1], columns=COLUMNS)

    for xml in glob.glob(path_pattern, recursive=True):
        root = ET.parse(xml).getroot()
        for sheetData in root:
            for child in sheetData.iter():
                # 特定要素(row)の抽出
                if child.tag != NS + 'row':
                    continue
                for child2 in child.iter():
                    # 特定要素(c)の抽出
                    if child2.tag != NS + 'c':
                        continue
                    ref = child2.attrib["r"]
                    if not (not ref.find('A') or not ref.find('B') or not ref.find('C')):
                        continue
                    for child3 in child2.iter():
                        if child3.tag != NS + 'v':
                            continue
                        if not ref.find('A'):
                            try:
                                VAL = int(child3.text)
                            except Exception:
                                VAL = child3.text
                            df_tmp.iloc[0, 0] = VAL
                        if not ref.find('B'):
                            try:
                                VAL = float(child3.text)
                            except Exception:
                                VAL = child3.text
                            df_tmp.iloc[0, 1] = VAL
                        if not ref.find('C'):
                            try:
                                if int(child3.text) < maxsslit:
                                    VAL = sslist[int(child3.text)]
                                else:
                                    VAL = child3.text
                            except Exception:
                                VAL = child3.text
                            df_tmp.iloc[0, 2] = VAL.ljust(500)  # 左寄せ

                    try:
                        # B列(時間)がない場合、例外が発生するので、その時は00:00にするしかない
                        df_tmp.iloc[0, 3] = datetime(
                            1899, 12, 30) + timedelta(df_tmp.iloc[0, 0]+df_tmp.iloc[0, 1])
                    except Exception:
                        try:
                            df_tmp.iloc[0, 3] = datetime(
                                1899, 12, 30) + timedelta(df_tmp.iloc[0, 0])
                        except Exception:
                            df_tmp.iloc[0, 3] = 0

                # 行の結合 concat　　axis=0は縦方向に追加する　1だと横
                df = pd.concat([df, df_tmp], ignore_index=True, axis=0)
                # 次の行への準備。C列(内容部分)だけクリア、A、B列は日時なのでクリアしたくない
                df_tmp.iloc[0, 2] = "-"

    try:
        # errors='coerce'だと変換できない値はNaTになる
        df['A'] = pd.to_timedelta(
            df['A'], unit='D', errors="coerce")+pd.to_datetime("1899/12/30")
    except Exception:
        print('Error')
    return df


def drop_unneeded_rows(df: pd.DataFrame, words: list) -> None:
    """C列が「-」の行と、wordsのどれかを含む行を削除する(dfを直接変更)"""
    df.drop(df[(df['C'] == "-")].index, inplace=True)
    for word in words:
        # 大文字小文字を無視するにはcase=False、NaNを無視するにはna=False。regex=Falseで単純な文字列検索
        df.drop(df[df['C'].str.contains(
            word, case=False, na=False, regex=False)].index, inplace=True)


# ============================================================================================
# 検索モード (--mode search)
# ============================================================================================

def run_search(df: pd.DataFrame) -> None:
    print("print Before drop ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    print(df)

    drop_unneeded_rows(df, DROP_WORDS_SEARCH)

    print(
        "print df.loc[:, [DT, C]]====================================================")
    print(df.loc[:, ['DT', 'C']])

    print(f"type: {type(df)}")
    print("Finish~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")


# ============================================================================================
# 運転集計モード (--mode summary)
# ============================================================================================

# 行番号と色を引数にする関数
def highlight_rows(x, rows_to_highlight, color="yellow"):
    style = f'background-color: {color}'
    return [style if x.name in rows_to_highlight else '' for _ in x]


# 列全体に色を付ける関数
def highlight_column_BL2(val):
    return 'background-color: gold'


def highlight_column_BL3(val):
    return 'background-color: dodgerblue'


def load_excel_to_dataframe(file_path_str: str, sheet_name: str) -> pd.DataFrame | None:
    """
    指定されたExcelファイルとシート名を読み込み、pandasのDataFrameとして返します。

    Args:
        file_path_str (str): 読み込むExcelファイル名 (例: 'test.xlsm')。
        sheet_name (str): 読み込むシート名 (例: 'sheet1')。

    Returns:
        pd.DataFrame | None: 読み込まれたDataFrame、またはエラーが発生した場合はNone。
    """
    from pandas.errors import EmptyDataError
    try:
        file_path = Path(file_path_str)
        print(f"file:'{file_path_str}', sheet:'{sheet_name}'")
        df = pd.read_excel(
            file_path,        # ファイルパス
            sheet_name=sheet_name,  # 読み込むシート名
            engine='openpyxl'  # .xlsm/.xlsxファイルに対応
        )
        return df

    except FileNotFoundError:
        print(f"エラー: ファイル '{file_path}' が見つかりません。ファイルパスを確認してください。")
        return None
    except ValueError as e:
        # シート名が存在しない場合などに発生
        print(
            f"エラー: 指定されたシート '{sheet_name}' が見つからないか、その他の読み込みエラーが発生しました。詳細: {e}")
        return None
    except EmptyDataError:
        print(f"警告: ファイル '{file_path}' のシート '{sheet_name}' にデータが含まれていません。")
        return None
    except Exception as e:
        print(f"予期せぬエラーが発生しました: {e}")
        return None


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
        # 列のデータ型をdatetimeに揃える
        if not pd.api.types.is_datetime64_any_dtype(df[column_name]):
            # utc=Trueで一度UTCに変換し、tz_localize(None)でタイムゾーン情報を除去（比較をしやすくするため）
            # ただし、date_to_checkもタイムゾーン情報がない（naive）前提での処理
            df[column_name] = pd.to_datetime(
                df[column_name], errors='coerce', utc=True).dt.tz_localize(None)
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
    mask = df[column_name].between(start_time, end_time)
    matching_indices = df[mask].index.tolist()

    # 4. 結果の返却 (1つだけならint、複数ならリスト、なければ-1)
    if matching_indices:
        if len(matching_indices) == 1:
            return matching_indices[0]
        else:
            return matching_indices
    else:
        return -1


def get_ical(url):
    import requests
    try:
        res = requests.get(url, timeout=(30.0, 30.0))
    except Exception as e:
        print(e.args)
        return ''
    else:
        res.raise_for_status()
        return res.text


def get_schedule_from_ical(df_lognote: pd.DataFrame, df_sig: pd.DataFrame) -> None:
    """ログノートの各行の時刻に対応するicalの予定を、BL2ical/BL3ical列に入れる"""
    import re
    from icalendar import Calendar
    JST = timezone(timedelta(hours=+9), 'JST')

    for n in range(len(df_sig)):
        print("label: ", str(df_sig.loc[n]['label']))
        icaldata = get_ical(str(df_sig.loc[n]['url']))
        cal = Calendar.from_ical(icaldata)

        for index, item in df_lognote.iterrows():
            for ev in cal.walk():
                if ev.name == 'VEVENT':
                    start_dt = ev.decoded("dtstart")
                    end_dt = ev.decoded("dtend")
                    try:
                        summary = ev['summary']
                    except Exception:
                        print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                    else:
                        try:
                            if (item['DT'].astimezone(JST) - start_dt).total_seconds() > 0 and (item['DT'].astimezone(JST) - end_dt).total_seconds() < 0:
                                tmp_summary = str(summary).replace(' ', '')
                                tmp_summary = re.sub(
                                    "（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す
                                tmp_summary = tmp_summary.rstrip('<br>')
                                tmp_summary = tmp_summary.replace("/30Hz", "")
                                tmp_summary = tmp_summary.replace("/60Hz", "")

                                if (df_sig.loc[n]['label'] == "BL2"):
                                    df_lognote.loc[index,
                                                   'BL2ical'] = tmp_summary
                                elif (df_sig.loc[n]['label'] == "BL3"):
                                    df_lognote.loc[index,
                                                   'BL3ical'] = tmp_summary
                                continue
                        except Exception:
                            pass


def run_summary(df: pd.DataFrame) -> None:
    import locale
    import webbrowser

    # ical用 Japanese
    locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
    df_sig = pd.read_excel("ical_SACLA.xlsx", sheet_name="sig")

    pd.options.display.max_colwidth = 2000
    pd.set_option('display.width', 1000)  # 少ないと改行されてしまうので増やす

    # 不要行削除
    # SRのログノート、日本語の文字列を含む行がdropできなかったが、前後の空白を削除することによって対処できた。
    df['C'] = df['C'].str.strip()
    drop_unneeded_rows(df, DROP_WORDS_SUMMARY)

    # ログノートA列の日付が00:00を過ぎても日付はそのままなので対処
    bf_itemDT = datetime(year=2000, month=1, day=1, hour=0, minute=0, second=0)
    for index, item in df.iterrows():
        if (type(item['DT']) is not datetime):
            print('DEBUG ~~~~~~~  type item[DT]=', type(item['DT']))
            continue
        try:
            if (item['DT'] - bf_itemDT).total_seconds() < 0:
                print(index, ' TIME INVERT: ログノートの時刻記載が間違ってる可能性があります。',  item['DT'], " - ", bf_itemDT, " = ", (
                    item['DT'] - bf_itemDT).total_seconds(), " NEW df.loc[index, 'DT'] = ",  df.loc[index, 'DT'])

            if item['DT'].hour == 0:
                df.loc[index, 'DT'] = item['DT'] + timedelta(days=1)

            bf_itemDT = df.loc[index, 'DT']
            df.loc[index, 'formatted_DT'] = df.loc[index,
                                                   'DT'].strftime('%Y/%#m/%#d %#H:%#M')
        except Exception as e:
            print(f"message:{e}")

    print("この処理には時間が掛かる~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    get_schedule_from_ical(df, df_sig)
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    styler = df.loc[:, ['formatted_DT', 'BL2ical', 'BL3ical', 'C']].style.map(
        lambda x: 'background-color: skyblue' if ('引渡' or '引き渡') in str(x) else '')
    styler = styler.map(lambda x: 'color: yellow' if ('切替') in str(x) else '')
    styler = styler.map(lambda x: 'color: pink' if ("加速器調整" in str(
        # 　なぜか or　が効かない
        x) or "BL-study" in str(x) or "BL調整" in str(x)) else '')
    styler = styler.map(lambda x: 'color: red' if ('終了') in str(x) else '')
    styler = styler.applymap(highlight_column_BL2, subset=[
                             'BL2ical'])  # BL2/BL3 ical列に色付け
    styler = styler.applymap(highlight_column_BL3, subset=[
                             'BL3ical'])  # BL2/BL3 ical列に色付け
    styler = styler.set_properties(**{'text-align': 'left'})  # 左寄せ

    for index, item in df.iterrows():
        try:
            if not ("加速器調整" in str(item['BL2ical']) or "BL-study" in str(item['BL2ical']) or "BL調整" in str(item['BL2ical'])) and not ("加速器調整" in str(item['BL3ical']) or "BL-study" in str(item['BL3ical']) or "BL調整" in str(item['BL3ical'])):  # ユーザー運転
                if "終了" in str(item['C']):
                    # 両方ユーザー運転中に終了したので行全体に色付け
                    styler = styler.apply(highlight_rows, rows_to_highlight=[
                                          index], color="lime", axis=1)
        except Exception as e:
            print(f"message:{e}")

    # セル内での改行をしたくないのでcssを噛ます
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
    webbrowser.open_new_tab('output.html')

    check_adjustment_time(df)


def check_adjustment_time(df: pd.DataFrame) -> None:
    """SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間(終了)がログノートに存在するか確認
    (get_schedule_from_ical(df)の前にこれをすると、icalのスケジュールがうまくdfに入らない)"""
    print("SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間終了がログノートに存在するか確認~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    first_dt = None
    for index, item in df.iterrows():  # ログノートの最初の日時を取得(ログノートが何月のなのかを確認するため)
        if (type(item['DT']) is datetime):
            first_dt = datetime(
                year=item['DT'].year, month=item['DT'].month, day=item['DT'].day, hour=0, minute=0, second=0)
            break
    if first_dt is None:
        print("ログノートに有効な日時がありません。")
        return

    with open(r"C:\me\unten\OperationSummary\dt_beg.txt", mode='r', encoding="UTF-8") as f:
        buff_dt_beg = f.read()
    with open(r"C:\me\unten\OperationSummary\dt_end.txt", mode='r', encoding="UTF-8") as f:
        buff_dt_end = f.read()
    dt_beg = datetime.strptime(buff_dt_beg, "%Y/%m/%d %H:%M")
    dt_end = datetime.strptime(buff_dt_end, "%Y/%m/%d %H:%M")
    df_kiroku = load_excel_to_dataframe(
        r"\\saclaopr18.spring8.or.jp\common\運転状況集計\最新\SACLA\SACLA運転集計記録.xlsm", "調整時間")
    if df_kiroku is None:
        return
    print("/    dt_beg=", dt_beg)
    start_row_index = 1  # 2行目以降
    column_index = 2  # 'end'列目  調整時間のstartにはチョッパーOFF時間になってる事があるので、ログノートの記載時間と合わないことがあるのでendで確認する。
    ans_line = -1
    for index, value in df_kiroku.iloc[start_row_index:, column_index].items():
        if value.month == first_dt.month and value >= dt_beg and value <= dt_end:  # 指定された月のログノートで、かつ、運転集計する期間内だけ確認
            result = check_datetime_existence_with_tolerance(
                df, 'DT', value, tolerance_minutes=1.0)  # ±1分の許容時間で確認
            if result == -1:
                ans_line = result
            elif isinstance(result, int):  # 1つのインデックス（int型）が返された場合
                ans_line = result
            else:  # 複数のインデックス（リスト型）が返された場合、とりあえず、最初のインデックスだけを使用。。。要改修
                ans_line = result[0]

            if ans_line != -1:
                matching_row = df.loc[ans_line, ['C']]
                try:
                    if not "引" in matching_row.to_string(header=False, index=False).replace('\n', ' ').strip():
                        print(
                            "🚨Warning  SACLA運転集計記録.xlsmのシート調整時間に記載されている調整「終了」時間( " + str(value) + " )がログノートに存在しますが、「引渡」と書かれていません。ログノートの内容：" + matching_row.to_string(header=False, index=False).replace('\n', '').strip())
                    else:
                        print(
                            "✅OK found 「引渡」 on Log note " + str(value))
                except Exception as e:
                    print(f"ERROR: {e}")
            else:
                print(
                    "🚨Warning SACLA運転集計記録.xlsmのシート調整時間に記載されている調整「終了」時間( " + str(value) + " )がログノートに存在しません。調整理由:" + str(df_kiroku.iloc[index, 3]))
        else:
            print(
                f"index: {index}, value: {value} is out of range.運転集計する期間内ではないのでスキップします。")
    print("dt_end=", dt_end, "  /")
    print("SACLA運転集計記録.xlsmのシート調整時間を読み込んで、調整時間終了がログノートに存在するか確認　が終了しました。~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")


# ============================================================================================
# main
# ============================================================================================

def main() -> None:
    parser = argparse.ArgumentParser(description="ログノート(xlsm)のXMLを解析する")
    parser.add_argument('-m', '--mode', choices=['search', 'summary'], default='search',
                        help="search: ログノート検索(既定) / summary: 運転集計(ical付きHTML出力+調整時間の確認)")
    parser.add_argument('shared_strings', help="sharedStrings.xml のパス")
    parser.add_argument('sheet', help="sheet1.xml のパス")
    args = parser.parse_args()

    print(f"============ ここから excelgrep_by_XMLparse.py (mode={args.mode}) ============")
    print("version", pd.__version__)
    pd.set_option('display.max_rows', None)
    pd.options.display.colheader_justify = 'left'  # 列名表示の右寄せ

    print('sys.stdout.encoding:', sys.stdout.encoding)
    # default でutf-8なのに、これをしないと文字化けする。なぜ？？
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    print('sys.stdout.encoding:', sys.stdout.encoding)

    print("Arg[sharedStrings.xml]:\t", args.shared_strings)
    print("Arg[sheet1.xml]:\t", args.sheet)

    sslist = load_shared_strings(args.shared_strings)
    if len(sslist) == 0:
        print("sharedStrings.xml のsiタグの数がlenght of maxsslit = 0 です。終了します")
        sys.exit()
    print("lenght of maxsslit = ", len(sslist))

    df = load_sheet(args.sheet, sslist)

    if args.mode == 'summary':
        run_summary(df)
    else:
        run_search(df)


if __name__ == '__main__':
    main()
