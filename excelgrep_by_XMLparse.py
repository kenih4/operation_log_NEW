import xml.etree.ElementTree as ET
import argparse
import functools
import io
import os
import re
import subprocess
import sys
import traceback
import zipfile
import numpy as np
import pandas as pd

from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Union, List

#
# 使い方:
#   python excelgrep_by_XMLparse.py [--mode search|summary] [-k 検索ワード] file1.xlsm [file2.xlsm ...]
#
#   --mode search  (既定) ログノート検索モード。DT,C列から -k の検索ワード(grep -iE)に一致する行を色付きで出力する
#   --mode summary        運転集計モード。icalカレンダーを付与してHTML出力し、
#                         SACLA運転集計記録.xlsmの調整時間がログノートに記載されているか確認する
#   --mode log            ログ出力モード。運転集計モードと同じ前処理(不要行削除・日付跨ぎ補正)をして
#                         日時とログ内容を「D:\LOGNOTE\output\エクセルファイル名.txt」にテキスト出力する(ical付与・突合チェックはしない)
#
# xlsmはzipとして直接読む(展開しない)。複数ファイルを渡しても、Pythonの起動とpandasのimportは1回だけ。
#
# 通常は excelgrep_by_XMLparse.sh から呼ばれる(-k=検索ワード があれば search、なければ summary)
#
# Formatter     Shift+Alt+F

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

LOG_OUTPUT_DIR = r'D:\LOGNOTE\output'  # ログ出力モードの出力先フォルダ

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

def load_shared_strings(source) -> list:
    """sharedStrings.xml のsiタグの部分(最初のtだけ)を配列に格納"""
    sslist = []
    root = ET.parse(source).getroot()
    for ssl in root:
        for child in ssl.iter():
            # 特定要素(si)の抽出
            if child.tag == NS + 'si':
                for child2 in child.iter():
                    if child2.tag == NS + 't':
                        sslist.append(child2.text)
                        break
    return sslist


def load_sheet(source, sslist: list) -> pd.DataFrame:
    """sheet1.xml のA,B,C列をピックアップしてDataFrameにする(行は貯めて最後に1回だけDataFrame化)"""
    maxsslit = len(sslist)
    root = ET.parse(source).getroot()

    # A,B,Cはその行にセルが無ければ前の行の値を引き継ぐ(A、B列は日時なのでクリアしない)。C列(内容部分)だけ行ごとにクリア
    A = B = C = DT = np.nan
    rows = []
    for row in root.iter(NS + 'row'):  # 特定要素(row)の抽出
        touched = False
        for cell in row.iter(NS + 'c'):  # 特定要素(c)の抽出
            ref = cell.attrib["r"]
            is_a, is_b, is_c = not ref.find('A'), not ref.find('B'), not ref.find('C')
            if not (is_a or is_b or is_c):
                continue
            touched = True
            for v in cell.iter(NS + 'v'):
                if is_a:
                    try:
                        A = int(v.text)
                    except Exception:
                        A = v.text
                if is_b:
                    try:
                        B = float(v.text)
                    except Exception:
                        B = v.text
                if is_c:
                    try:
                        if int(v.text) < maxsslit:
                            VAL = sslist[int(v.text)]
                        else:
                            VAL = v.text
                    except Exception:
                        VAL = v.text
                    C = (VAL if VAL is not None else '').ljust(500)  # 左寄せ

        if touched:
            try:
                # B列(時間)がない場合、例外が発生するので、その時は00:00にするしかない
                DT = datetime(1899, 12, 30) + timedelta(A + B)
            except Exception:
                try:
                    DT = datetime(1899, 12, 30) + timedelta(A)
                except Exception:
                    DT = 0
        rows.append({'A': A, 'B': B, 'C': C, 'DT': DT})
        C = "-"

    df = pd.DataFrame(rows, columns=COLUMNS, dtype=object)
    try:
        # errors='coerce'だと変換できない値はNaTになる
        df['A'] = pd.to_timedelta(
            df['A'], unit='D', errors="coerce")+pd.to_datetime("1899/12/30")
    except Exception:
        print('Error')
    return df


def load_xlsm(path: str):
    """xlsmをzipのまま読み、(sharedStringsの配列, ログノートのDataFrame)を返す"""
    with zipfile.ZipFile(path) as zf:
        with zf.open('xl/sharedStrings.xml') as f:
            sslist = load_shared_strings(f)
        with zf.open('xl/worksheets/sheet1.xml') as f:
            df = load_sheet(f, sslist)
    return sslist, df


def drop_unneeded_rows(df: pd.DataFrame, words: list) -> None:
    """C列が「-」の行と、wordsのどれかを含む行を削除する(dfを直接変更。判定は1回でまとめて行う)"""
    # 大文字小文字を無視するにはcase=False、NaNを無視するにはna=False。wordsはエスケープして単純な文字列として検索
    pattern = '|'.join(re.escape(w) for w in words)
    mask = (df['C'] == "-") | df['C'].str.contains(
        pattern, case=False, na=False)
    df.drop(index=df.index[mask], inplace=True)


# ============================================================================================
# 検索モード (--mode search)
# ============================================================================================

def grep_text(text: str, pattern: str) -> None:
    """textをgrepにかけて色付きで出力する(grepはGit Bashのものを使う)"""
    sys.stdout.flush()
    env = dict(os.environ, GREP_COLOR='0;33')
    try:
        subprocess.run(['grep', '-a', '--color', '-n', '-A', '0', '-iE', pattern],
                       input=text.encode('utf-8'), env=env)
    except FileNotFoundError:
        print("❌ grep が見つかりません。Git Bashから実行してください。")


def run_search(df: pd.DataFrame, keyword: str | None) -> None:
    drop_unneeded_rows(df, DROP_WORDS_SEARCH)

    # print(df)だとC列が約50文字で切り詰められ、後半の語が検索にヒットしないので、1行に全文を出す
    text = '\n'.join(
        f"{index:<6}{dt}  {str(c).replace(chr(10), ' ').rstrip()}"
        for index, dt, c in zip(df.index, df['DT'], df['C']))
    if keyword:
        grep_text(text, keyword)
    else:
        print(text)


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


def load_ical_events(icaldata: str) -> list:
    """icalからイベント(開始, 終了, 整形済みの予定名)のリストを返す。開始・終了はtimezone付きdatetimeのみ(終日予定は除外)"""
    import re
    from icalendar import Calendar

    events = []
    for ev in Calendar.from_ical(icaldata).walk():
        if ev.name != 'VEVENT':
            continue
        try:
            start_dt = ev.decoded("dtstart")
            end_dt = ev.decoded("dtend")
            summary = ev['summary']
        except Exception:
            print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
            continue
        if not (isinstance(start_dt, datetime) and isinstance(end_dt, datetime)
                and start_dt.tzinfo is not None and end_dt.tzinfo is not None):
            continue
        tmp_summary = str(summary).replace(' ', '')
        tmp_summary = re.sub("（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す
        tmp_summary = tmp_summary.rstrip('<br>')
        tmp_summary = tmp_summary.replace("/30Hz", "")
        tmp_summary = tmp_summary.replace("/60Hz", "")
        events.append((start_dt.timestamp(), end_dt.timestamp(), tmp_summary))
    return events


_ICAL_EVENTS_CACHE = {}  # url -> events (複数ファイルを処理しても同じicalは1回だけ取得する)


def get_ical_events(urls: list) -> list:
    """各urlのイベントのリストを返す。未取得のicalは並列に取得する"""
    from concurrent.futures import ThreadPoolExecutor

    todo = [u for u in dict.fromkeys(urls) if u not in _ICAL_EVENTS_CACHE]
    if todo:
        with ThreadPoolExecutor(max_workers=len(todo)) as ex:
            texts = list(ex.map(get_ical, todo))
        for url, text in zip(todo, texts):
            _ICAL_EVENTS_CACHE[url] = load_ical_events(text)
    return [_ICAL_EVENTS_CACHE[u] for u in urls]


def get_schedule_from_ical(df_lognote: pd.DataFrame, df_sig: pd.DataFrame) -> None:
    """ログノートの各行の時刻に対応するicalの予定を、BL2ical/BL3ical列に入れる
    (イベントを先に1回だけ展開し、全ログ行をnumpyでまとめて判定する。複数の予定に重なる場合は後のイベントが優先)"""
    JST = timezone(timedelta(hours=+9), 'JST')

    # 各ログ行のDTをJSTのtimestamp(秒)にする。DTが日時でない行はNaN(どのイベントにも一致しない)
    dt_ts = np.full(len(df_lognote), np.nan)
    for i, dt in enumerate(df_lognote['DT']):
        if isinstance(dt, datetime):
            try:
                dt_ts[i] = dt.astimezone(JST).timestamp()
            except (OSError, OverflowError, ValueError):  # 1970年より前などの異常な日時は対象外
                pass

    labels = [df_sig.loc[n]['label'] for n in range(len(df_sig))]
    for label in labels:
        print("label: ", str(label))
    events_list = get_ical_events(
        [str(df_sig.loc[n]['url']) for n in range(len(df_sig))])

    for label, events in zip(labels, events_list):
        if label not in ("BL2", "BL3"):
            continue
        col = label + 'ical'
        result = np.full(len(df_lognote), None, dtype=object)
        for start_ts, end_ts, summary in events:
            result[(dt_ts > start_ts) & (dt_ts < end_ts)] = summary
        matched = result != None  # noqa: E711  (object配列の要素ごとの比較)
        if matched.any():
            values = df_lognote[col].to_numpy(dtype=object).copy()
            values[matched] = result[matched]
            df_lognote[col] = values


@functools.cache
def setup_summary() -> pd.DataFrame:
    """運転集計モードの準備(ロケール設定とical_SACLA.xlsxの読込)。複数ファイルでも1回だけ"""
    import locale
    try:
        locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
    except locale.Error as e:
        print(f"警告: ロケールを設定できませんでした: {e}")
    return pd.read_excel("ical_SACLA.xlsx", sheet_name="sig")


def prepare_lognote(df: pd.DataFrame) -> None:
    """運転集計モード/ログ出力モード共通の前処理(不要行の削除と、日付跨ぎの補正+formatted_DTの作成)"""
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


def run_log(df: pd.DataFrame) -> str:
    """ログ出力モード: 運転集計モードの前処理だけを行い、日時とログ内容をテキストにして返す
    (色付け/HTML出力、icalの付与、運転集計記録との突合は行わない)"""
    prepare_lognote(df)
    return '\n'.join(
        f"{'' if pd.isna(formatted) else formatted}\t{str(c).replace(chr(10), ' ')}"
        for formatted, c in zip(df['formatted_DT'], df['C']) if not pd.isna(c))  # C列が空の行は出力しない


def run_summary(df: pd.DataFrame) -> None:
    import webbrowser

    df_sig = setup_summary()

    pd.options.display.max_colwidth = 2000
    pd.set_option('display.width', 1000)  # 少ないと改行されてしまうので増やす

    prepare_lognote(df)

    print("この処理には時間が掛かる~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    get_schedule_from_ical(df, df_sig)
    print("~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")

    styler = df.loc[:, ['formatted_DT', 'BL2ical', 'BL3ical', 'C']].style.map(
        lambda x: 'background-color: skyblue' if any(w in str(x) for w in ('引渡', '引き渡')) else '')
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
        if not isinstance(value, datetime):  # 空欄(NaT)や日時以外はスキップ
            print(f"index: {index}, value: {value} は日時ではないのでスキップします。")
            continue
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
    parser = argparse.ArgumentParser(description="ログノート(xlsm)を解析する")
    parser.add_argument('-m', '--mode', choices=['search', 'summary', 'log'], default='search',
                        help="search: ログノート検索(既定) / summary: 運転集計(ical付きHTML出力+調整時間の確認) / "
                             f"log: ログ出力({LOG_OUTPUT_DIR}にエクセルファイル名.txtでテキスト出力)")
    parser.add_argument('-k', '--keyword', help="searchモードの検索ワード(grep -iE の正規表現)")
    parser.add_argument('files', nargs='+', help="xlsm/xlsファイル(複数可)")
    args = parser.parse_args()

    # default でutf-8なのに、これをしないと文字化けする。なぜ？？
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    pd.set_option('display.max_rows', None)
    pd.options.display.colheader_justify = 'left'  # 列名表示の右寄せ

    for path in args.files:
        print(f"📘 File: {path}__________________________________________________________________________")
        try:
            sslist, df = load_xlsm(path)
            if len(sslist) == 0:
                print("sharedStrings.xml のsiタグの数が0です。スキップします")
                continue
            if args.mode == 'summary':
                run_summary(df)
            elif args.mode == 'log':
                out_dir = Path(LOG_OUTPUT_DIR)
                out_dir.mkdir(parents=True, exist_ok=True)
                out_path = out_dir / (Path(path).stem + '.txt')  # エクセルファイル名.txt
                with open(out_path, 'w', encoding='utf-8') as f:
                    f.write(run_log(df) + '\n')
                print(f"📝 ログを {out_path} に出力しました")
            else:
                run_search(df, args.keyword)
        except zipfile.BadZipFile:
            print("❌ ZIPファイルは異常です。指定されたファイルはZIP形式ではないか、壊れている可能性があります。")
        except KeyError as e:
            print(f"❌ xlsm内に必要なXMLが見つかりません: {e}")
        except Exception:
            traceback.print_exc()


if __name__ == '__main__':
    main()
