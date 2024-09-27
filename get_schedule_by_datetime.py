#   運転集計用にテスト的に作成
#	python get_schedule_by_datetime.py


# Formatter     Shift+Alt+F

import requests
from requests.exceptions import Timeout
import re
import pandas as pd
import sys
from icalendar import Calendar, Event
import datetime
import time

""" Japanese"""
import locale
dt = datetime.datetime(2018, 1, 1)
print(locale.getlocale(locale.LC_TIME))
print(dt.strftime('%A, %a, %B, %b'))
locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
print(locale.getlocale(locale.LC_TIME))
#print(dt.strftime('%A, %a, %B, %b'))
"""------"""



config_file_sig = "ical_TEST.xlsx"
df_sig = pd.read_excel(config_file_sig, sheet_name="sig")
# print(df_sig)


"""  -------------------------------------------------------------------------------------  """


def get_ical(url):

    # print(url)
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



JST = datetime.timezone(datetime.timedelta(hours=+9), 'JST')

now = datetime.datetime.now()
dt1 = now + datetime.timedelta(days=-1)
dt2 = now + datetime.timedelta(days=23)

list_dt = []
list_dt.append(now)
list_dt.append(dt1)
list_dt.append(dt2)


while True:


    for n, s in enumerate(sig, 0):
        print("label: ",str(df_sig.loc[n]['label']))
        s.icaldata = get_ical(str(df_sig.loc[n]['url']))
        # print(s.icaldata)
        cal = Calendar.from_ical(s.icaldata)
#        m = 0


        for searchdt in list_dt:
#            print(searchdt)

            for ev in cal.walk():

#                print('ev.decoded("dtstart")):  ',  str(ev.decoded("dtstart")))
                """
                try:
                    print('ev:', ev.decoded("dtstart"))
                    start_dt_datetime = datetime.datetime.strptime(
                        str(ev.decoded("dtstart")), '%Y-%m-%d %H:%M:%S+09:00')
                except Exception as e:
                    print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                    continue
                else:
                    if (start_dt_datetime - now).days < -60:
                        continue
    #		        else:
    #			        print('(start_dt_datetime - now).days	' + str((start_dt_datetime - now).days))
                """

                if ev.name == 'VEVENT':
                    start_dt = ev.decoded("dtstart")
                    end_dt = ev.decoded("dtend")
                    try:
                        summary = ev['summary']
                    except Exception as e:
                        print('Exception!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!	')
                    else:
                        d = {}
                        d["Task"] = str(df_sig.loc[n]['label'])
                        d["Start"] = start_dt
                        d["Finish"] = end_dt

                        tmp_summary = str(summary).replace(' ', '')


#                        print('start_dt	' + str(start_dt))
#                        print('end_dt	' + str(end_dt))

                        tmp_summary = re.sub("（.+?）", "", tmp_summary)  # カッコで囲まれた部分を消す

                        tmp_summary = tmp_summary.rstrip('<br>')
                        tmp_summary = tmp_summary.replace("/30Hz", "")
                        tmp_summary = tmp_summary.replace("/60Hz", "")
                        tmp_summary = tmp_summary.replace("SEED", "<i>SEED</i>")

                        if (searchdt.astimezone(JST) - start_dt).total_seconds() > 0 and (searchdt.astimezone(JST) - end_dt).total_seconds() < 0:
                            print("searchdt= ",searchdt,"   :    ", tmp_summary)

#                        print(str(start_dt) + " ~ " + str(end_dt) + " :	" + tmp_summary)

    break

