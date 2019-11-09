from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
import json
from flask import *

#ServiceAccountCredentials：Googleの各サービスへアクセスできるservice変数を生成します。
from oauth2client.service_account import ServiceAccountCredentials

#2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']

#認証情報設定
#ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
credentials = ServiceAccountCredentials.from_json_keyfile_name('python-google-sheet.json', scope)

#OAuth2の資格情報を使用してGoogle APIにログインします。
gc = gspread.authorize(credentials)

#共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
SPREADSHEET_KEY2 = "1YHbhRnTtiON1nQMyo5O8kk8fDwiXcOJQihYxcO0-cVc"
worksheet2 = gc.open_by_key(SPREADSHEET_KEY2).worksheet("GUiDEE利用状況レポート")
worksheet3 = gc.open_by_key(SPREADSHEET_KEY2).worksheet("週次")
mentor_names = set(worksheet3.col_values(3))#「週次」の2列目からメンターの名前をsetで拾う
num_address_worksheet = gc.open_by_key(SPREADSHEET_KEY2).worksheet("メアド一覧")#「離脱者週次報告」の「メアド一覧」を開く→これをCAOのファイルに変えたい

name_id_dic = {k:v for k,v in zip(num_address_worksheet.col_values(2),num_address_worksheet.col_values(4))}#「離脱者週次報告」の「メアド一覧」を開いて,2行目と4行目から名前、アドレスを取得
id_name_dic = {v:k  for k, v in name_id_dic.items()}#上記のname_id_dicの{name:address}を{address:name}に変える

pair_id_list = []
exception_pair_id_list = []

#メンターのカレンダーのみ見る
mentor_id_list = []
for name in mentor_names:#「離脱者週次報告」の「週次」から取得したmentorの名前のsetについて繰り返し処理
    if name != "メンター" and name!= "":#空欄と「メンター」の文字を除く
        
        mentor_id_list.append(name_id_dic[name])#mentorの名前をmentorのアドレスに変更(離脱者週次報告の「メアド一覧」)し、mentor_id_listに格納

#ペアのアドレスを取得
test_pairs= zip(worksheet3.col_values(4),worksheet3.col_values(7))#離脱者週次報告「週次」から、ペアのアドレスをzipで取得
test_pair_id_list =[]
for test_pair_id in test_pairs:
    test_pair_id_list.append(list(test_pair_id))#検証対象ペアを（メンター、メンティー）の順でリストを取得

#上記と同じように、順番逆のペアリストを作る
inverse_test_pairs= zip(worksheet3.col_values(7),worksheet3.col_values(4))
inverse_test_pair_id_list = []

for inverse_test_pair_id in inverse_test_pairs:
    inverse_test_pair_id_list.append(list(inverse_test_pair_id))#検証対象ペアを（メンティー、メンター）の順で取得
now = datetime.datetime.utcnow().isoformat() + 'Z'

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

def tra_Z_JST_isoformat(date):
    date = (datetime.datetime.strptime(date, "%Y-%m-%d %H:%M")-datetime.timedelta(hours=9)).isoformat()+"Z"
    return date

def tra_Z_JST_datetime(date2):
    date2 = datetime.datetime.strptime(date2, "%Y-%m-%d　%H:%M")-datetime.timedelta(hours=9)
    return date2

app = Flask(__name__)
@app.route("/", methods=["GET", "POST"])
def check1on1():#カレンダー上で1on1があるペアの中で、検証対象であるペアのリストを返す 。例外のリストprintされる。#期間のみ指定可能
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    """引数はどちらも、isoformatである必要がある"""
    if request.method == "GET":
        return """
        <p>いつ以降の1on1を把握したいですか？</p>
        <p>記入例）2019-11-01 10:00</p>
        <form action="/" method="POST">
        <input name="checkbefore" value = "2019-11-01 10:00"></input>
     
        <p>いつまでの1on1を把握したいですか？</p>
        <p>記入例）2019-11-01 19:00</p>
        <input name="checkafter" value = "2019-11-01 19:00"></input>
        <input type="submit" value="check">
        </form>
     """
    else:
        try:
            check_timeMin = tra_Z_JST_isoformat(request.form["checkbefore"])
            check_timeMax = tra_Z_JST_isoformat(request.form["checkafter"])
            creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
            if os.path.exists('token.pickle'):
                with open('token.pickle','rb') as token:
                    creds = pickle.load(token)
            # If there are no (valid) credentials available, let the user log in.
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                    creds = flow.run_local_server(port=0)
                # Save the credentials for the next run
                with open('token.pickle', 'wb') as token:
                    pickle.dump(creds, token)
                    
            service = build('calendar','v3',credentials=creds)

        # Call the Calendar API

            for ids in mentor_id_list:#検証シートの「新検証ペア」に記載されている、メンターのアドレスについて繰り返し処理
                try :
                    events_result = service.events().list(calendarId=ids, timeMin=check_timeMin , singleEvents=True,timeMax = check_timeMax,orderBy='startTime').execute()
                except :
                    continue
            
                events = events_result.get('items', [])
        
                if not events:
                    print(id_name_dic[ids],'指定された期間に予定はありません')
            
                for event in events:
                    pair=[]
                    if "summary" in event.keys() and "attendees" in event.keys():#attendeesがいなくなる場合は？attendeesがいなくても1on1をする場合があるか？
                        if ("振り返り" in event["summary"]) or ("目標設定" in event["summary"]) or ( "評価面談" in event["summary"]) or  ("1on1" in event["summary"]) or ("パフォーマンスレビュー" in event["summary"]) or ("１on１" in event["summary"]) or ("１on1"in event["summary"]) or ("１on1"in event["summary"]) or ("1 on 1" in event["summary"]) or ("1 on1" in event["summary"]) or ("1on 1" in event["summary"]) :
                            start = event['start'].get('dateTime', event['start'].get('date'))
                    
                            for attendee in event["attendees"]:#1on1とタイトルにあり、attendeeがいるイベントについて、各attendeeについて繰り返し処理
                                if attendee["email"] in id_name_dic.keys():#「東京本社 Daisy」とかを省くために、メール一覧に名前があるかを用いて制約を課す
                                    pair.append(attendee["email"])#1on1があるペアのアドレスのリストを作成
        
                            if (pair in test_pair_id_list) or (pair in inverse_test_pair_id_list):
                                if (pair not in pair_id_list) and (pair.reverse() not in pair_id_list):
                                    pair_id_list.append(pair)
                        
                        #print(start, event["summary"])
                        #print(pair)
                        #print("-------------------")
                        
                            elif len(pair) >2:
                        #print("例外",id_name_dic[ids],start,event["summary"])
                                if pair not in exception_pair_id_list:
                                    exception_pair_id_list.append([id_name_dic[ids],start,event["summary"]])


    #if exception_pair_id_list:
     #   print("1on1の疑いあり:",exception_pair_id_list)
      #  print("---------------------------------------")

            pair_name_list=[]
            for pair_ids in  pair_id_list:
                if '#N/A' not in pair_ids:
                        pair_name_list.append([id_name_dic[pair_ids[0]],id_name_dic[pair_ids[1]]])

            return """1on1がある人はこの人達です！{}
        <p>1on1の疑いがある人たちは以下の人達です</p>{}""".format(pair_name_list,exception_pair_id_list)
        
        except ValueError:
            return"""値が不正でした！もう一度お願いします！
                <p>いつ以降の1on1を把握したいですか？</p>
                <form action="/" method="POST">
                <input name="checkbefore"></input>
                </form>
            
                <p>いつまでの1on1を把握したいですか？</p>
                <form action="/" method="POST">
                <input name="checkafter"></input>
                </form>"""
        

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=8888, threaded=True)
    



