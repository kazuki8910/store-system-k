# import csv

# with open('data/reservations.csv', encoding='utf8') as f_in:
#     with open('data/sjis.csv', 'w', encoding='cp932') as f_out:
#         f_out.write(f_in.read())


import gspread
import json
from oauth2client.service_account import ServiceAccountCredentials

# スプシにアクセスする関数
def connect_gspread(jsonf,key):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonf, scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_key(key)
    return worksheet

# 基幹シートにアクセス
jsonf = "config/sheet_api.json"
sheet_key_kikan = "1rlaJptyRr-MkbLejSakF9zK1vdLoP6pMAhYcoj6p7dI"
ws_kikan = connect_gspread(jsonf,sheet_key_kikan) # 基幹シート


####################
# Sheet関連の変数
####################
sheets_kikan = ws_kikan.worksheet("2022年1月") # 基幹シートのワークシート一覧

####################
# Sheet関連の関数
####################

# シート名が存在するか検証
def exsit_sheet(sheets, sheet_name):
    for sheet in sheets:
        if sheet.title == sheet_name:
            return False
    return True


sheets_kikan.update_acell('A1', "aaaaa")