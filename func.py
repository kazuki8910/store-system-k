####################
# 
# モジュール
# 
####################

import json

import gspread
from oauth2client.service_account import ServiceAccountCredentials



####################
# 
# 関数
# 
####################

####################
# スプシにアクセスする関数
####################
def connect_gspread(key):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name("config/sheet_api.json", scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_key(key)
    return worksheet