#!/usr/bin/env python
# coding: utf-8

# In[1]:


####################
# 
# モジュール
# 
####################

# セレニウム関連
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_binary

# 基本モジュール
import time
from pathlib import Path
import csv
import os
import datetime


# In[2]:


####################
# 
# 変数
# 
####################

# 基幹システム関連
url_kikan  = "https://stlassh.com/admin/" #URL

mail_kikan = "info@stlassh.com" #メールアドレス
pass_kikan = "2sYsLpebQ3sx"     #PASS
input_mail_id_kikan = 'user_email'    #メールアドレスのインプットタグのID
input_pass_id_kikan = 'user_password' #パスワードのインプットタグのID
xpath_login_btn = '//*[@id="new_user"]/div[3]/input' #ログインボタンのXPath

xpath_reserve_elm = '/html/body/main/div/div[1]/a'    # 予約/問合せの要素
xpath_csv_btn     ='/html/body/main/div/form/input[3]'# CSV出力ボタン

file_path_kikan = "data/reservations.csv" # CSVの保存先パス

# 日付関連
this_time  = datetime.datetime.now() # 今の時間
this_year  = this_time.year  # 今の年
this_month = this_time.month # 今の月
sheet_name_this_month = str(this_year) + "年" + str(this_month) + "月" # 今月のシート名


# In[3]:


####################
# 
# 関数
# 
####################

# 要素に入力する関数
def input_for_elm(elm, text):
    elm.clear() #元から入っているテキスト削除
    elm.send_keys(text) #指定のテキスト入力
    time.sleep(1)

# IDから要素に入力
def input_text_by_id(id, text):
    input_tag = driver.find_element_by_id(id) #要素取得
    input_for_elm(input_tag, text) #要素に入力
    
# Xpathから要素に入力
def input_text_by_xpath(xpath, text):
    input_tag = driver.find_element_by_xpath(xpath) #要素取得
    input_for_elm(input_tag, text) #要素に入力
    

# Xpathから要素クリック
def click_by_xpath(xpath):
    elm = driver.find_element_by_xpath(xpath) # 要素取得
    elm.click()
    time.sleep(1)


# ファイル削除
def delete_file(path):
    if(os.path.isfile(path)):
        os.remove(path)


# In[4]:


####################
# 
# Sheet API設定
# 
####################

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
sheets_kikan = ws_kikan.worksheets() # 基幹シートのワークシート一覧

####################
# Sheet関連の関数
####################

# シート名が存在するか検証
def exsit_sheet(sheets, sheet_name):
    for sheet in sheets:
        if sheet.title == sheet_name:
            return False
    return True


# In[5]:


####################
# 
# メイン動作
# 
####################


# In[6]:


####################
# クロームの起動設定
####################

# オプションを格納
options = webdriver.ChromeOptions()

# ヘッドレスモード
options.add_argument('--headless')

# ファイルの保存先指定
dldir_name = 'data'  # 保存先フォルダ名
dldir_path = Path(dldir_name)
dldir_path.mkdir(exist_ok=True)  # 存在していてもOKとする（エラーで止めない）
download_dir = str(dldir_path.resolve())  # 絶対パス
options.add_experimental_option("prefs", {
    "download.default_directory": download_dir,
    "plugins.always_open_pdf_externally": True
})


# In[7]:


# クローム起動
driver = webdriver.Chrome(options=options)
driver.set_window_size(1280, 720)
time.sleep(1)


# In[8]:


# 基幹システムログイン
driver.get(url_kikan) # 基幹システムにアクセス
time.sleep(1)

input_text_by_id(input_mail_id_kikan, mail_kikan) #メールアドレス入力
input_text_by_id(input_pass_id_kikan, pass_kikan) #パスワード入力
click_by_xpath(xpath_login_btn) # ログインボタンクリック


# In[9]:


# CSVダウンロード
click_by_xpath(xpath_reserve_elm) # 予約/問合せをクリック
delete_file(file_path_kikan) # 既存のCSV削除
click_by_xpath(xpath_csv_btn) # CSVダウンロードクリック
time.sleep(5)
driver.quit() # ブラウザ終了


# In[10]:


# シートが存在しなかったら今月のシート追加
if(exsit_sheet(sheets_kikan, sheet_name_this_month)):
    ws_kikan.add_worksheet(title=sheet_name_this_month, rows=1, cols=1)


# In[11]:


# CSVをシートにアップロード
ws_kikan.values_update(
    sheet_name_this_month + '!A1',
    params={
        'valueInputOption': 'USER_ENTERED'
    },
    body={'values': list(csv.reader(open(file_path_kikan)))}
)


# In[ ]:




