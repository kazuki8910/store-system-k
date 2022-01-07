# 関数を定義



####################
# 
# モジュール
# 
####################

# 一般
import time
import os
from pathlib import Path
import csv

# スプシ関連
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# セレニウム関連
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_binary





####################
# 
# スプシ関連
# 
####################

# スプシにアクセスする関数
def connect_gspread(key):
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name("config/sheet_api.json", scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_key(key)
    return worksheet


# シート名が存在するか検証
def exsit_sheet(sheets, sheet_name):
    for sheet in sheets:
        if sheet.title == sheet_name:
            return False
    return True


# CSVをシートにアップロード
def upload_csv(workbook, sheet_name, file_path):
    workbook.values_update(
        sheet_name,
        params={
            'valueInputOption': 'USER_ENTERED'
        },
        body={'values': list(csv.reader(open(file_path, encoding='cp932')))}
    )





####################
# 
# セレニウム関連
# 
####################

# クローム起動
def start_chrome():
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

    # クローム起動
    driver = webdriver.Chrome(options=options)
    driver.set_window_size(1280, 720)
    time.sleep(1)
    print("クローム起動")
    return driver


# 要素に入力する関数
def input_for_elm(elm, text):
    elm.clear() #元から入っているテキスト削除
    elm.send_keys(text) #指定のテキスト入力
    time.sleep(1)


# IDから要素に入力
def input_text_by_id(driver, id, text):
    input_tag = driver.find_element_by_id(id) #要素取得
    input_for_elm(input_tag, text) #要素に入力


# Xpathから要素クリック
def click_by_xpath(driver, xpath):
    elm = driver.find_element_by_xpath(xpath) # 要素取得
    elm.click()
    time.sleep(1)





####################
# 
# ファイル操作
# 
####################

# ファイル削除
def delete_file(path):
    if(os.path.isfile(path)):
        os.remove(path)
