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
import shutil

# GoogleAPI関連
import gspread
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
from oauth2client.service_account import ServiceAccountCredentials

# セレニウム関連
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
import chromedriver_binary

# 自作
import conf





####################
# 
# GoogleAPI関連
# 
####################

# サービスアカウントキー読み込み
def get_credentials():
    credentials = ServiceAccountCredentials.from_json_keyfile_name(conf.google_api_filepath, conf.google_api_scope)
    return credentials


# pydrive認証
def auth_pydrive():
    gauth = GoogleAuth()
    gauth.credentials = get_credentials()
    drive = GoogleDrive(gauth)
    return drive


# スプシにアクセスする関数
def connect_gspread(key):
    gc = gspread.authorize(get_credentials())
    workbook = gc.open_by_key(key)
    return workbook


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


# 今月の分析シートがない場合作成
def creat_new_ana_sheet():
    # pydrive認証
    drive = auth_pydrive()

    # ファイル名のリスト取得
    query = "'{}' in parents and trashed=false".format(conf.drive_folder_key)
    file_names = []
    for file_list in drive.ListFile({'q': query}):
        for file in file_list:
            file_names.append(file["title"])

    # 今月の分析シートがない場合実行
    if(not conf.book_name_ana in file_names):

        # 新規ファイル作成
        f = drive.CreateFile({
            'title': conf.book_name_ana,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            "parents": [{"id": conf.drive_folder_key}]})
        f.Upload()
        file_key = f["id"]

        # キーを設定ファイルに保存
        workbook = connect_gspread(conf.sheet_key_conf)
        worksheet = workbook.worksheet(conf.sheet_name_conf)
        len_wb = len(worksheet.get_all_values())
        worksheet.update_cell(len_wb+1, 1, conf.book_name_ana)
        worksheet.update_cell(len_wb+1, 2, file_key)

        print("新規分析シート作成")

    else:
        print("分析シートはすでに存在します")


# 設定シートから分析シートのキーを取得
def get_ana_key():
    # 分析シートがない場合作成
    creat_new_ana_sheet()

    # シートキー取得
    wb_conf = connect_gspread(conf.sheet_key_conf)
    ws_conf = wb_conf.worksheet(conf.sheet_name_conf)
    rows_ws_conf = len(ws_conf.get_all_values())
    return ws_conf.cell(rows_ws_conf, 2).value






####################
# 
# セレニウム関連
# 
####################

# クローム起動
def start_chrome(headless=True):
    # オプションを格納
    options = webdriver.ChromeOptions()

    # ヘッドレスモード
    if(headless):
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

    
# Xpathから要素に入力
def input_text_by_xpath(driver, xpath, text):
    input_tag = driver.find_element_by_xpath(xpath) #要素取得
    input_for_elm(input_tag, text) #要素に入力


# Xpathから要素クリック
def click_by_xpath(driver, xpath, delay=1):
    elm = driver.find_element_by_xpath(xpath) # 要素取得
    elm.click()
    time.sleep(delay)


# Xpathから要素にホバー
def hover(driver, xpath):
    elm = driver.find_element_by_xpath(xpath)
    actions = ActionChains(driver)
    actions.move_to_element(elm).perform()



####################
# 
# ファイル操作
# 
####################

# ファイル削除
def delete_file(path):
    if(os.path.isfile(path)):
        os.remove(path)


# データフォルダ内のファイル全削除
def delete_data_files():
    dir_name_data = "data"
    shutil.rmtree(dir_name_data)
    os.mkdir(dir_name_data)