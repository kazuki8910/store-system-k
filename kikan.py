#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 基幹データ反映


####################
# 
# モジュール
# 
####################

# 基本モジュール
import time

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")


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

print("変数の定義完了")


# In[3]:


####################
# 
# 基幹シートにアクセス
# 
####################

ws_kikan = func.connect_gspread(conf.sheet_key_kikan) # 基幹シート
sheets_kikan = ws_kikan.worksheets() # 基幹シートのワークシート一覧

print("シートapiの設定完了")


# In[4]:


####################
# 
# メイン動作
# 
####################
try:
    # クローム起動
    driver = func.start_chrome()

    # 基幹システムログイン
    driver.get(url_kikan) # 基幹システムにアクセス
    time.sleep(1)

    func.input_text_by_id(driver, input_mail_id_kikan, mail_kikan) #メールアドレス入力
    func.input_text_by_id(driver, input_pass_id_kikan, pass_kikan) #パスワード入力
    func.click_by_xpath(driver, xpath_login_btn) # ログインボタンクリック
    print("基幹システムログイン")

    # CSVダウンロード
    func.click_by_xpath(driver, xpath_reserve_elm) # 予約/問合せをクリック
    func.delete_file(file_path_kikan) # 既存のCSV削除
    func.click_by_xpath(driver, xpath_csv_btn) # CSVダウンロードクリック
    time.sleep(5)
    print("csvダウンロード完了")

finally:
    driver.quit() # ブラウザ終了
    print("ブラウザ閉じる")


# シートが存在しなかったら今月のシート追加
if(func.exsit_sheet(sheets_kikan, conf.sheet_name_this_month)):
    ws_kikan.add_worksheet(title=conf.sheet_name_this_month, rows=1, cols=1)
    print("今月のシート追加")


# CSVをシートにアップロード
func.upload_csv(ws_kikan, conf.sheet_name_this_month+'!A1', file_path_kikan)

print("CSVアップロード")
print("終了")


# In[ ]:





# In[ ]:




