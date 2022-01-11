#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 基幹データ取り込み


# In[2]:


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


# In[3]:


####################
# 
# 変数
# 
####################

# 基幹システムのログインURL
url_kikan  = "https://stlassh.com/admin/" #URL

# ログイン関連
mail_kikan = "info@stlassh.com" #メールアドレス
pass_kikan = "2sYsLpebQ3sx"     #PASS
input_mail_id_kikan = 'user_email'    #メールアドレスのインプットタグのID
input_pass_id_kikan = 'user_password' #パスワードのインプットタグのID
xpath_login_btn = '//*[@id="new_user"]/div[3]/input' #ログインボタンのXPath

# ダウンロード関連
xpath_reserve_elm = '/html/body/main/div/div[1]/a'    # 予約/問合せの要素のXpath
xpath_csv_btn     ='/html/body/main/div/form/input[3]'# CSV出力ボタンXpath
file_path_kikan = "data/reservations.csv" # CSVの保存先パス

print("変数の定義完了")


# In[4]:


####################
# 
# 基幹からデータダウンロード
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


# In[5]:


####################
# 
# シートにデータ反映
# 
####################

# 分析シートキー取得
sheet_key_ana = func.get_ana_key()

# 分析シートにアクセス
wb_ana = func.connect_gspread(sheet_key_ana) # 分析

# 基幹シートが存在しなかったら今月のシート追加
sheets_ana = wb_ana.worksheets() # 基幹シートのワークシート一覧
sheet_name = "基幹"
if(func.exsit_sheet(sheets_ana, sheet_name)):
    wb_ana.add_worksheet(title=sheet_name, rows=1, cols=1)
    print("基幹シート追加")

# CSVをシートにアップロード
func.upload_csv(wb_ana, sheet_name+'!A1', file_path_kikan)
print("CSVアップロード")

# 「data」ディレクトリのデータ削除
func.delete_data_files()

print("終了")

