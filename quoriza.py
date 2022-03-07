#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 成果確認シート取り込み


# In[1]:


####################
# 
# モジュール
# 
####################

# 一般
import time
import glob

# 自作
import func
import conf

print("モジュールのインポート完了")


# In[2]:


####################
# 
# 変数
# 
####################

# クオリザ関連
url_quoriza  = "https://quoriza.net/client/login" # クオリザのURL
id_quoriza   = "RvKVjPRn" # ID
pass_quoriza = "jbqoDbrn" # パスワード

print("変数の定義完了")


# In[6]:


####################
# 
# クオリザにログイン
# 
####################

try:
    
    # クローム起動
    driver = func.start_chrome(False)

    # ログイン画面にアクセス
    driver.get(url_quoriza)
    time.sleep(1)

    # ログイン操作
    xpath_id = "/html/body/div[2]/form/div[1]/input"
    func.input_text_by_xpath(driver, xpath_id, id_quoriza) # ID入力
    xpath_pass = "/html/body/div[2]/form/div[2]/input"
    func.input_text_by_xpath(driver, xpath_pass, pass_quoriza) # パスワード入力
    xpath_login_btn = "/html/body/div[2]/form/div[3]/input"
    func.click_by_xpath(driver, xpath_login_btn, 7) # ボタンクリック

    print("クオリザのログイン成功")

    
####################
# 
# データをダウンロード
# 
####################

    # 「成果コンバージョン」をクリック
    xpath_hov_menu = "/html/body/div[1]/div[2]/nav/div/div/ul/li[3]"
    func.click_by_xpath(driver, xpath_hov_menu)

    # 「成果コンバージョン一覧」をクリック
    xpath_cv_list = "/html/body/div[1]/div[2]/nav/div/div/ul/li[3]/ul/li[1]/a"
    func.click_by_xpath(driver, xpath_cv_list, 3)

    # 「当月」ボタンをクリック
    xpath_this_month_btn = "/html/body/div[1]/div[3]/div/div[2]/div/form/div/div[2]/div/div/div[2]/a[1]"
    func.click_by_xpath(driver, xpath_this_month_btn)

    # ダウンロードボタンをクリック
    xpath_download_btn = "/html/body/div[1]/div[3]/div/div[2]/div/form/div/div[8]/div/button[2]"
    func.click_by_xpath(driver, xpath_download_btn, 5)

    print("ダウンロード完了")


finally:
    driver.quit() # ブラウザ終了
    print("ブラウザ閉じる")


# In[4]:


####################
# 
# シートにデータ反映
# 
####################

# 分析シートキー取得
sheet_key_ana = func.get_ana_key()

# 分析シートにアクセス
wb_ana = func.connect_gspread(sheet_key_ana) # 分析

# クオリザシートが存在しなかったら今月のシート追加
sheets_ana = wb_ana.worksheets() # クオリザシートのワークシート一覧
sheet_name = "クオリザ"
if(func.exsit_sheet(sheets_ana, sheet_name)):
    wb_ana.add_worksheet(title=sheet_name, rows=1, cols=1)
    print("クオリザシート追加")

# ダウンロードしたファイルパス取得
file_path = glob.glob("data/*")[0]

# CSVをシートにアップロード
func.upload_csv(wb_ana, sheet_name+'!A1', file_path)
print("CSVアップロード")

# 「data」ディレクトリのデータ削除
func.delete_data_files()

print("終了")

