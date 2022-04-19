#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 集客表のデータ取得


# In[2]:


########################
# 
# モジュール
# 
########################

# 一般モジュール
import pandas as pd
import re
from datetime import datetime as dt

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")


# In[3]:


########################
# 
# 設定
# 
########################

# シートキー
key_02 = "1WORdOLMmZU-7xyEhtZUJbjwyKe_q1Ye47E30xSA4ZS4"

key_list = [
    ["サルクルー", "1pcnplPHoRmxm4YJ0U9jIEUCGaInI_ANB4JqiaO3dz_Y"],
    ["FORCE", "14HWWhVcjK3-P_G7vcPuEmJ-mBTCsUVc5UB8KhzAPWXY"],
    ["セレス", "1YFt3uNb8fLaqYx3wP5POycCHoGPtpRDnCFS_FhM-tdY"],
    ["パフォテク", "1dZTiyKiJhXKAJWIzQHdY2EBA6WAo1U2V1GF88HiOWY0"],
    ["アレテコ", "1dqga1fiKNRFmjUDfAbtDG0zqp2aFGzzlBuelwM3c7eQ"],
    ["ブリーチ", "146GNHcMifTg4al17p4xYhLrf8mATZaVrQedHmWDi7HU"],
    ["FA", "14o-xmRmliJRZXIQ-0KiQntLeVeC6wkhnAfosOzzRa08"],
    ["ナハト", "1GftsMKvAbUxL3W7VNjxSrPUOn1KaxSU9COl14Hx9LnM"],
    ["エンジョイ", "1KA_TO7-NfiYzd7llUbWhOzKxMF37-3UMPdpYrgxfSeE"],
    ["無限", "1viSsDKXcVEX7ovcI_zIij3gxX-NE9YkUkttiM6jGmdk"],
    ["クラン", "12_AHStXqWHtD2sZbzyDtUMNKbQ27Vn21i2AsBAvvXkg"]
]

# シート名
name_02 = "22年4月"
name_to = "集客表"


# In[4]:


########################
# 
# コピペ
# 
########################

# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)


# In[5]:


# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
df_origin = df_origin[df_origin[4] != ""]   # ID空白除去


# In[6]:


# コピー先にアクセス
wb_origin = func.connect_gspread(key_02)
ws_origin = wb_origin.worksheet(name_02)


# In[7]:


# 後でループにする
for this_key in key_list:
    asp_name = this_key[0]
    key_to = this_key[1]

    # ROAS管理シートにアクセス
    wb_to = func.connect_gspread(key_to)
    ws_to = wb_to.worksheet(name_to)

    # まとめデータを貼り付け
    set_with_dataframe(ws_to, df_origin)

    # タイムスタンプ
    cell = ws_to.range(1,1,1,1)
    cell[0].value = str(dt.today())
    ws_to.update_cells(cell)
    
    print(asp_name)


# In[8]:


print("完了")


# In[ ]:




