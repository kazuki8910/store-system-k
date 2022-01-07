#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 集客表反映


########################
# 
# モジュール
# 
########################

# 一般モジュール
import pandas as pd

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")


# In[2]:


########################
# 
# 集客表元データ整形
# 
########################

# 集客表元データのシートにアクセス
ws_origin = func.connect_gspread(conf.sheet_key_syukyaku_origin) # 基幹シート
sheet_origin = ws_origin.worksheet(conf.sheet_name_this_month)

# 元データ取得
df_origin = pd.DataFrame(sheet_origin.get_all_values()) # 元データ取得
df_origin = df_origin.drop(index=df_origin.index[[0,1,2]], columns=df_origin.columns[[0,1,12,13]]) # 不要行列削除

# カラム名つける
df_origin = df_origin.rename(columns={
    2:'店舗名',
    3:"問合せ日",
    4:"成果識別ID",
    5:"名前",
    6:"来店予定日",
    7:"来店状況",
    8:"契約状況",
    9:"契約金額",
    10:"入金金額",
    11:"媒体",
    14:"進捗状況",
    15:"年齢"
})

# 列の並び替え
df_origin = df_origin.reindex(columns=[
    "成果識別ID",
    "名前",
    "年齢",
    "店舗名",
    "媒体",
    "問合せ日",
    "来店予定日",
    "来店状況",
    "契約状況",
    "進捗状況",
    "契約金額",
    "入金金額"
])

df_origin = df_origin[df_origin['成果識別ID'] != ""]   # ID空白除去
df_origin = df_origin[df_origin['成果識別ID'] != "ID"] # 「ID」の文字列除去
df_origin = df_origin.sort_values('成果識別ID')        # ID順に並び替え
df_origin = df_origin.reset_index(drop=True)           # 番号振り直し

print("顧客データの整形完了")


# In[3]:


########################
# 
# 集客表シートに反映
# 
########################

# 集客表シートにアクセス
ws_syukyaku = func.connect_gspread(conf.sheet_key_syukyaku) # 集客表シート
sheets_syukyaku = ws_syukyaku.worksheets() # 集客表シートのワークシート一覧

# シートが存在しなければ追加
if(func.exsit_sheet(sheets_syukyaku, conf.sheet_name_this_month)):
    ws_syukyaku.add_worksheet(title=conf.sheet_name_this_month, rows=1, cols=1)
    print("今月のシート追加")

# 今月のシートにアクセス
sheet_this_month = ws_syukyaku.worksheet(conf.sheet_name_this_month)

# データフレームをスプシに反映
set_with_dataframe(sheet_this_month, df_origin)

print("集客表シートへの反映完了")


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:




