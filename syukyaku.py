#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 集客表のデータ取得
########################
# 
# モジュール
# 
########################

# 一般モジュール
import pandas as pd
import re

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作モジュール
import func
import conf

print("モジュールのインポート完了")


# In[3]:


########################
# 
# 集客表元データ整形
# 
########################

# 集客表元データのシートにアクセス
wb_origin = func.connect_gspread(conf.sheet_key_03)       # 03シートにアクセス
ws_origin = wb_origin.worksheet(conf.sheet_name_syukyaku) # 今月の集客表取得

# 元データ取得
df_origin = pd.DataFrame(ws_origin.get_all_values()) # 元データ取得
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

# 空欄データを除去
df_origin = df_origin[df_origin['成果識別ID'] != ""]   # ID空白除去
df_origin = df_origin[df_origin['成果識別ID'] != "ID"] # ｢ID｣の文字列除去
df_origin = df_origin.sort_values('成果識別ID')        # ID順に並び替え
df_origin = df_origin.reset_index(drop=True)          # 番号振り直し

# 問合せ順に並び替え
df_origin = df_origin.sort_values('問合せ日')

# 番号振る
serial_num = pd.RangeIndex(start=1, stop=len(df_origin.index) + 1, step=1)
df_origin['No'] = serial_num

# 列の並び替え
df_origin = df_origin.reindex(columns=[
    "No",
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

print("顧客データの整形完了")


# In[ ]:


# 入金、契約金を数値型に変換
def to_num(string):
    try:
        cost = int(re.sub("\¥|\,","",string))
    except:
        cost = ""
    return cost
df_origin["契約金額"] = df_origin["契約金額"].map(to_num)
df_origin["入金金額"] = df_origin["入金金額"].map(to_num)


# In[ ]:


########################
# 
# 集客表シートに反映
# 
########################

# 分析シートキー取得
sheet_key_ana = func.get_ana_key()

# 分析シートにアクセス
wb_ana = func.connect_gspread(sheet_key_ana)

# 集客表シートが存在しなかったら今月のシート追加
sheets_ana = wb_ana.worksheets() # 集客表シートのワークシート一覧
sheet_name = "集客表"
if(func.exsit_sheet(sheets_ana, sheet_name)):
    wb_ana.add_worksheet(title=sheet_name, rows=1, cols=1)
    print("集客表シート追加")

# 集客表シートにアクセス
ws_syukyaku = wb_ana.worksheet(sheet_name)

# データフレームをスプシに反映
set_with_dataframe(ws_syukyaku, df_origin)

print("集客表シートへの反映完了")


# In[ ]:


# In[ ]:




