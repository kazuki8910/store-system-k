#!/usr/bin/env python
# coding: utf-8

# In[1]:


# 成果確認シートの内容取得


# In[2]:


####################
# 
# モジュール
# 
####################

# 一般
import datetime as dt
import pandas as pd

# スプシ関連
from gspread_dataframe import set_with_dataframe

# 自作
import conf
import func

print("モジュールインポート")


# In[3]:


####################
# 
# 変数
# 
####################

# 今月の成果シート名
today = dt.date.today()
this_year  = str(today.year)
this_month = str(today.month)
sheet_name_seika = this_year + "年" + this_month + "月成果" 

# 月末月初取得
beginning_day = dt.datetime.combine(today.replace(day=1), dt.time())             # 月初の日付
end_day       = dt.datetime(today.year, today.month+1, 1) - dt.timedelta(days=1) # 月末の日付

# 成果確認シートの情報
sheets_info = [
    {
        "name":"ブリーチ",
        "key" :"1tYvW0ntpbCVOpY2UGcqdQD_DQhA7HesfKkg0NPAp65g"
    },
    {
        "name":"フォースリー",
        "key" :"1O3JayPUeW91PaEqcjX9Zm3_Jg2Ltzn94UYuqya5Il68"
    },
    {
        "name":"レントラ",
        "key" :"1jlgzK2U__YYCZO6wk_klEAkXbQbYXWuxhxE-AZTWkCQ"
    },
    {
        "name":"サルクルー",
        "key" :"1cbrrpc2GDTEGn2BauduPaC6-NRmAVJX5YMszz4CSCKQ"
    },
    {
        "name":"セレス",
        "key" :"1LNkA6RZ-__kvcwe3qhOg604ITnCj0W6DJxZqWQzRJOc"
    },
    {
        "name":"FA",
        "key" :"1qxvzPJn3c84VCfWrrKLdO03wc6fhauKdTjEqqYhyUNU"
    },
    {
        "name":"パフォーマンステクノロジー",
        "key" :"1ULCaCoNvoWCAW7khU5fqZIgVFt5ZWibYB0GvP2wQ3fc"
    },
    {
        "name":"クラン",
        "key" :"1wZGNHbgyoEQDGqWiePYSWOboR5xOLW8_nSalkTB716s"
    }
]

print("変数定義完了")


# In[4]:


####################
# 
# 成果シートの情報取得
# 
####################
# データを格納するDataframe
df_seika_sum = pd.DataFrame()

# 成果を抽出するループ
for info in sheets_info:
    key  = info["key"]
    name = info["name"]

    # 成果シートにアクセス
    wb_seika = func.connect_gspread(key)
    ws_seika = wb_seika.worksheet(sheet_name_seika)

    # 成果シートのデータ取得
    df_seika = pd.DataFrame(ws_seika.get_all_values()) # 元データ取得
    df_seika = df_seika.drop(index=df_seika.index[[0,1]]) # 先頭行削除
    df_seika = df_seika.reindex(columns=[9,1,2,4,5,6]) # 必要列抽出

    # カラム名つける
    df_seika = df_seika.rename(columns={
        9:"成果識別番号",
        1:"成果発生日時",
        2:"獲得方法",
        4:"メディア名",
        5:"LP名",
        6:"グロス(税込)"
    })

    # 日付と時間に列を分ける（レントラだけ仕様が違う）
    ach_day  = df_seika['成果発生日時'].str.extract('(\d+[/-]\d+[/-]+\d+)', expand=True) # 成果発生日
    ach_day  = ach_day.apply(lambda x: x.str.replace('-','/'))                           # "-"を"/"に変換
    ach_time = df_seika['成果発生日時'].str.extract('(\d+:\d:*\d*)', expand=True)        # 成果発生時間

    # Macの場合
    df_seika.insert(1,"成果発生日",ach_day[0])    # 成果発生日を追加
    df_seika.insert(2,"成果発生時間",ach_time[0]) # 成果発生時間を追加

    # Windowsの場合
    # df_seika.insert(1,"成果発生日",ach_day)    # 成果発生日を追加
    # df_seika.insert(2,"成果発生時間",ach_time) # 成果発生時間を追加

    # 成果発生日時削除
    df_seika = df_seika.drop("成果発生日時",axis=1) 

    # ASP名追加
    df_seika.insert(5,"ASP",name) 

    # 該当しない日付を削除
    df_seika['成果発生日'] = pd.to_datetime(df_seika['成果発生日'], format='%Y/%m/%d') # 成果発生日の列を日付型に変換
    df_seika = df_seika[(df_seika['成果発生日'] >= beginning_day) & (df_seika['成果発生日'] <= end_day)] # 該当する日付を抽出
    df_seika.reset_index(drop=True, inplace=True) # 行番号振り直し

    # データ出力
    df_seika_sum = pd.concat([df_seika_sum,df_seika])
    print(name+":"+str(len(df_seika)))
    
print("データ出力完了")


# In[5]:


########################
# 
# シートに反映
# 
########################

# 分析シートキー取得
sheet_key_ana = func.get_ana_key()

# 分析シートにアクセス
wb_ana = func.connect_gspread(sheet_key_ana)

# 集客表シートが存在しなかったら今月のシート追加
sheets_ana = wb_ana.worksheets() # 集客表シートのワークシート一覧
sheet_name = "成果"
if(func.exsit_sheet(sheets_ana, sheet_name)):
    wb_ana.add_worksheet(title=sheet_name, rows=1, cols=1)
    print("成果シート追加")

# 集客表シートにアクセス
ws = wb_ana.worksheet(sheet_name)

# データフレームをスプシに反映
set_with_dataframe(ws, df_seika_sum)

print("成果シートへの反映完了")


# In[ ]:




