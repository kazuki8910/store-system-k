# 変数を定義するファイル



####################
# 
# モジュール
# 
####################

# 一般
import datetime



####################
# 
# 日付関連
# 
####################

# 今日の日付関連
today        = datetime.date.today() # 今日の日付
this_year    = str(today.year)[2:4]  # 年
this_month   = str(today.month)      # 月

# 月（二桁）
if(len(this_month) == 1):
    this_month_d = "0" + this_month
else:
    this_month_d = this_month

# 今月の分析シートのファイル名
book_name_ana = this_year + "年" + this_month_d + "月"

# 今月の集客表のシート名
sheet_name_syukyaku = this_year + "年" + this_month + "月"




####################
# 
# GoogleAPI関連
# 
####################

# APIサービスアカウントキーのファイルパス
google_api_filepath = 'config/google_api.json'

# スコープ
google_api_scope = ['https://spreadsheets.google.com/feeds',
                    'https://www.googleapis.com/auth/drive']




####################
# 
# スプシ関連
# 
####################

# データを保管してるドライブのキー
drive_folder_key = '1rpxNAZqFJQQNnVpkSPzO90UyY8YDOYx'

# 設定シートのキー
sheet_key_conf = "1kvGW49-FnnlWf05FTs0ybDcqXIakycdyFf3cmnuXLus"
# 設定シートのシート名
sheet_name_conf = "シートキー"

#03シートのシートキー
sheet_key_03 = "1u95ZTRxAF64RM1t4Gyi3RYv_OcuazVgGg8cHhARFxbA"
