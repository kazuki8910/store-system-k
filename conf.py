# 変数を定義



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
today      = datetime.date.today() # 今日の日付
this_year  = str(today.year)[2:4]  # 年
this_month = str(today.month)      # 月

# 今月のシート名
sheet_name_this_month = this_year + "年" + this_month + "月"




####################
# 
# スプシ関連
# 
####################

# 基幹反映シートのシートキー
sheet_key_kikan = "1rlaJptyRr-MkbLejSakF9zK1vdLoP6pMAhYcoj6p7dI"

# 集客表元データのシートキー
sheet_key_syukyaku_origin = "1u95ZTRxAF64RM1t4Gyi3RYv_OcuazVgGg8cHhARFxbA"

# 集客表のシートキー
sheet_key_syukyaku = "1p8zfjMPWjGaV-HkkC-AE-o5nYBx5f8WiUNCtPno6XIw"