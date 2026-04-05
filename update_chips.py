def get_tpex_date(date_str):
    """
    將 'YYYYMMDD' 轉換為上櫃 API 需要的 'YYY/MM/DD' 格式
    例如: '20260403' -> '115/04/03'
    """
    # 將西元年轉成民國年
    roc_year = int(date_str[:4]) - 1911
    month = date_str[4:6]
    day = date_str[6:]
    
    return f"{roc_year}/{month}/{day}"

# 假設你在爬蟲迴圈中的原始日期變數叫做 current_date_str (例如 '20260403')
# --- 在組合 TPEx API 網址時 ---

# 1. 轉換日期格式
tpex_date = get_tpex_date(current_date_str)

# 2. 組合 API (請確認你的 Endpoint 是否為這個)
# 上櫃三大法人買賣超 API 通常長這樣：
tpex_url = f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={tpex_date}&s=0,asc"

# 3. 發送 Request
# response = requests.get(tpex_url, headers=your_headers)
# ... 下面的解析邏輯維持不變 ...
