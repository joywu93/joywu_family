# ==========================================
# 📂 檔案名稱： update_chips.py (上櫃自動尋標版)
# 💡 任務：自動測試多種上櫃 API 格式，強制把資料挖出來！
# ==========================================

import os
import requests
import gspread
import json
import time
import re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import urllib3

# 關閉 SSL 警告，避免被阻擋
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

MASTER_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1s4dIaZb4FLOHrn_hwreHPkDKSobgtlaqFJjnsQiO1F4/edit"

def get_gspread_client():
    key_data = os.environ.get("GOOGLE_CREDENTIALS") or os.environ.get("GOOGLE_KEY_JSON")
    if not key_data: raise ValueError("找不到 Google 金鑰，請檢查 GitHub Secrets 設定！")
    creds_dict = json.loads(key_data)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return gspread.authorize(creds)

def clean_num(s):
    s = re.sub(r'<[^>]*>', '', str(s)).replace(',', '').strip()
    if not s or s in ['-', '--', '---']: return 0
    try: return int(float(s))
    except: return 0

def fetch_10_days_chips():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest'
    }
    chip_stats = {}
    valid_days = 0
    current_date = datetime.now()
    
    print("🚀 開始啟動籌碼雷達，啟用上櫃自動尋標模式...")

    for _ in range(30):
        if valid_days >= 10: break
        if current_date.weekday() >= 5:
            current_date -= timedelta(days=1)
            continue
            
        dt_str = current_date.strftime("%Y%m%d")
        
        # 準備各種日期格式
        roc_y = current_date.year - 1911
        west_y = current_date.year
        m_pad = current_date.strftime('%m')
        d_pad = current_date.strftime('%d')
        m_no = str(current_date.month)
        d_no = str(current_date.day)

        twse_url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={dt_str}&selectType=ALL&response=json"
        
        day_has_data = False
        twse_count = 0
        tpex_count = 0
        
        # ==========================================
        # 1. 抓取上市 (TWSE) - 維持不變
        # ==========================================
        try:
            res_twse = requests.get(twse_url, headers=headers, timeout=10, verify=False).json()
            if res_twse.get('stat') == 'OK' and 'data' in res_twse:
                twse_count = len(res_twse['data'])
                day_has_data = True
                fields = res_twse['fields']
                c_idx = fields.index('證券代號')
                f_idx = next((i for i, f in enumerate(fields) if '外陸資買賣超股數' in f or '外資及陸資買賣超股數' in f), -1)
                t_idx = next((i for i, f in enumerate(fields) if '投信買賣超股數' in f), -1)

                if f_idx != -1 and t_idx != -1:
                    for row in res_twse['data']:
                        code = str(row[c_idx]).strip()
                        f_net = clean_num(row[f_idx]) // 1000 
                        t_net = clean_num(row[t_idx]) // 1000

                        if code not in chip_stats: chip_stats[code] = {'f_days': 0, 'f_vol': 0, 't_days': 0, 't_vol': 0}
                        if f_net > 0: chip_stats[code]['f_days'] += 1
                        chip_stats[code]['f_vol'] += f_net
                        if t_net > 0: chip_stats[code]['t_days'] += 1
                        chip_stats[code]['t_vol'] += t_net
        except Exception as e:
            print(f"  [上市錯誤] {dt_str}: {e}")

        # ==========================================
        # 2. 抓取上櫃 (TPEx) - 自動尋標探測器
        # ==========================================
        # 櫃買 API 常常偷偷變更參數，我們一次準備 4 種最常見的合法組合來撞擊
        tpex_urls_to_try = [
            # 策略A：傳統民國年，省略 se 參數 (預設抓全市場，最可能成功)
            f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={roc_y}/{m_pad}/{d_pad}",
            # 策略B：改用西元年 (部分新版 API 已經強制改西元)
            f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={west_y}/{m_pad}/{d_pad}",
            # 策略C：傳統民國年 + se=EW (退而求其次，只抓電子類)
            f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={roc_y}/{m_pad}/{d_pad}&se=EW",
            # 策略D：民國年但不補零
            f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={roc_y}/{m_no}/{d_no}"
        ]

        res_tpex_data = None
        success_strategy = ""

        # 開始輪替測試
        for idx, test_url in enumerate(tpex_urls_to_try):
            try:
                res_tpex = requests.get(test_url, headers=headers, timeout=5, verify=False).json()
                if 'aaData' in res_tpex and len(res_tpex['aaData']) > 0:
                    res_tpex_data = res_tpex['aaData']
                    success_strategy = f"策略 {idx+1}"
                    break  # 只要其中一個成功，就立刻跳出測試迴圈
            except Exception:
                continue # 失敗就默默換下一個網址測試

        if res_tpex_data:
            tpex_count = len(res_tpex_data)
            day_has_data = True
            for row in res_tpex_data:
                code = str(row[0]).strip()
                try:
                    # 🎯 【凡甲除錯器】印出凡甲的原始資料！
                    if code == "3526":
                        print(f"  👉 找到凡甲(3526)！使用{success_strategy} | 外資格: {row[8]} | 投信格: {row[11]}")

                    if len(row) > 11:
                        f_net = clean_num(row[8]) // 1000  
                        t_net = clean_num(row[11]) // 1000 
                        
                        if code not in chip_stats: chip_stats[code] = {'f_days': 0, 'f_vol': 0, 't_days': 0, 't_vol': 0}
                        if f_net > 0: chip_stats[code]['f_days'] += 1
                        chip_stats[code]['f_vol'] += f_net
                        if t_net > 0: chip_stats[code]['t_days'] += 1
                        chip_stats[code]['t_vol'] += t_net
                except Exception as e: 
                    if code == "3526": print(f"  ❌ 凡甲解析失敗: {e}")
        else:
            # 如果這天上市有資料，但上櫃 4 個策略都失敗，才印出警告
            if twse_count > 0:
                print(f"  [上櫃警告] {dt_str}: 所有 API 參數策略皆無法取得 aaData！")

        if day_has_data:
            valid_days += 1
            print(f"✅ {dt_str} | 上市: {twse_count}筆 | 上櫃: {tpex_count}筆 (進度: {valid_days}/10)")
            
        current_date -= timedelta(days=1)
        time.sleep(2) 
        
    return chip_stats

def update_gsheet_chips(chip_stats):
    print("\n📝 準備將籌碼戰情報告寫入 Google 表單...")
    client = get_gspread_client()
    spreadsheet = client.open_by_url(MASTER_GSHEET_URL)
    target_sheets = [ws for ws in spreadsheet.worksheets() if any(n in ws.title for n in ["當年度表", "個股總表", "總表", "金融股"])]
    
    total_cells = 0
    for ws in target_sheets:
        data = ws.get_all_values()
        if not data: continue
        h = data[0]
        
        i_code = next((i for i, x in enumerate(h) if "代號" in str(x)), -1)
        i_t_days = next((i for i, x in enumerate(h) if "投信10日買天數" in str(x)), -1)
        i_t_vol = next((i for i, x in enumerate(h) if "投信10日買賣超" in str(x)), -1)
        i_f_days = next((i for i, x in enumerate(h) if "外資10日買天數" in str(x)), -1)
        i_f_vol = next((i for i, x in enumerate(h) if "外資10日買賣超" in str(x)), -1)
        
        if i_code == -1 or i_t_days == -1: continue

        cells = []
        for r_idx, row in enumerate(data[1:], start=2):
            if i_code < len(row):
                code = str(row[i_code]).split('.')[0].strip()
                if code in chip_stats:
                    d = chip_stats[code]
                    if i_t_days != -1: cells.append(gspread.Cell(row=r_idx, col=i_t_days+1, value=d['t_days']))
                    if i_t_vol != -1: cells.append(gspread.Cell(row=r_idx, col=i_t_vol+1, value=d['t_vol']))
                    if i_f_days != -1: cells.append(gspread.Cell(row=r_idx, col=i_f_days+1, value=d['f_days']))
                    if i_f_vol != -1: cells.append(gspread.Cell(row=r_idx, col=i_f_vol+1, value=d['f_vol']))
                    
        if cells:
            ws.update_cells(cells, value_input_option='USER_ENTERED')
            total_cells += len(cells)
            print(f"📊 分頁 [{ws.title}] 籌碼更新完成！")
            
    print(f"🎉 全部任務執行完畢！共更新了 {total_cells} 個籌碼欄位。")

if __name__ == "__main__":
    stats = fetch_10_days_chips()
    if stats: update_gsheet_chips(stats)
