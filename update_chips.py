# ==========================================
# 📂 檔案名稱： update_chips.py (上櫃終極除錯版)
# 💡 任務：掛上最強偽裝器，並印出「凡甲」原始資料來抓 Bug！
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
    # 🔥 換上最強的瀏覽器偽裝，突破官方防火牆
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'X-Requested-With': 'XMLHttpRequest'
    }
    chip_stats = {}
    valid_days = 0
    current_date = datetime.now()
    
    print("🚀 開始啟動籌碼雷達，專注測試上櫃資料...")

    for _ in range(30):
        if valid_days >= 10: break
        if current_date.weekday() >= 5:
            current_date -= timedelta(days=1)
            continue
            
        dt_str = current_date.strftime("%Y%m%d")
        roc_y = current_date.year - 1911
        tpex_str = f"{roc_y}/{current_date.strftime('%m/%d')}"
        
        twse_url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={dt_str}&selectType=ALL&response=json"
        
        # 🐛 這裡把原本的 se=EW 改成 se=AL (全部)，才能抓到凡甲等所有上櫃股票
        tpex_url = f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={tpex_str}&se=AL"
        
        day_has_data = False
        twse_count = 0
        tpex_count = 0
        
        try:
            # 1. 上市 (TWSE)
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

        try:
            # 2. 上櫃 (TPEx) - 加入除錯追蹤
            res_tpex = requests.get(tpex_url, headers=headers, timeout=10, verify=False).json()
            if 'aaData' in res_tpex and len(res_tpex['aaData']) > 0:
                tpex_count = len(res_tpex['aaData'])
                day_has_data = True
                for row in res_tpex['aaData']:
                    code = str(row[0]).strip()
                    try:
                        # 🎯 【凡甲除錯器】印出凡甲的原始資料！
                        if code == "3526":
                            print(f"  👉 找到凡甲(3526)！總欄數: {len(row)} | 外資格: {row[8]} | 投信格: {row[11]}")

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
                print(f"  [上櫃警告] {dt_str}: 成功連線，但 aaData 裡面沒有資料！")
        except Exception as e:
            print(f"  [上櫃連線失敗] {dt_str}: {e}")

        if day_has_data:
            valid_days += 1
            print(f"✅ {dt_str} | 上市: {twse_count}筆 | 上櫃: {tpex_count}筆 (收集進度: {valid_days}/10)")
            
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
