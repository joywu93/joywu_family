# ==========================================
# 📂 檔案名稱： update_chips.py (GitHub 雲端全自動版)
# 💡 任務：拿掉憑證檢查，並啟用 Cloudscraper 嘗試突破海外 IP 封鎖
# ==========================================

import os
import gspread
import json
import time
import re
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials
import urllib3
import cloudscraper

# 關閉 SSL 警告
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
    # 啟動 Cloudscraper 模擬真實的 Chrome 瀏覽器
    scraper = cloudscraper.create_scraper(browser={
        'browser': 'chrome',
        'platform': 'windows',
        'desktop': True
    })
    
    chip_stats = {}
    valid_days = 0
    current_date = datetime.now()
    
    print("🚀 開始啟動籌碼雷達，GitHub 雲端全自動版執行中...")

    for _ in range(30):
        if valid_days >= 10: break
        if current_date.weekday() >= 5:
            current_date -= timedelta(days=1)
            continue
            
        dt_str = current_date.strftime("%Y%m%d")
        roc_y = current_date.year - 1911
        m_pad = current_date.strftime('%m')
        d_pad = current_date.strftime('%d')
        
        twse_url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={dt_str}&selectType=ALL&response=json"
        tpex_url = f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={roc_y}/{m_pad}/{d_pad}"
        
        day_has_data = False
        twse_count = 0
        tpex_count = 0
        
        # ==========================================
        # 1. 抓取上市 (TWSE)
        # ==========================================
        try:
            # 💡 這裡已經把 verify=False 拿掉了
            res_twse = scraper.get(twse_url, timeout=15).json()
            if res_twse.get('stat') == 'OK' and 'data' in res_twse:
                twse_count = len(res_twse['data'])
                day_has_data = True
                fields = res_twse['fields']
                
                c_idx = next((i for i, f in enumerate(fields) if '代號' in f), -1)
                f_idx = next((i for i, f in enumerate(fields) if '外陸資買賣超股數' in f or '外資及陸資買賣超股數' in f), -1)
                t_idx = next((i for i, f in enumerate(fields) if '投信買賣超股數' in f), -1)

                if c_idx != -1 and f_idx != -1 and t_idx != -1:
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
        # 2. 抓取上櫃 (TPEx)
        # ==========================================
        if twse_count > 0: 
            try:
                # 💡 這裡已經把 verify=False 拿掉了
                res_tpex_raw = scraper.get(tpex_url, timeout=15)
                
                if res_tpex_raw.status_code != 200 or 'html' in res_tpex_raw.text.lower()[:100]:
                    print(f"  [上櫃阻擋] {dt_str}: 伺服器回應異常，可能仍被防火牆阻擋。")
                else:
                    res_tpex = res_tpex_raw.json()
                    
                    if 'aaData' in res_tpex and len(res_tpex['aaData']) > 0:
                        tpex_count = len(res_tpex['aaData'])
                        for row in res_tpex['aaData']:
                            code = str(row[0]).strip()
                            try:
                                if len(row) > 11:
                                    f_net = clean_num(row[8]) // 1000  
                                    t_net = clean_num(row[11]) // 1000 
                                    
                                    if code not in chip_stats: chip_stats[code] = {'f_days': 0, 'f_vol': 0, 't_days': 0, 't_vol': 0}
                                    if f_net > 0: chip_stats[code]['f_days'] += 1
                                    chip_stats[code]['f_vol'] += f_net
                                    if t_net > 0: chip_stats[code]['t_days'] += 1
                                    chip_stats[code]['t_vol'] += t_net
                            except Exception: 
                                pass
            except Exception as e:
                print(f"  [上櫃連線失敗] {dt_str}: {e}")

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
        
        if i_code == -1:
            print(f"  ⚠️ 分頁 [{ws.title}] 找不到包含「代號」的欄位，跳過更新。")
            continue

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
    if stats: 
        update_gsheet_chips(stats)
