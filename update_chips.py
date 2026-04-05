# ==========================================
# 📂 檔案名稱： update_chips.py (籌碼搬運工)
# 💡 任務：自動往回抓取 10 個「交易日」的上市/上櫃法人買賣超，並精準寫入表單！
# ==========================================

import os
import requests
import gspread
import json
import time
from datetime import datetime, timedelta
from google.oauth2.service_account import Credentials

# 您的 Google Sheet 網址
MASTER_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1s4dIaZb4FLOHrn_hwreHPkDKSobgtlaqFJjnsQiO1F4/edit"

def get_gspread_client():
    key_data = os.environ.get("GOOGLE_CREDENTIALS") or os.environ.get("GOOGLE_KEY_JSON")
    if not key_data: raise ValueError("找不到 Google 金鑰，請檢查 GitHub Secrets 設定！")
    creds_dict = json.loads(key_data)
    creds = Credentials.from_service_account_info(creds_dict, scopes=['https://www.googleapis.com/auth/spreadsheets'])
    return gspread.authorize(creds)

def fetch_10_days_chips():
    headers = {'User-Agent': 'Mozilla/5.0'}
    chip_stats = {}
    valid_days = 0
    current_date = datetime.now()
    
    print("🚀 開始啟動籌碼雷達，往前推算 10 個交易日...")

    # 最多往前找 30 天，確保能湊滿 10 個交易日 (避開連假)
    for _ in range(30):
        if valid_days >= 10: break
        
        # 六日直接跳過
        if current_date.weekday() >= 5:
            current_date -= timedelta(days=1)
            continue
            
        dt_str = current_date.strftime("%Y%m%d")
        roc_y = current_date.year - 1911
        tpex_str = f"{roc_y}/{current_date.strftime('%m/%d')}"
        
        twse_url = f"https://www.twse.com.tw/rwd/zh/fund/T86?date={dt_str}&selectType=ALL&response=json"
        tpex_url = f"https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php?l=zh-tw&o=json&d={tpex_str}&se=EW"
        
        day_has_data = False
        
        try:
            # 1. 抓取上市 (TWSE)
            res_twse = requests.get(twse_url, headers=headers, timeout=10).json()
            if res_twse.get('stat') == 'OK' and 'data' in res_twse:
                day_has_data = True
                fields = res_twse['fields']
                c_idx = fields.index('證券代號')
                # 動態尋找外資與投信的欄位
                f_idx = next((i for i, f in enumerate(fields) if '外陸資買賣超股數' in f or '外資及陸資買賣超股數' in f), -1)
                t_idx = next((i for i, f in enumerate(fields) if '投信買賣超股數' in f), -1)

                if f_idx != -1 and t_idx != -1:
                    for row in res_twse['data']:
                        code = str(row[c_idx]).strip()
                        # 官方單位是「股」，除以 1000 換算成「張」
                        f_net = int(str(row[f_idx]).replace(',', '')) // 1000 
                        t_net = int(str(row[t_idx]).replace(',', '')) // 1000

                        if code not in chip_stats:
                            chip_stats[code] = {'f_days': 0, 'f_vol': 0, 't_days': 0, 't_vol': 0}

                        if f_net > 0: chip_stats[code]['f_days'] += 1
                        chip_stats[code]['f_vol'] += f_net

                        if t_net > 0: chip_stats[code]['t_days'] += 1
                        chip_stats[code]['t_vol'] += t_net

            # 2. 抓取上櫃 (TPEx)
            res_tpex = requests.get(tpex_url, headers=headers, timeout=10).json()
            if 'aaData' in res_tpex and len(res_tpex['aaData']) > 0:
                day_has_data = True
                for row in res_tpex['aaData']:
                    code = str(row[0]).strip()
                    try:
                        if len(row) >= 14:
                            f_net = int(str(row[4]).replace(',', '')) // 1000
                            t_net = int(str(row[13]).replace(',', '')) // 1000
                            
                            if code not in chip_stats:
                                chip_stats[code] = {'f_days': 0, 'f_vol': 0, 't_days': 0, 't_vol': 0}

                            if f_net > 0: chip_stats[code]['f_days'] += 1
                            chip_stats[code]['f_vol'] += f_net

                            if t_net > 0: chip_stats[code]['t_days'] += 1
                            chip_stats[code]['t_vol'] += t_net
                    except: pass
        except Exception as e:
            print(f"⚠️ {dt_str} 抓取發生錯誤 (可能為國定假日無資料)")

        if day_has_data:
            valid_days += 1
            print(f"✅ 成功獲取 {dt_str} 籌碼數據 (已收集 {valid_days}/10 天)")
            
        current_date -= timedelta(days=1)
        time.sleep(2) # 禮貌性延遲，避免被證交所封鎖
        
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
        
        # 尋找您剛建立的 4 個新欄位
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
