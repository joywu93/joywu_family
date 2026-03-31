import gspread
from google.oauth2.service_account import Credentials
import json
import os

# ==========================================
# ⚙️ 設定區：測試表單網址
# ==========================================
TEST_GSHEET_URL = "https://docs.google.com/spreadsheets/d/174Aq8nLEXqgGX2lbeancNqwp_sppi-dsquP8kBlkELc/edit"

def get_gspread_client():
    # 讀取 GitHub Secrets 中的金鑰
    key_data = os.environ.get("GOOGLE_CREDENTIALS")
    if not key_data:
        raise ValueError("找不到 GOOGLE_CREDENTIALS 環境變數")
    creds = Credentials.from_service_account_info(
        json.loads(key_data) if isinstance(key_data, str) else dict(key_data), 
        scopes=['https://www.googleapis.com/auth/spreadsheets']
    )
    return gspread.authorize(creds)

def yearly_rollover_mission():
    try:
        print("🚀 啟動多房間跨年大搬風演習...")
        client = get_gspread_client()
        ss = client.open_by_url(TEST_GSHEET_URL)
        
        # 抓取這份試算表裡「所有的」分頁
        all_sheets = ss.worksheets()
        
        # 篩選出標題包含「當年度表」的分頁 (例如: 當年度表01, 當年度表02...)
        target_sheets = [ws for ws in all_sheets if "當年度表" in ws.title]
        
        if not target_sheets:
            print("❌ 找不到任何名稱包含「當年度表」的分頁，請檢查分頁名稱！")
            return

        for ws in target_sheets:
            print(f"\n=============================")
            print(f"🎯 正在處理分頁：【{ws.title}】")
            
            # 第一步：備份歸檔
            # 名稱會自動變成「歷史表單_2025_01」、「歷史表單_2025_02」...
            archive_name = ws.title.replace("當年度表", "歷史表單_2025_")
            print(f"📦 正在備份至：{archive_name}")
            try:
                ss.duplicate_sheet(ws.id, new_sheet_name=archive_name)
                print(f"✅ 備份成功！")
            except Exception as e:
                print(f"⚠️ 備份略過 (可能已存在同名分頁)。")

            # 第二步：欄位平移與升級 (25年 -> 26年)
            print("🧠 正在計算欄位平移邏輯...")
            headers = ws.row_values(1)
            new_headers = []
            cols_to_clear = [] 
            
            for i, h in enumerate(headers):
                if "24M11" in h or "24M12" in h: 
                    new_h = h.replace("24M", "25M")
                elif "25M" in h: 
                    new_h = h.replace("25M", "26M")
                    cols_to_clear.append(i + 1)
                elif "25Q" in h: 
                    new_h = h.replace("25Q", "26Q")
                    cols_to_clear.append(i + 1)
                else:
                    new_h = h
                new_headers.append(new_h)
                
            # 第三步：覆蓋新標題
            ws.update('1:1', [new_headers])
            print("✅ 標題年份升級成功！(25 -> 26)")
            
            # 第四步：清空舊年度格子 (保留標題，只清空下方的數字)
            if cols_to_clear:
                print(f"🧹 正在清空新年度戰場格子...")
                num_rows = len(ws.get_all_values())
                if num_rows > 1:
                    ranges_to_clear = []
                    for col_idx in cols_to_clear:
                        # 標記需要清除的範圍 (第 2 列到最後一列)
                        range_label = f"{gspread.utils.rowcol_to_a1(2, col_idx)}:{gspread.utils.rowcol_to_a1(num_rows, col_idx)}"
                        ranges_to_clear.append(range_label)
                    
                    # 批次清除，速度更快
                    ws.batch_clear(ranges_to_clear)
                print("✅ 該分頁資料清空完成！")

        print("\n🎉 跨年演習大獲全勝！所有當年度表均已更新完畢。")

    except Exception as e:
        print(f"❌ 演習發生未預期錯誤：{e}")

if __name__ == "__main__":
    yearly_rollover_mission()
