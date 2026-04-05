import gspread
from google.oauth2.service_account import Credentials
import json
import os

# ==========================================
# 🎯 正式實戰區：主力總表網址
# ==========================================
MASTER_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1s4dIaZb4FLOHrn_hwreHPkDKSobgtlaqFJjnsQiO1F4/edit?usp=sharing"

def get_gspread_client():
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
        print("🚀 [正式環境] 啟動多房間跨年大搬風...")
        client = get_gspread_client()
        ss = client.open_by_url(MASTER_GSHEET_URL)
        all_sheets = ss.worksheets()
        
        # 🔥 升級點 1：將所有主力分頁（包含金融股）納入跨年雷達
        target_names = ["當年度表", "個股總表", "總表", "金融股"]
        target_sheets = [ws for ws in all_sheets if any(n in ws.title for n in target_names)]
        
        if not target_sheets:
            print("❌ 找不到任何符合跨年條件的分頁！")
            return

        for ws in target_sheets:
            print(f"\n=============================")
            print(f"🎯 正在處理主力分頁：【{ws.title}】")
            
            # 🔥 升級點 2：智慧命名備份分頁，確保「金融股」備份時不會撞名
            if "當年度表" in ws.title:
                archive_name = ws.title.replace("當年度表", "歷史表單_2025_")
            else:
                archive_name = f"歷史表單_2025_{ws.title}"
                
            print(f"📦 正在備份至：{archive_name}")
            try:
                ss.duplicate_sheet(ws.id, new_sheet_name=archive_name)
                print(f"✅ 備份成功！")
            except Exception as e:
                print(f"⚠️ 備份略過 (可能已存在同名歷史分頁)。")

            # 第二步：欄位平移邏輯 (月增與年增完全不介入)
            print("🧠 正在計算欄位平移與清空邏輯...")
            headers = ws.row_values(1)
            new_headers = []
            cols_to_clear = [] 
            
            for i, h in enumerate(headers):
                new_h = h 
                
                if "24M11" in h or "24M12" in h: 
                    new_h = h.replace("24M", "25M")
                elif "25M" in h: 
                    new_h = h.replace("25M", "26M")
                    cols_to_clear.append(i + 1)
                elif "25Q" in h: 
                    new_h = h.replace("25Q", "26Q")
                    cols_to_clear.append(i + 1)
                
                # 注意：只要標題包含 "月增" 或 "年增"，程式會直接無視，保留原標題與原數據。
                new_headers.append(new_h)
                
            # 第三步：覆蓋新標題
            ws.update('1:1', [new_headers])
            print("✅ 標題年份全面升級完成！")
            
            # 第四步：精準清空舊年度格子
            if cols_to_clear:
                print(f"🧹 正在清空新年度戰場格子...")
                num_rows = len(ws.get_all_values())
                if num_rows > 1:
                    ranges_to_clear = []
                    for col_idx in cols_to_clear:
                        range_label = f"{gspread.utils.rowcol_to_a1(2, col_idx)}:{gspread.utils.rowcol_to_a1(num_rows, col_idx)}"
                        ranges_to_clear.append(range_label)
                    ws.batch_clear(ranges_to_clear)
                print("✅ 營收與 EPS 格子已清空，月增/年增與靜態資料已完美保留！")

        print("\n🎉 主系統跨年大搬風完美竣工！")

    except Exception as e:
        print(f"❌ 執行發生錯誤：{e}")

if __name__ == "__main__":
    yearly_rollover_mission()
