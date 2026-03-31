import gspread
from google.oauth2.service_account import Credentials
import json
import time

# ==========================================
# ⚙️ 設定區：請填入您的測試表單網址
# ==========================================
TEST_GSHEET_URL = "https://docs.google.com/spreadsheets/d/174Aq8nlEXqgGX2IbeancNqwp_sppi-dsquP8kBIkELc/edit"
CREDENTIALS_NAME = "GOOGLE_CREDENTIALS" # 讀取您 GitHub Actions 裡的金鑰

def yearly_rollover_mission():
    try:
        # 1. 登入 Google Sheet
        # (這裡假設您在 GitHub 環境跑，會讀取您設定好的 Secrets)
        import os
        from google.oauth2 import service_account
        
        # 這裡為了方便測試，手動讀取您環境中的金鑰
        # 如果是在 GitHub 跑，這段會自動抓取您存好的 JSON
        # 若在本地測試，請確保環境變數已設定
        print("🚀 啟動跨年大搬風演習...")
        
        # 這裡模擬 GitHub Actions 的讀取方式
        # (如果您是手動執行，請確保您的 Secret 已正確注入)
        # 為求保險，我們直接使用您已經運作順利的 gspread 授權邏輯
        
        # ... 授權代碼略 (與您 update_finance.py 相同) ...

        # 2. 開啟測試表單與定位分頁
        client = get_gspread_client() # 呼叫授權函式
        ss = client.open_by_url(TEST_GSHEET_URL)
        ws_current = ss.worksheet("當年度表")
        
        # 3. 第一步：備份歸檔
        # 將目前的「當年度表」複製一份，改名為「歷史表單_2025」
        print("📦 正在執行備份歸檔：當年度表 -> 歷史表單_2025")
        archive_name = "歷史表單_2025"
        try:
            ss.duplicate_sheet(ws_current.id, insert_sheet_index=1, new_sheet_name=archive_name)
            print(f"✅ 備份成功！已建立分頁：{archive_name}")
        except Exception as e:
            print(f"⚠️ 備份提醒：可能已存在同名分頁，略過複製步驟。({e})")

        # 4. 第二步：欄位平移與升級 (25年 -> 26年)
        print("🧠 正在計算欄位平移邏輯...")
        headers = ws_current.row_values(1)
        new_headers = []
        
        cols_to_clear = [] # 記錄哪些欄位需要清空數據
        
        for i, h in enumerate(headers):
            # 邏輯 A：24M11, 24M12 -> 變成 25M11, 25M12 (作為去年度基期)
            if "24M11" in h: new_h = h.replace("24M", "25M")
            elif "24M12" in h: new_h = h.replace("24M", "25M")
            
            # 邏輯 B：所有的 25M 或 25Q -> 升級為 26M 或 26Q (新戰場)
            elif "25M" in h: 
                new_h = h.replace("25M", "26M")
                cols_to_clear.append(i + 1)
            elif "25Q" in h: 
                new_h = h.replace("25Q", "26Q")
                cols_to_clear.append(i + 1)
            
            # 邏輯 C：其餘欄位（代號、名稱、發行量等）維持不變
            else:
                new_h = h
            
            new_headers.append(new_h)
            
        # 5. 第三步：執行表單更新
        # 覆蓋新標題
        ws_current.update('1:1', [new_headers])
        print("✅ 標題年份升級成功！(25 -> 26)")
        
        # 6. 第四步：清空舊年度格子 (保留標題，只清空下方的數字)
        if cols_to_clear:
            print(f"🧹 正在清空新年度戰場格子 (共 {len(cols_to_clear)} 個欄位)...")
            num_rows = len(ws_current.get_all_values())
            if num_rows > 1:
                for col_idx in cols_to_clear:
                    # 清除該欄第 2 列到最後一列的內容
                    # 使用 batch 操作會更快，這裡先用簡單邏輯演示
                    range_label = f"{gspread.utils.rowcol_to_a1(2, col_idx)}:{gspread.utils.rowcol_to_a1(num_rows, col_idx)}"
                    ws_current.batch_clear([range_label])
            print("✅ 資料清空完成，永動機已準備好迎接新數據！")

        print("\n🎉 跨年演習大獲全勝！請前往 Google Sheet 驗收成果。")

    except Exception as e:
        print(f"❌ 演習失敗，錯誤回報：{e}")

# ... (其餘授權輔助函式略) ...
