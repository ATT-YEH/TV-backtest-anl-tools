import pandas as pd

# === 請依需求修改以下路徑 ===
new_file_path = "C:/Users/user/Downloads/CME_MINI_ES1!, 60 (1).csv"
# ✅ 新資料（CSV）
old_file_path = "C:/Users/user/OneDrive/桌面/TV/SP500/標普近月1H歷史_210402-250624.xlsx"  # ✅ 舊資料（XLSX）
output_file_path = "C:/Users/user/OneDrive/桌面/TV/SP500/標普近月1H歷史_210402-250726.xlsx"  # ✅ 匯出也是 XLSX

# === 載入資料 ===
df_new = pd.read_csv(new_file_path)
df_old = pd.read_excel(old_file_path)

# === 移除整列都是空值的 row ===
df_new = df_new.dropna(how='all')

# === 欄位名稱標準化 ===
df_new.columns = [col.strip().lower() for col in df_new.columns]
df_old.columns = [col.strip().lower() for col in df_old.columns]

# === 移除 time 欄位空值或非數字 ===
df_new = df_new[pd.to_numeric(df_new["time"], errors="coerce").notnull()]
df_new["time"] = df_new["time"].astype(int)

# === 新增 UTC 時間欄位 ===
df_new["utc time"] = pd.to_datetime(df_new["time"], unit='s')

# === 確保舊資料時間欄位是 datetime 格式 ===
if df_old["utc time"].dtype != '<M8[ns]':
    df_old["utc time"] = pd.to_datetime(df_old["utc time"])

# === 合併、去重、排序 ===
df_merged = pd.concat([df_old, df_new], ignore_index=True)
df_merged = df_merged.drop_duplicates(subset=["utc time"])
df_merged = df_merged.sort_values("utc time").reset_index(drop=True)

# === ✅ 輸出為 Excel 格式 ===
df_merged.to_excel(output_file_path, index=False)
print(f"✅ 合併完成，已輸出到：{output_file_path}")
