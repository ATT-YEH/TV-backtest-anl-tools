import pandas as pd

# Step 1: 讀取資料
file1_path = "C:/Users/user/OneDrive/桌面/TV/SP500/標普超藍/標普utc0實盤_2025-04-05.xlsx"
file2_path = "C:/Users/user/OneDrive/桌面/TV/SP500/SP500-1h滾動波動率250726測試.csv"

# 讀取交易清單
xls1 = pd.ExcelFile(file1_path)
df1 = xls1.parse(sheet_name="交易清單")

# 讀取波動資料（含 zigzag 與結構）
df2 = pd.read_csv(file2_path)
df1['日期/時間'] = pd.to_datetime(df1['日期/時間'])
df2['utc_time'] = pd.to_datetime(df2['utc_time'])

# Step 2: 擷取波動欄位並準備合併（新增 zigzag 與 structure）
df2_trimmed = df2[['utc_time', 'atr', 'rolling_volatility', 'zigzag_point', 'structure']].copy()

# 建立區間分類
atr_bins = pd.qcut(df2_trimmed['atr'], q=4)
vol_bins = pd.qcut(df2_trimmed['rolling_volatility'], q=4)
atr_categories = atr_bins.cat.categories
vol_categories = vol_bins.cat.categories

atr_labels = ['低', '中低', '中高', '高']
vol_labels = ['低', '中低', '中高', '高']

def get_bin_label(value, bin_edges, labels):
    for i, interval in enumerate(bin_edges):
        if value in interval:
            return labels[i]
    return None

# Step 3: 合併交易資料（以波動表為主體）
df_merged = pd.merge(
    df2_trimmed,
    df1[['日期/時間', '獲利 USD', '種類', '交易 #']],
    left_on='utc_time',
    right_on='日期/時間',
    how='left'
)

# Step 4: 區間標記
df_merged['ATR區間'] = df_merged['atr'].apply(lambda x: get_bin_label(x, atr_categories, atr_labels))
df_merged['報酬波動區間'] = df_merged['rolling_volatility'].apply(lambda x: get_bin_label(x, vol_categories, vol_labels))

# Step 5: 淨資產變化（出場累計）
df_merged['淨資產變化'] = df_merged['獲利 USD'].where(df_merged['種類'].str.contains("出場", na=False))
df_merged['淨資產變化'] = df_merged['淨資產變化'].cumsum()

# Step 6: 欄位整理
final_df = df_merged[['atr', 'utc_time', 'rolling_volatility', '獲利 USD', '種類', '交易 #',
                      'ATR區間', '報酬波動區間', '淨資產變化', 'zigzag_point', 'structure']]

# Step 7: 匯出
output_path = "C:/Users/user/OneDrive/桌面/TV/SP500/策略波動整合結果+新指標.xlsx"
final_df.to_excel(output_path, index=False)

print("✅ 完整整合含 ZigZag 結構完成，已儲存為：", output_path)
