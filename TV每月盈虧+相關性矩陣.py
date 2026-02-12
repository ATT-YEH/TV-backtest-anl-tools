import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import glob
import os
import openpyxl
from openpyxl.drawing.image import Image

# 設定 Matplotlib 中文字體
plt.rcParams['font.sans-serif'] = ['SimHei']  # 或 'Microsoft YaHei'
plt.rcParams['axes.unicode_minus'] = False  # 避免負號顯示錯誤

# 📌 讀取多個交易策略的交易明細
file_paths = glob.glob(r"C:\Users\user\OneDrive\桌面\TV\SP500\實單\三月\*.xlsx")
strategies = {}

for file_path in file_paths:
    strategy_name = file_path.split("\\")[-1].replace(".xlsx", "")
    xls = pd.ExcelFile(file_path)
    sheet_name = "交易清單"
    if sheet_name not in xls.sheet_names:
        print(f"⚠️ {file_path} 沒有找到工作表: {sheet_name}，跳過此檔案")
        continue

    df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
    df = df[df["種類"].str.contains("出場")]
    if "日期/時間" in df.columns:
        df["日期/時間"] = pd.to_datetime(df["日期/時間"])
    else:
        print(f"⚠️ {file_path} 沒有 '日期/時間' 欄位，跳過此檔案")
        continue

    strategies[strategy_name] = df

if not strategies:
    print("❌ 沒有找到任何可用的交易清單，請檢查 Excel 檔案！")
    exit()

# 📌 1. 計算每個策略的「每月盈虧報表」
monthly_pnl = {}
for name, df in strategies.items():
    df["月份"] = df["日期/時間"].dt.to_period("M")
    monthly_pnl[name] = df.groupby("月份")["獲利 USD"].sum()

df_monthly_pnl = pd.DataFrame(monthly_pnl).fillna(0)
correlation_monthly = df_monthly_pnl.corr()

# 📌 2. 計算每個策略的「每週盈虧報表」
weekly_pnl = {}
for name, df in strategies.items():
    df["周"] = df["日期/時間"].dt.to_period("W")
    weekly_pnl[name] = df.groupby("周")["獲利 USD"].sum()

df_weekly_pnl = pd.DataFrame(weekly_pnl).fillna(0)
correlation_weekly = df_weekly_pnl.corr()

# 📌 3. 計算「做多 / 做空」交易數量相關性
long_pnl = {}  # 做多
short_pnl = {}  # 做空

for name, df in strategies.items():
    df["月份"] = df["日期/時間"].dt.to_period("M")

    # **做多交易數量**
    long_trades = df[df["種類"] == "出場做多"].groupby("月份").size()
    long_pnl[name] = long_trades

    # **做空交易數量**
    short_trades = df[df["種類"] == "出場做空"].groupby("月份").size()
    short_pnl[name] = short_trades

# **合併數據**
df_long_pnl = pd.DataFrame(long_pnl).fillna(0)  # 做多數據
df_short_pnl = pd.DataFrame(short_pnl).fillna(0)  # 做空數據

# 計算相關性
correlation_long = df_long_pnl.corr()  # 做多交易數量的相關性
correlation_short = df_short_pnl.corr()  # 做空交易數量的相關性

# 📌 4. 繪製「每月盈虧」折線圖
plt.figure(figsize=(12, 5))  # 調整圖表寬度，避免擠在一起
for name in df_monthly_pnl.columns:
    plt.plot(df_monthly_pnl.index.astype(str), df_monthly_pnl[name], marker='o', label=name)

plt.xlabel("月份")
plt.ylabel("盈虧 (USD)")
plt.title("每個策略的每月盈虧變化")
plt.legend()
plt.grid(True)

# ✅ **調整 X 軸標籤**
plt.xticks(rotation=90, fontsize=10)  # **旋轉 X 軸標籤 90 度，並調整字體大小**
plt.savefig("盈虧變化.png")  # 儲存圖片
plt.close()

# 📌 5. 輸出 Excel 檔案
output_path = r"C:\Users\user\OneDrive\桌面\TV\交易數據分析.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df_monthly_pnl.to_excel(writer, sheet_name="每月盈虧數據")
    correlation_monthly.to_excel(writer, sheet_name="每月盈虧相關性")
    df_weekly_pnl.to_excel(writer, sheet_name="每週盈虧數據")
    correlation_weekly.to_excel(writer, sheet_name="每週盈虧相關性")
    df_long_pnl.to_excel(writer, sheet_name="做多交易數量")
    correlation_long.to_excel(writer, sheet_name="做多交易數量相關性")
    df_short_pnl.to_excel(writer, sheet_name="做空交易數量")
    correlation_short.to_excel(writer, sheet_name="做空交易數量相關性")

# 📌 6. 插入折線圖到 Excel
workbook = openpyxl.load_workbook(output_path)
worksheet = workbook.create_sheet("折線圖")
img = Image("盈虧變化.png")
worksheet.add_image(img, "A1")
workbook.save(output_path)

print(f"✅ Excel 文件已輸出到: {output_path}")
