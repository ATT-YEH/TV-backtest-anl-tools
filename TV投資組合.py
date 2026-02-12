import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
import glob
import os

# 設定無風險年利率（轉換為月利率）
risk_free_rate_annual = 0.02  # 2%
risk_free_rate_monthly = risk_free_rate_annual / 12

# 讀取所有 Excel 檔案
file_paths = glob.glob(r"C:\Users\user\OneDrive\桌面\TV\ALL\*.xlsx")
strategies = {}

for file_path in file_paths:
    strategy_name = os.path.basename(file_path).replace(".xlsx", "")
    xls = pd.ExcelFile(file_path)
    sheet_name = "交易清單"

    if sheet_name not in xls.sheet_names:
        print(f"⚠️ {file_path} 沒有找到工作表: {sheet_name}，跳過此檔案")
        continue

    df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
    df = df[df["種類"].str.contains("出場")]
    df["日期/時間"] = pd.to_datetime(df["日期/時間"])
    df["獲利 USD"] = pd.to_numeric(df["獲利 USD"], errors="coerce").fillna(0)
    strategies[strategy_name] = df

if not strategies:
    print("❌ 沒有找到任何可用的交易清單，請檢查 Excel 檔案！")
    exit()

# 生成組合數據
portfolio = pd.concat(strategies.values(), ignore_index=True)
portfolio["月份"] = portfolio["日期/時間"].dt.to_period("M")
portfolio_pnl = portfolio.groupby("月份")["獲利 USD"].sum()


# 計算投資組合的風險指標
def max_drawdown(equity_curve):
    peak = equity_curve.cummax()
    drawdown = equity_curve - peak
    return drawdown.min()


def omega_ratio(returns):
    gains = returns[returns > 0].sum()
    losses = -returns[returns < 0].sum()
    return gains / losses if losses != 0 else np.nan


def sortino_ratio(returns, risk_free=risk_free_rate_monthly):
    downside_std = returns[returns < 0].std()
    mean_return = returns.mean() - risk_free
    return mean_return / downside_std if downside_std != 0 else np.nan


def rolling_sharpe(returns, window=3):
    return (returns.rolling(window).mean() - risk_free_rate_monthly) / returns.rolling(window).std()


# 計算淨值曲線
equity_curve = portfolio_pnl.cumsum()

# 計算各種風險指標
total_net_profit = portfolio_pnl.sum()
mdd = max_drawdown(equity_curve)
omega = omega_ratio(portfolio_pnl)
sortino = sortino_ratio(portfolio_pnl)
rolling_sharpe_ratio = rolling_sharpe(portfolio_pnl, window=3)
rolling_sharpe_ratio.index = rolling_sharpe_ratio.index.strftime('%Y-%m')
rolling_sharpe_df = rolling_sharpe_ratio.rename_axis("月份").reset_index(name="滾動夏普率")

# 繪製淨值曲線
plt.figure(figsize=(12, 6))
plt.plot(equity_curve.index.astype(str), equity_curve, marker='o', label="投資組合淨值")
plt.xlabel("月份")
plt.ylabel("淨值 (USD)")
plt.title("投資組合淨值曲線")
plt.legend()
plt.grid(True)
plt.xticks(rotation=90)
plt.savefig("equity_curve.png")
plt.close()

# 輸出到 Excel
output_path = r"C:\Users\user\OneDrive\桌面\TV\投資組合分析.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    portfolio_pnl.to_excel(writer, sheet_name="每月盈虧")
    rolling_sharpe_df.to_excel(writer, sheet_name="滾動夏普率", index=False)
    summary_df = pd.DataFrame({"指標": ["總淨利", "最大回撤（USD）", "Omega Ratio", "Sortino Ratio"],
                               "數值": [total_net_profit, mdd, omega, sortino]})
    summary_df.to_excel(writer, sheet_name="風險指標", index=False)

# 插入淨值曲線圖到 Excel
workbook = openpyxl.load_workbook(output_path)
worksheet = workbook.create_sheet("淨值曲線")
img = openpyxl.drawing.image.Image("equity_curve.png")
worksheet.add_image(img, "A1")
workbook.save(output_path)

print(f"✅ Excel 文件已輸出到: {output_path}")
