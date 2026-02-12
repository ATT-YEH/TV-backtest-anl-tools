
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl

# === 參數設定 ===
initial_capital = 1606  # ✅ 新增初始資金填入保證金*3
risk_free_rate_annual = 0.02 #輸入無風險利率
risk_free_rate_monthly = risk_free_rate_annual / 12

# === 路徑設定（請自行修改） ===
file_path = r"C:/Users/user/OneDrive/桌面/TV/運行策略 .xlsx"
sheet_name = "交易清單"
output_path = r"C:/Users/user/OneDrive/桌面/TV/標普回測評估+MAR_完整版.xlsx"

# === 讀取資料 ===
xls = pd.ExcelFile(file_path)
df = pd.read_excel(xls, sheet_name=sheet_name, engine="openpyxl")
#以下行測tv原始報表記得打開
#df = df[df["種類"].str.contains("出場")]
df["日期/時間"] = pd.to_datetime(df["日期/時間"])
df["淨損益USD"] = pd.to_numeric(df["淨損益USD"], errors="coerce").fillna(0)

# === 每月盈虧 ===
df["月份"] = df["日期/時間"].dt.to_period("M")
monthly_pnl = df.groupby("月份")["淨損益USD"].sum()

# === 滾動夏普率 ===
def rolling_sharpe(returns, window=3):
    return (returns.rolling(window).mean() - risk_free_rate_monthly) / returns.rolling(window).std()

rolling_sharpe_ratio = rolling_sharpe(monthly_pnl, window=3)
rolling_sharpe_ratio.index = rolling_sharpe_ratio.index.strftime('%Y-%m')
rolling_sharpe_ratio_df = rolling_sharpe_ratio.rename_axis("月份").reset_index(name="滾動夏普率")

# === 績效指標 ===
def omega_ratio(returns):
    gains = returns[returns > 0].sum()
    losses = -returns[returns < 0].sum()
    return gains / losses if losses != 0 else np.nan

def sortino_ratio(returns, risk_free=risk_free_rate_monthly):
    downside_std = returns[returns < 0].std()
    mean_return = returns.mean() - risk_free
    return mean_return / downside_std if downside_std != 0 else np.nan

def calculate_mar(returns, capital):
    returns = returns.reset_index(drop=True)
    equity_curve = returns.cumsum() + capital
    peak = equity_curve.cummax()
    drawdown = peak - equity_curve
    max_drawdown = drawdown.max()
    total_return = equity_curve.iloc[-1] - capital
    n_years = len(returns) / 12
    cagr = (equity_curve.iloc[-1] / capital) ** (1 / n_years) - 1
    return cagr / max_drawdown if max_drawdown != 0 else np.nan

def calculate_ret_dd(returns, capital):
    returns = returns.reset_index(drop=True)
    equity_curve = returns.cumsum() + capital
    peak = equity_curve.cummax()
    drawdown = peak - equity_curve
    max_drawdown = drawdown.max()
    total_return = equity_curve.iloc[-1] - capital
    return total_return / max_drawdown if max_drawdown != 0 else np.nan

# === 計算指標 ===
returns = monthly_pnl.reset_index(drop=True)
equity_curve = returns.cumsum() + initial_capital
n_years = len(returns) / 12
cagr = (equity_curve.iloc[-1] / initial_capital) ** (1 / n_years) - 1

omega = omega_ratio(monthly_pnl)
sortino = sortino_ratio(monthly_pnl)
mar = calculate_mar(monthly_pnl, initial_capital)
ret_dd = calculate_ret_dd(monthly_pnl, initial_capital)

# === 策略評分（針對微型期貨）===
score = (
    (cagr * 100) * 0.4 +
    sortino * 0.3 +
    omega * 0.2 +
    ret_dd * 0.1
)

# === 匯出 Excel ===
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    monthly_pnl.to_excel(writer, sheet_name="每月盈虧")
    rolling_sharpe_ratio_df.to_excel(writer, sheet_name="滾動夏普率", index=False)
    summary_df = pd.DataFrame({
        "指標": ["CAGR", "Omega Ratio", "Sortino Ratio", "MAR", "RET/DD", "策略評分 (微型期貨版)"],
        "數值": [cagr, omega, sortino, mar, ret_dd, score]
    })
    summary_df.to_excel(writer, sheet_name="風險指標", index=False)

print(f"✅ Excel 文件已輸出到: {output_path}")
