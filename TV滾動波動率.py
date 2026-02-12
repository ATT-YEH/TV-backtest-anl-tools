import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os

# 計算滾動波動率
def calculate_rolling_volatility(df, close_column, window_size=5):
    df[close_column] = pd.to_numeric(df[close_column], errors='coerce')
    df['returns'] = df[close_column].pct_change() * 100
    df['rolling_volatility'] = df['returns'].rolling(window=window_size).std() * np.sqrt(24)
    return df

# 將 Unix timestamp 轉換為 UTC
def convert_unix_to_utc(df, time_column):
    df['utc_time'] = pd.to_datetime(df[time_column], unit='s', utc=True)
    df['utc_time'] = df['utc_time'].dt.strftime('%Y/%m/%d %H:%M')
    return df

# ZigZag 與 HH/LL 結構判斷 # ZigZag 閾值 1D 用 2%、4H 用 1%、1H 用 0.2%
def calculate_zigzag_and_structure(df, close_column, threshold=0.002):
    close = df[close_column].values
    timestamps = df['utc_time'].values
    zigzag = [np.nan] * len(close)

    last_pivot_index = 0
    last_pivot_price = close[0]
    trend = None

    for i in range(1, len(close)):
        change = (close[i] - last_pivot_price) / last_pivot_price
        if trend is None:
            if abs(change) >= threshold:
                trend = 'up' if change > 0 else 'down'
                zigzag[i] = close[i]
                last_pivot_price = close[i]
                last_pivot_index = i
        else:
            if (trend == 'up' and close[i] < last_pivot_price * (1 - threshold)) or \
               (trend == 'down' and close[i] > last_pivot_price * (1 + threshold)):
                trend = 'down' if trend == 'up' else 'up'
                zigzag[i] = close[i]
                last_pivot_price = close[i]
                last_pivot_index = i

    df['zigzag_point'] = zigzag

    # 計算結構狀態 HH/LL 判斷
    structure = ['Unknown'] * len(close)
    zz_idx = [i for i, val in enumerate(zigzag) if not pd.isna(val)]
    for j in range(3, len(zz_idx)):
        p1, p2, p3, p4 = zigzag[zz_idx[j-3]], zigzag[zz_idx[j-2]], zigzag[zz_idx[j-1]], zigzag[zz_idx[j]]
        if p1 < p3 and p2 < p4:
            s = 'Uptrend'
        elif p1 > p3 and p2 > p4:
            s = 'Downtrend'
        else:
            s = 'Ranging'
        structure[zz_idx[j]] = s

    df['structure'] = structure
    return df

def main():
    file_path = "C:/Users/user/OneDrive/桌面/TV/SP500/標普近月1H歷史_210402-250726.xlsx"
    output_path = "C:/Users/user/OneDrive/桌面/TV/SP500/SP500-1h滾動波動率250726測試.csv"
    time_col_name = 'time'
    close_col_name = 'close'
    window_size = 24
    zigzag_threshold_pct = 0.002  # ZigZag 閾值 1D 用 2%、4H 用 1%、1H 用 0.2%

    try:
        print(f"正在讀取 Excel 數據從: {file_path}")
        df = pd.read_excel(file_path)

        print(f"轉換時間欄位 '{time_col_name}' 為 UTC 格式...")
        df = convert_unix_to_utc(df, time_column=time_col_name)

        print(f"計算滾動波動率 (窗口大小: {window_size})")
        result_df = calculate_rolling_volatility(df, close_column=close_col_name, window_size=window_size)

        print(f"計算 ZigZag 與 HH/LL 結構 (閾值: {zigzag_threshold_pct*100:.1f}%)")
        result_df = calculate_zigzag_and_structure(result_df, close_column=close_col_name, threshold=zigzag_threshold_pct)

        print(f"保存結果到: {output_path}")
        result_df.to_csv(output_path, index=False, encoding='utf-8-sig')

        print("\n計算結果摘要:")
        print(result_df[[ 'utc_time', close_col_name, 'returns', 'rolling_volatility', 'zigzag_point', 'structure']].dropna().head(10))
        print("\n處理完成!")

    except FileNotFoundError:
        print(f"錯誤: 找不到檔案 '{file_path}'。請檢查路徑是否正確。")
    except KeyError as e:
        print(f"錯誤: 在 Excel 檔案中找不到指定的欄位 {e}。請檢查 'time_col_name' 和 'close_col_name' 是否設定正確。")
    except Exception as e:
        print(f"發生未預期的錯誤: {e}")

if __name__ == "__main__":
    main()
