# TV-backtest-anl-tools
整理trading view的策略分析報告
多策略組合回測：支援合併多個 Excel 交易清單，計算組合淨值曲線與風險指標。

策略相關性矩陣：分析不同策略間的盈虧相關性，降低帳戶回撤同步風險。

進階績效評估：計算 CAGR、MAR、Sortino Ratio、Omega Ratio 等專業機構級指標。

市場結構分析：自動標註 ZigZag 轉折點與 HH/LL (Higher High / Lower Low) 趨勢結構。

環境因子關聯：將交易紀錄與滾動波動率 (Rolling Volatility) 及 CNN 貪婪恐懼指數整合，尋找策略失效的邊界條件。


1. 策略組合與績效分析
TV投資組合.py:

自動讀取指定資料夾內所有策略清單。

計算組合的每月盈虧、滾動夏普率 (Rolling Sharpe Ratio)。

輸出包含淨值曲線圖的 Excel 報告。

TV策略評估.py:

針對單一策略進行深度「健檢」。

提供微型期貨版策略評分系統，綜合考量 CAGR、Sortino、MAR 等權重。

TV每月盈虧+相關性矩陣.py:

分析各策略在不同時間維度（周、月）的相關性。

細分「做多」與「做空」的交易次數相關性，優化策略多樣性。

2. 資料處理與市場環境因子
合併K棒資料_加UTC時間.py:

整合 CSV (新資料) 與 XLSX (舊資料)，自動去重並統一轉換為 UTC 時間戳。

TV滾動波動率.py:

計算市場的滾動波動率，並透過 ZigZag 演算法判斷當前市場狀態（Uptrend / Downtrend / Ranging）。

波動率因子比較.py:

將交易損益與 ATR、波動率區間進行交叉比對，找出最適合策略運行的波動環境。

貪婪恐懼指數.py:

爬取 CNN Fear & Greed Index 歷史數據，建立情緒面分析資料庫。


語言: Python 3.x

資料處理: pandas, numpy

圖表繪製: matplotlib

檔案操作: openpyxl (Excel 整合), glob, os

爬蟲/API: requests, fake_useragent
