from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression


def align_common_prices(prices: pd.DataFrame, tickers: list[str]) -> pd.DataFrame:
    missing = [t for t in tickers if t not in prices.columns]
    if missing:
        raise ValueError(f"缺少股票欄位：{', '.join(missing)}")
    result = prices[tickers].copy().sort_index().dropna(how="any")
    result = result[~result.index.duplicated(keep="last")]
    if len(result) < 80:
        raise ValueError("有效共同交易日少於 80 筆，請延長期間或選擇資料較完整的股票。")
    if not np.isfinite(result.to_numpy()).all() or (result.to_numpy() <= 0).any():
        raise ValueError("價格包含非有限值或非正值。")
    return result


def make_spread(prices: pd.DataFrame, train_end_pos: int) -> tuple[pd.Series, float, float]:
    logs = np.log(prices)
    x_train = logs.iloc[:train_end_pos]
    if len(x_train) < 30:
        raise ValueError("訓練資料不足 30 筆，無法估計價差係數。")
    model = LinearRegression().fit(x_train.iloc[:, [1]], x_train.iloc[:, 0])
    alpha, beta = float(model.intercept_), float(model.coef_[0])
    spread = logs.iloc[:, 0] - alpha - beta * logs.iloc[:, 1]
    spread.name = "spread"
    return spread, alpha, beta


def supervised_frame(spread: pd.Series, train_end_pos: int) -> pd.DataFrame:
    # x_t predicts y_t = x_(t+1)-x_t. The date index is t (information available then).
    x = spread.to_numpy()
    y = np.diff(x)
    dates = spread.index[:-1]
    dx = np.r_[np.nan, y[:-1]]
    vol = pd.Series(dx, index=dates).rolling(20, min_periods=5).std(ddof=0).to_numpy()
    frame = pd.DataFrame({"x": x[:-1], "dx_lag1": dx, "vol20": vol, "target": y}, index=dates)
    # train_end_pos is a raw-date position. Keep target t only when t+1 remains
    # in the same partition. NaN rolling values are excluded, fit only on training.
    return frame.replace([np.inf, -np.inf], np.nan)
