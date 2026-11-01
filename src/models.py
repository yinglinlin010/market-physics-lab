from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path
import re

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression


@dataclass
class Split:
    train_end: int
    valid_end: int
    n: int
    train_dates: tuple[str, str]
    valid_dates: tuple[str, str]
    test_dates: tuple[str, str]


def time_split(index: pd.Index, train_fraction: float = .6, valid_fraction: float = .2) -> Split:
    if train_fraction <= 0 or valid_fraction <= 0 or train_fraction + valid_fraction >= 1:
        raise ValueError("訓練與驗證比例必須大於零，且總和小於 1。")
    n = len(index)
    a, b = int(n * train_fraction), int(n * (train_fraction + valid_fraction))
    if a < 40 or b - a < 10 or n - b < 10:
        raise ValueError("時間切分後每段資料不足；至少需要 80 個有效交易日。")
    return Split(a, b, n,
                 (str(index[0].date()), str(index[a-1].date())),
                 (str(index[a].date()), str(index[b-1].date())),
                 (str(index[b].date()), str(index[-1].date())))


def make_samples(spread: pd.Series, split: Split) -> tuple[pd.DataFrame, pd.Series, dict[str, np.ndarray]]:
    x = spread.to_numpy(dtype=float)
    # targets are x[t+1]-x[t], keyed to the information date t.
    y = np.diff(x)
    frame = pd.DataFrame({"x": x[:-1], "dx_lag1": np.r_[np.nan, y[:-1]]}, index=spread.index[:-1])
    frame["vol20"] = pd.Series(frame.dx_lag1, index=frame.index).rolling(20, min_periods=5).std(ddof=0)
    frame["target"] = y
    # Strict purge: last sample whose label crosses a partition boundary is removed.
    train = frame.iloc[:split.train_end - 1].dropna()
    valid = frame.iloc[split.train_end:split.valid_end - 1].dropna()
    test = frame.iloc[split.valid_end:split.n - 1].dropna()
    if min(len(train), len(valid), len(test)) < 10:
        raise ValueError("切分及前處理後的有效樣本太少。")
    X_cols = ["x", "dx_lag1", "vol20"]
    sets = {"train": train[X_cols].to_numpy(), "valid": valid[X_cols].to_numpy(), "test": test[X_cols].to_numpy()}
    targets = {"train": train.target.to_numpy(), "valid": valid.target.to_numpy(), "test": test.target.to_numpy()}
    return train, valid, {**sets, **{f"y_{k}": v for k,v in targets.items()}, **{f"index_{k}": df.index for k,df in [("train",train),("valid",valid),("test",test)]}}


def fit_ou(train: pd.DataFrame) -> tuple[LinearRegression, dict]:
    fit = LinearRegression().fit(train[["x"]], train.target)
    a, b = float(fit.intercept_), float(fit.coef_[0])
    phi = 1 + b
    diag = {"a": a, "b": b, "phi": phi, "supported": bool(b < 0 and 0 < phi < 1), "kappa": None, "half_life": None, "mu": None}
    if diag["supported"]:
        kappa = -float(np.log(phi))
        diag.update(kappa=kappa, half_life=float(np.log(2) / kappa), mu=float(a / (1 - phi)))
    return fit, diag


def fit_pysr(X: np.ndarray, y: np.ndarray, out_dir: Path, seed: int = 42, niterations: int = 30, maxsize: int = 18):
    try:
        from pysr import PySRRegressor
    except ImportError as exc:
        raise RuntimeError("PySR 未安裝；離線示範需要先安裝 requirements.txt 中的 PySR。") from exc
    out_dir.mkdir(parents=True, exist_ok=True)
    model = PySRRegressor(binary_operators=["+", "-", "*"], unary_operators=[], niterations=niterations,
                          maxsize=maxsize, deterministic=True, random_state=seed, parallelism="serial",
                          model_selection="best", verbosity=0,
                          output_directory=str(out_dir / "pysr"), run_id="midterm-formula")
    model.fit(X, y, variable_names=["x", "dx_lag1", "vol20"])
    equations = model.equations_.copy()
    equations.to_csv(out_dir / "candidate_equations.csv", index=False)
    # Selection happens in experiment.py using validation MAE, never the test split.
    payload = {"features": ["x", "dx_lag1", "vol20"], "seed": seed,
               "niterations": niterations, "maxsize": maxsize,
               "pysr_version": getattr(__import__("pysr"), "__version__", "unknown"),
               "binary_operators": ["+", "-", "*"], "unary_operators": [],
               "deterministic": True, "parallelism": "serial",
               "preprocessing": "none; natural units; features x, dx_lag1, vol20"}
    return model, equations, payload


def evaluate_expression(expression: str, X: np.ndarray) -> np.ndarray:
    """Evaluate the limited arithmetic grammar emitted for this project."""
    allowed = re.compile(r"^[\w\s.+*()\-/]+$")
    if not allowed.fullmatch(expression) or "__" in expression:
        raise ValueError("公式包含安全計算器不支援的語法。")
    x0, x1, x2 = X[:, 0], X[:, 1], X[:, 2]
    value = eval(expression, {"__builtins__": {}}, {
        "x0": x0, "x1": x1, "x2": x2,
        "x": x0, "dx_lag1": x1, "vol20": x2,
    })
    result = np.asarray(value, dtype=float).reshape(-1)
    if len(result) != len(X) or not np.isfinite(result).all():
        raise ValueError("公式輸出含非有限值或維度不符。")
    return result
