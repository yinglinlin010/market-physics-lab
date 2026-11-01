from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error


def score_predictions(actual: np.ndarray, predicted: dict[str, np.ndarray], dates: pd.Index) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not predicted:
        raise ValueError("沒有可評估的預測結果。")
    lengths = {len(actual), len(dates), *(len(v) for v in predicted.values())}
    if len(lengths) != 1:
        raise ValueError("模型預測、實際值與日期長度不一致。")
    rows = []
    for name, values in predicted.items():
        if not np.isfinite(values).all():
            raise ValueError(f"{name} 預測包含 NaN 或無限值。")
        rows.append({"模型": name, "MAE": mean_absolute_error(actual, values), "RMSE": np.sqrt(mean_squared_error(actual, values))})
    metrics = pd.DataFrame(rows).set_index("模型")
    baseline = metrics.loc["無變化基準", "MAE"]
    metrics["相對基準改善"] = (baseline - metrics.MAE) / baseline if baseline > 0 else np.nan
    predictions = pd.DataFrame({"date": dates, "actual_dx": actual, **{f"{n}_pred_dx": v for n,v in predicted.items()}})
    return metrics, predictions
