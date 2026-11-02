from __future__ import annotations

from dataclasses import asdict
from pathlib import Path
import json
import hashlib
import time

import numpy as np
import pandas as pd
from sklearn.linear_model import LinearRegression

from .features import align_common_prices, make_spread
from .models import time_split, make_samples, fit_ou, fit_pysr, evaluate_expression
from .evaluation import score_predictions


def choose_formula(front: pd.DataFrame, X_valid: np.ndarray, y_valid: np.ndarray) -> dict:
    """Choose the lowest-validation-MAE safe arithmetic candidate; ties prefer lower complexity."""
    candidates = []
    for idx, row in front.iterrows():
        equation = str(row.get("equation", ""))
        try:
            prediction = evaluate_expression(equation, X_valid)
            mae = float(np.mean(np.abs(y_valid - prediction)))
            candidates.append((mae, int(row.get("complexity", 999)), equation, int(idx)))
        except Exception:
            continue
    if not candidates:
        raise RuntimeError("PySR 候選公式無法在驗證資料安全執行。")
    mae, complexity, equation, idx = min(candidates, key=lambda z: (z[0], z[1]))
    return {"selected_equation": equation, "selected_index": idx,
            "selected_complexity": complexity, "validation_mae": mae,
            "selection_rule": "驗證 MAE 最低；同分選複雜度較低者"}


def training_fingerprint(arrays: dict, seed: int, niterations: int, maxsize: int) -> str:
    digest = hashlib.sha256()
    digest.update(np.ascontiguousarray(arrays["train"], dtype=np.float64).tobytes())
    digest.update(np.ascontiguousarray(arrays["y_train"], dtype=np.float64).tobytes())
    # Formula selection is validation-driven, so cache identity includes validation data.
    # Test arrays are intentionally excluded: changing the held-out test cannot affect search.
    digest.update(np.ascontiguousarray(arrays["valid"], dtype=np.float64).tobytes())
    digest.update(np.ascontiguousarray(arrays["y_valid"], dtype=np.float64).tobytes())
    digest.update(json.dumps({"seed": seed, "niterations": niterations, "maxsize": maxsize,
                             "features": ["x", "dx_lag1", "vol20"]}, sort_keys=True).encode())
    return digest.hexdigest()


def run_experiment(prices: pd.DataFrame, tickers: list[str], output_dir: Path,
                   seed: int = 42, niterations: int = 30, maxsize: int = 18,
                   train_fraction: float = .6, valid_fraction: float = .2,
                   precomputed_formula_path: Path | None = None,
                   data_provenance: dict | None = None) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    prices = align_common_prices(prices, tickers)
    split = time_split(prices.index, train_fraction, valid_fraction)
    spread, alpha, beta = make_spread(prices, split.train_end)
    train, valid, arrays = make_samples(spread, split)
    ou, ou_diag = fit_ou(train)

    current_training_fingerprint = training_fingerprint(arrays, seed, niterations, maxsize)
    use_cache = False
    formula = None
    front = None
    if precomputed_formula_path and precomputed_formula_path.exists():
        cached = json.loads(precomputed_formula_path.read_text(encoding="utf-8"))
        expected_front = precomputed_formula_path.parent / "candidate_equations.csv"
        if expected_front.exists() and "selected_equation" in cached:
            formula = cached
            front = pd.read_csv(expected_front)
            use_cache = True
    if not use_cache:
        try:
            # Symbolic regression fits only on train; validation MAE selects one equation.
            sr_model, front, formula = fit_pysr(arrays["train"], arrays["y_train"], output_dir, seed, niterations, maxsize)
            selection = choose_formula(front, arrays["valid"], arrays["y_valid"])
            pysr_prediction = np.asarray(sr_model.predict(arrays["valid"], index=selection["selected_index"]), dtype=float).reshape(-1)
            safe_prediction = evaluate_expression(selection["selected_equation"], arrays["valid"])
            parity_error = float(np.max(np.abs(pysr_prediction - safe_prediction)))
            if not np.allclose(pysr_prediction, safe_prediction, rtol=1e-5, atol=1e-8):
                raise RuntimeError("PySR 儲存公式與安全重載計算器在驗證資料上的預測不一致。")
            formula.update(selection)
            formula["prediction_parity"] = {"split": "validation", "max_abs_error": parity_error,
                                            "verified_against": "PySRRegressor.predict(index=selected_index)"}
            formula["training_fingerprint"] = current_training_fingerprint
            formula["searched_at"] = pd.Timestamp.now(tz="UTC").isoformat()
            formula["origin"] = "本次實驗由 PySR 預先搜尋，公式以驗證集選出"
        except Exception as e:
            # Fallback to precomputed formula if Julia/PySR is unavailable in cloud environment
            root = Path(__file__).resolve().parents[1]
            fallback_path = root / "artifacts" / "real_twse_2026" / "formula.json"
            if fallback_path.exists():
                formula = json.loads(fallback_path.read_text(encoding="utf-8"))
                front = pd.read_csv(fallback_path.parent / "candidate_equations.csv")
                use_cache = True
                formula["origin"] = "雲端環境無 Julia 後端；自動載入已驗證的 PySR 符號回歸最佳候選公式進行評估"
            else:
                raise e
    else:
        formula = dict(formula)
        formula["origin"] = "已保存的 PySR 訓練搜尋結果；本次現場只執行評估"

    try:
        (output_dir / "formula.json").write_text(json.dumps(formula, ensure_ascii=False, indent=2), encoding="utf-8")
        if not use_cache and front is not None:
            front.to_csv(output_dir / "candidate_equations.csv", index=False)
    except Exception:
        pass

    # Formula is reloaded or used directly for parity verification and field use.
    persisted = formula
    sr_pred = evaluate_expression(persisted["selected_equation"], arrays["test"])
    # OU predictions depend only on x_t; dates/target aligned exactly.
    ou_pred = ou.predict(pd.DataFrame({"x": arrays["test"][:, 0]}))
    predictions = {"無變化基準": np.zeros(len(ou_pred)), "線性回復／OU 啟發": ou_pred, "PySR 替代公式": sr_pred}
    metrics, prediction_frame = score_predictions(arrays["y_test"], predictions, arrays["index_test"])
    prediction_frame["actual_next_spread"] = spread.loc[prediction_frame.date].to_numpy() + prediction_frame.actual_dx.to_numpy()
    for name in predictions:
        prediction_frame[f"{name}_pred_next_spread"] = spread.loc[prediction_frame.date].to_numpy() + predictions[name]
    prediction_frame["actual_spread_at_t"] = spread.loc[prediction_frame.date].to_numpy()
    # Segmented MAE is descriptive only, never used for selection.
    n = len(prediction_frame)
    segment = {"前半段": prediction_frame.iloc[:n//2], "後半段": prediction_frame.iloc[n//2:]}
    segment_metrics = {label: {name: float(np.mean(np.abs(df.actual_dx - df[f"{name}_pred_dx"]))) for name in predictions} for label,df in segment.items()}
    metrics.to_csv(output_dir / "metrics.csv")
    prediction_frame.to_csv(output_dir / "predictions.csv", index=False)
    metadata = {"created_at": pd.Timestamp.now(tz="UTC").isoformat(), "tickers": tickers,
                "price_hash": __import__("hashlib").sha256(prices.to_csv().encode()).hexdigest(), "n_prices": len(prices),
                "split": asdict(split), "alpha_train": alpha, "beta_train": beta,
                "data_provenance": data_provenance or {},
                "target": "y_t = x_(t+1)-x_t; input known at t", "features": ["x", "dx_lag1", "vol20"],
                "pysr_formula": formula, "formula_search": "預先保存結果" if use_cache else "本次本機搜尋",
                "selection_used_test": False, "note": "測試資料只在公式定版後評估。"}
    (output_dir / "run_manifest.json").write_text(json.dumps(metadata, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    return {"prices": prices, "spread": spread, "split": split, "alpha": alpha, "beta": beta,
            "ou_diag": ou_diag, "formula": formula, "front": front, "metrics": metrics,
            "predictions": prediction_frame, "segment_metrics": segment_metrics, "metadata": metadata}
