from pathlib import Path
import os

from src.data import make_synthetic_prices
from src.experiment import run_experiment


ROOT = Path(__file__).resolve().parents[1]


def main() -> None:
    # Keep Julia threads deterministic and prevent an interactive REPL prompt.
    os.environ.setdefault("PYTHON_JULIACALL_THREADS", "auto")
    os.environ.setdefault("PYTHON_JULIACALL_AUTOLOAD_IPYTHON_EXTENSION", "no")
    prices = make_synthetic_prices(seed=42)
    result = run_experiment(
        prices.prices,
        ["A_SYN", "B_SYN"],
        ROOT / "artifacts" / "precomputed" / "synthetic-seed-42",
        seed=42,
        niterations=20,
        maxsize=16,
        train_fraction=.60,
        valid_fraction=.20,
        data_provenance={"source": prices.source, "retrieved_at": prices.retrieved_at,
                         "adjustment": prices.adjustment, "version": prices.version,
                         "synthetic": True},
    )
    print("已保存候選公式、選式、測試預測與實驗紀錄。")
    print(result["formula"]["selected_equation"])
    print(result["metrics"].to_string())


if __name__ == "__main__":
    main()
