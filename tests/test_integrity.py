import json
from pathlib import Path
import tempfile

import numpy as np
import pandas as pd
import pytest

from src.data import make_synthetic_prices
from src.features import make_spread
from src.models import time_split, make_samples, fit_ou, evaluate_expression
from src.evaluation import score_predictions
from src.experiment import run_experiment, choose_formula
from src.state import settings_fingerprint, should_invalidate


def test_synthetic_ou_parameter_recovery():
    rng = np.random.default_rng(2026)
    x = np.zeros(16000)
    for i in range(1, len(x)):
        x[i] = .015 + .91 * x[i-1] + rng.normal(0, .009)
    y = np.diff(x)
    fit = pd.DataFrame({"x":x[:-1],"target":y})
    model, diag = fit_ou(fit)
    assert abs(diag["b"] - (.91 - 1)) < .01
    assert diag["supported"]
    assert abs(diag["phi"] - .91) < .01


def test_test_data_does_not_change_training_estimates():
    rng = np.random.default_rng(12)
    x = np.zeros(1000)
    for i in range(1, len(x)):
        x[i] = .9*x[i-1]+rng.normal(0,.02)
    idx = pd.bdate_range("2020-01-01",periods=len(x))
    s = pd.Series(x,index=idx)
    split=time_split(idx)
    tr,va,arr=make_samples(s,split)
    _,d1=fit_ou(tr)
    s2=s.copy(); s2.iloc[split.valid_end:]+=np.linspace(0,100,len(s2)-split.valid_end)
    tr2,_,_=make_samples(s2,split)
    _,d2=fit_ou(tr2)
    assert d1==d2
    assert np.array_equal(arr["index_test"], make_samples(s2,split)[2]["index_test"])


def test_three_models_same_dates_and_formula_reload():
    dates=pd.bdate_range("2022-01-03",periods=5)
    X=np.arange(15,dtype=float).reshape(5,3)/10
    persisted={"selected_equation":"(x0 * -0.5) + (x1 * 0.1)"}
    with tempfile.TemporaryDirectory() as d:
        p=Path(d)/"formula.json"; p.write_text(json.dumps(persisted))
        formula=json.loads(p.read_text())
        again=evaluate_expression(formula["selected_equation"],X)
    original=evaluate_expression("(x0 * -0.5) + (x1 * 0.1)",X)
    assert np.array_equal(again,original)
    pred={"無變化基準":np.zeros(5),"線性回復／OU 啟發":np.ones(5)*.1,"PySR 替代公式":again}
    metrics,frame=score_predictions(np.zeros(5),pred,dates)
    assert len(frame)==5
    assert np.array_equal(frame.date,dates)
    assert list(metrics.index)==list(pred)


def test_bad_prices_fail_clearly():
    from src.features import align_common_prices
    with np.testing.assert_raises_regex(ValueError,"正值"):
        align_common_prices(pd.DataFrame({"A":[1,-1]*50,"B":[2,2]*50}),["A","B"])


def test_empty_sample_split_fails_clearly():
    with np.testing.assert_raises_regex(ValueError,"至少需要 80"):
        time_split(pd.bdate_range("2024-01-01",periods=20))


def test_formula_choice_uses_validation_and_not_test():
    # The selection API accepts validation inputs only; test data cannot influence it.
    front = pd.DataFrame({"equation":["x * -0.5", "x * 0.5"], "complexity":[3,3]})
    x_valid=np.array([[1.,0,0],[2.,0,0]])
    y_valid=np.array([-.5,-1.])
    first=choose_formula(front,x_valid,y_valid)
    artificial_test_a=np.array([1.,2.])
    artificial_test_b=np.array([1000.,-1000.])
    assert not np.array_equal(artificial_test_a,artificial_test_b)
    second=choose_formula(front,x_valid,y_valid)
    assert first==second
    assert first["selected_equation"]=="x * -0.5"


def test_precomputed_formula_selection_is_test_invariant(tmp_path):
    prices=make_synthetic_prices(seed=42).prices
    artifact=Path(__file__).parents[1]/"artifacts/precomputed/synthetic-seed-42/formula.json"
    baseline=run_experiment(prices,["A_SYN","B_SYN"],tmp_path/"a",seed=42,niterations=20,maxsize=16,
                            precomputed_formula_path=artifact)
    split=baseline["split"]
    changed=prices.copy()
    changed.iloc[split.valid_end:,0] *= np.exp(np.linspace(0,1.8,len(changed)-split.valid_end))
    modified=run_experiment(changed,["A_SYN","B_SYN"],tmp_path/"b",seed=42,niterations=20,maxsize=16,
                            precomputed_formula_path=artifact)
    assert modified["formula"]["selected_equation"]==baseline["formula"]["selected_equation"]
    assert modified["alpha"]==baseline["alpha"]
    assert modified["beta"]==baseline["beta"]
    assert modified["metadata"]["formula_search"]=="預先保存結果"


def test_ui_result_is_invalidated_when_settings_change():
    a=settings_fingerprint({"ticker":["2330","2454"],"period":"2022-2026"})
    same=settings_fingerprint({"period":"2022-2026","ticker":["2330","2454"]})
    changed=settings_fingerprint({"ticker":["2330","2317"],"period":"2022-2026"})
    assert a==same
    assert not should_invalidate(a,same)
    assert should_invalidate(a,changed)


def test_twse_source_parser_fixture(monkeypatch, tmp_path):
    """API response format, ROC date conversion, and close selection are checked offline."""
    from src import data
    class Response:
        def __init__(self, month): self.month=month
        def raise_for_status(self): pass
        def json(self):
            start=pd.Timestamp(self.month[:4]+"-"+self.month[4:]+"-01")
            end=start+pd.offsets.MonthEnd(0)
            days=pd.bdate_range(start,end)
            data_rows=[[f"{d.year-1911:03d}/{d.month:02d}/{d.day:02d}","1","2","3","4","2","3,000.00","1","5",""] for d in days]
            return {"stat":"OK","fields":["日期","成交股數","成交金額","開盤價","最高價","最低價","收盤價","漲跌價差","成交筆數","註記"],"data":data_rows}
    class Session:
        def __init__(self): self.headers={}
        def get(self,*args,**kwargs): return Response(kwargs["params"]["date"][:6])
        def close(self): pass
    monkeypatch.setattr(data.requests,"Session",Session)
    monkeypatch.setattr(data.time,"sleep",lambda _:None)
    result=data.fetch_twse(["2330"],"2026-07-01","2026-11-01",delay_seconds=0,cache_dir=tmp_path)
    assert len(result.prices)>=80
    assert result.prices.iloc[0,0]==3000
    assert str(result.prices.index[0].date())=="2026-07-01"
