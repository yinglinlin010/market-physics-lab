from __future__ import annotations

from datetime import date
from pathlib import Path
import json
import hashlib

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.data import make_synthetic_prices, fetch_twse, load_local_csv_bytes, load_bundled_twse_sample
from src.experiment import run_experiment
from src.state import settings_fingerprint, should_invalidate

ROOT = Path(__file__).parent
ARTIFACTS = ROOT / "artifacts" / "latest"
st.set_page_config(page_title="股票也有回復力嗎？", page_icon="↗", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
:root { --ink:#172A35; --muted:#5F727A; --paper:#F6F4EE; --panel:#FFFFFF; --teal:#137C78; --teal-light:#E9F3F1; --orange:#D16B3C; --line:#DCE2DD; }
html, body, [class*="css"] { font-family:'PingFang TC','Microsoft JhengHei',sans-serif; color:var(--ink); }
.stApp { background:var(--paper); }
[data-testid="stSidebar"] { background:#EAF0ED; border-right:1px solid var(--line); }
[data-testid="stSidebar"] > div:first-child { padding-top:1.6rem; }
[data-testid="stSidebar"] h2 { font-size:18px; }
h1 { letter-spacing:-.035em; font-size:2.1rem !important; font-weight:700 !important; line-height:1.2 !important; }
h2,h3 { letter-spacing:-.02em; }
.block-container { max-width:1360px; padding-top:2.3rem; padding-bottom:4rem; }
[data-testid="stMetric"] { background:var(--panel); border:1px solid var(--line); border-radius:8px; padding:16px 18px; min-height:112px; }
[data-testid="stMetricLabel"] { color:var(--ink); font-weight:600; }
[data-testid="stMetricValue"] { font-variant-numeric:tabular-nums; }
.stButton>button[kind="primary"] { background:var(--teal); border-color:var(--teal); border-radius:8px; min-height:2.8rem; font-weight:600; }
.stButton>button[kind="primary"]:hover { background:#0D625F; border-color:#0D625F; }
.stButton>button:focus-visible, input:focus-visible { outline:3px solid #D16B3C !important; outline-offset:2px; }
.evidence { border:1px solid var(--line); border-radius:8px; padding:16px 18px; background:#fff; line-height:1.9; }
.evidence b { color:var(--ink); font-size:.88rem; font-weight:600; }
.formula { font-family:ui-monospace,SFMono-Regular,Menlo,monospace; font-size:14px; line-height:1.65; color:#134E4B; background:#E9F3F1; border:1px solid #DCE2DD; border-radius:5px; padding:10px 12px; overflow-wrap:anywhere; }
.experiment-note { background:#E9F3F1; color:#134E4B; border:1px solid #DCE2DD; border-radius:8px; padding:12px 15px; font-size:14px; line-height:1.65; }
.experiment-note b { font-weight:650; }
.stAlert { border-radius:8px; }
[data-testid="stExpander"] { border-color:var(--line); border-radius:8px; }
[data-testid="stDataFrame"] { border:1px solid var(--line); border-radius:8px; overflow:hidden; }
::selection { background:#E9F3F1; color:#172A35; }
.small-muted { color:var(--muted); font-size:.9rem; }
hr { border-color:var(--line); }
@media (max-width:700px) {
  h1 {font-size:26px !important;}
  .block-container {padding:1.2rem .85rem 2.5rem;}
  [data-testid="stMetric"] {min-height:96px; padding:12px 13px;}
  .evidence {padding:13px 14px;}
}
</style>
""", unsafe_allow_html=True)


def fmt(x, digits=4):
    return "—" if pd.isna(x) else f"{x:.{digits}f}"

st.title("股票也有回復力嗎？")
st.caption("物理模型與 AI 公式的比較實驗｜以固定時間切分，檢驗股票對數價差的樣本外變化。")
st.markdown("<div class='experiment-note'><b>研究資料與公式已隨專案保存。</b>預設載入 TWSE 官方歷史樣本；PySR 公式預先搜尋，評估於本機執行。</div>", unsafe_allow_html=True)

with st.sidebar:
    st.header("研究設定")
    source = st.radio("資料來源", ["TWSE 官方樣本（離線，2026-05 至 10）", "合成資料（流程驗證）", "TWSE 官方歷史資料（線上查詢）", "本機 CSV"], help="官方樣本已附於專案，可離線重現；亦可用合成資料驗證已知回復機制。")
    if source.startswith("TWSE 官方樣本"):
        tickers = ["2330", "2454"]
        period_label = "2026-05-04 至 2026-10-30（123 個共同交易日）"
        seed = 42
        st.caption("台積電（2330）與聯發科（2454），共 123 個共同交易日；價格調整與來源可於實驗紀錄查看。")
    elif source.startswith("合成"):
        tickers = ["A_SYN", "B_SYN"]
        seed = st.number_input("合成資料種子", 1, 999999, 42)
        period_label = "固定 900 個模擬交易日"
    else:
        default_pair = ["2330", "2454"]
        tickers = st.text_input("兩檔代碼（逗號分隔）", ",".join(default_pair)).upper().replace(" ", "").split(",")[:2]
        if len(tickers) != 2 or not all(tickers):
            st.error("請填入兩個股票代碼。")
            tickers = default_pair
        start = st.date_input("開始日期", date(2026, 5, 1) if source.startswith("TWSE") else date(2022, 1, 1))
        end = st.date_input("結束日期（不含）", date.today())
        period_label = f"{start} 至 {end}"
        seed = 42
    uploaded = None
    if source == "本機 CSV":
        uploaded = st.file_uploader("上傳 date,ticker,close CSV", type=["csv"])
    train_fraction = st.slider("訓練比例", .50, .75, .60, .05)
    valid_fraction = st.slider("驗證比例", .10, .30, .20, .05)
    iterations = st.select_slider("PySR 搜尋迭代", options=[10, 20, 30, 50], value=20, help="只用訓練資料搜尋；展示可載入預先搜尋公式。")
    maxsize = st.slider("公式複雜度上限", 8, 24, 16, 2)
    run = st.button("執行比較", type="primary", use_container_width=True)
    st.caption("官方樣本已保存 PySR 公式，展示時只評估；改資料或搜尋設定會在本機搜尋 PySR。首次搜尋需初始化 Julia。")

settings = {"source": source, "tickers": tickers, "period": period_label, "train": train_fraction, "valid": valid_fraction,
            "seed": int(seed), "iterations": int(iterations), "maxsize": int(maxsize),
            "uploaded_csv_sha256": hashlib.sha256(uploaded.getvalue()).hexdigest() if uploaded else None}
fingerprint = settings_fingerprint(settings)
if "result" not in st.session_state:
    st.session_state.result = None
if should_invalidate(st.session_state.get("result_fingerprint"), fingerprint):
    st.session_state.result = None
    st.session_state.result_fingerprint = None

left, right = st.columns([1.25, 1])
with left:
    st.subheader("研究問題")
    st.markdown("兩檔股票常同受市場與產業消息影響。若其價格長期共同變動，兩者的**對數價差**可能在偏離後出現回復。本研究檢驗線性回復假設與 PySR 搜尋公式，能否比「價差變化為零」更準確預測下一期變化。")
with right:
    st.subheader("本次實驗")
    st.markdown(f"<div class='evidence'><b>資料</b>　{', '.join(tickers)}<br><b>期間</b>　{period_label}<br><b>切分</b>　訓練 {train_fraction:.0%}／驗證 {valid_fraction:.0%}／測試 {1-train_fraction-valid_fraction:.0%}<br><b>評估目標</b>　下一期對數價差變化 Δx<br><b>狀態</b>　{'已完成' if st.session_state.result is not None else '尚未執行'}</div>", unsafe_allow_html=True)

st.markdown("---")

with st.expander("研究背景、目的與分析架構", expanded=True):
    st.markdown("**研究背景。** 股票價格各自波動，但相關公司的價格也可能共享市場或產業趨勢。將兩檔價格取對數並以訓練期估計的線性關係組成價差，可把焦點由個別價格水準轉為相對偏離。若價差高於其估計均衡位置後傾向下降，這是可檢驗的回復假說，而非先驗結論。")
    st.markdown("**研究目的。** 在完全依時間排序的訓練、驗證、測試流程中，比較無變化基準、線性回復／OU 啟發模型，以及 PySR 從訓練資料搜尋並由驗證集選出的替代公式。所有模型預測相同的下一期對數價差變化，並以測試集 MAE、RMSE 與相對基準改善量比較。")
    b1, b2, b3 = st.columns(3)
    b1.markdown("**價差定義**\n\n`xₜ = ln(Pₐ,ₜ) − α − β ln(Pᵦ,ₜ)`\n\nα、β 僅用訓練資料估計。")
    b2.markdown("**待檢驗假說**\n\n當 `xₜ` 偏高，線性模型是否估出負向調整係數？這項關係能否延伸到未見日期？")
    b3.markdown("**比較目標**\n\n`Δxₜ₊₁ = xₜ₊₁ − xₜ`\n\n以相同目標、相同測試日期評分。")

if run:
    try:
        with st.status("正在執行研究流程…", expanded=True) as status:
            st.write("讀取並檢查共同交易日與價格…")
            if source.startswith("TWSE 官方樣本"):
                data = load_bundled_twse_sample(ROOT / "data" / "twse_2330_2454_2026-05_to_2026-10.csv")
            elif source.startswith("合成"):
                data = make_synthetic_prices(seed=int(seed))
            elif source.startswith("TWSE"):
                api_progress = st.progress(0, text="準備讀取 TWSE 月資料…")
                data = fetch_twse(tickers, str(start), str(end), progress_callback=lambda done,total,label: api_progress.progress(done/total,text=f"取得 {label}（{done}/{total}）"))
            elif uploaded is not None:
                data = load_local_csv_bytes(uploaded.getvalue(), uploaded.name, tickers)
            else:
                raise ValueError("請上傳本機 CSV。")
            if source.startswith("TWSE 官方樣本") or data.synthetic:
                prices = data.prices
            else:
                prices = data.prices.loc[(data.prices.index >= pd.Timestamp(start)) & (data.prices.index < pd.Timestamp(end))]
            st.write(f"使用 {len(prices)} 筆共同價格；資料來源：{data.source}。")
            st.write("訓練期估計價差與 OU 係數；載入符合本次訓練與驗證資料的 PySR 公式，否則本機搜尋…")
            formula_cache = ROOT / "artifacts" / "precomputed" / "synthetic-seed-42" / "formula.json" if source.startswith("合成") else ROOT / "artifacts" / "real_twse_2026" / "formula.json" if source.startswith("TWSE 官方樣本") else None
            result = run_experiment(prices, tickers, ARTIFACTS, seed=int(seed), niterations=int(iterations), maxsize=int(maxsize),
                                    train_fraction=float(train_fraction), valid_fraction=float(valid_fraction),
                                    precomputed_formula_path=formula_cache,
                                    data_provenance={"source": data.source, "retrieved_at": data.retrieved_at,
                                                     "adjustment": data.adjustment, "version": data.version,
                                                     "synthetic": data.synthetic})
            result["data"] = data
            result["settings"] = settings
            st.write("公式已保存；使用同一套公式評估未參與選擇的測試期間…")
            status.update(label="比較完成", state="complete", expanded=False)
        st.session_state.result = result
        st.session_state.result_fingerprint = fingerprint
        st.rerun()
    except Exception as exc:
        st.session_state.result = None
        st.session_state.result_fingerprint = None
        st.error(f"本次實驗未完成：{exc}")
        st.info("可改用合成資料、延長期間或檢查 CSV 欄位。PySR 首次執行需完成 Julia 後端初始化。")

result = st.session_state.result
if result is None:
    st.caption("選好資料後按「執行比較」。設定一旦改變，先前結果會清除，避免新設定沿用舊結果。")
    st.subheader("三種比較方式")
    c1,c2,c3=st.columns(3)
    c1.markdown("**無變化基準**\n\n假設價差下一期不變，Δx 預測為 0。")
    c2.markdown("**線性回復／OU 啟發**\n\n用訓練資料估計 Δx = a + b·x；只有參數符合條件才談 OU 半衰期。")
    c3.markdown("**PySR 替代公式**\n\n只在訓練資料搜尋，再用驗證資料選式；不是物理定律或因果機制。")
    st.stop()

res = result
metrics = res["metrics"]
preds = res["predictions"]
data = res["data"]

st.subheader("測試結果")
if data.synthetic:
    st.warning("目前是合成資料，用來驗證程式流程與已知回復機制；不是股票實證結果。")
else:
    st.success(f"真實行情來源：{data.source}｜取得：{data.retrieved_at}｜價格：{data.adjustment}")
cols=st.columns(3)
for col, name in zip(cols, metrics.index):
    col.metric(name, f"MAE {metrics.loc[name,'MAE']:.5f}", f"RMSE {metrics.loc[name,'RMSE']:.5f}")
    improvement=metrics.loc[name,"相對基準改善"]
    col.caption("基準模型" if name=="無變化基準" else (f"相對基準 MAE 改善 {improvement:.1%}" if pd.notna(improvement) else "基準 MAE 為零，不適用"))

with st.expander("檢視資料來源、切分與參數", expanded=False):
    st.write({"來源":data.source,"取得時間":data.retrieved_at,"價格調整":data.adjustment,"資料版本／SHA256":data.version,
              "價格筆數":len(res["prices"]),"切分日期":res["metadata"]["split"],"訓練 alpha":res["alpha"],"訓練 beta":res["beta"],
              "PySR 搜尋設定":{"seed":seed,"iterations":iterations,"maxsize":maxsize},"測試參與選擇":False})
    st.dataframe(metrics, use_container_width=True)
    st.json(res["ou_diag"])

st.markdown("---")
st.subheader("測試期間：實際與預測價差變化")
fig=go.Figure()
fig.add_trace(go.Scatter(x=preds.date,y=preds.actual_dx,name="實際 Δx",line=dict(color="#172A35",width=2)))
for name,color in [("無變化基準","#9AA9A5"),("線性回復／OU 啟發","#137C78"),("PySR 替代公式","#D16B3C")]:
    fig.add_trace(go.Scatter(x=preds.date,y=preds[f"{name}_pred_dx"],name=name,line=dict(color=color,width=1.6,dash="dot")))
fig.update_layout(height=390,margin=dict(l=8,r=8,t=20,b=8),paper_bgcolor="#F6F4EE",plot_bgcolor="#FFFFFF",legend=dict(orientation="h",y=-.22),yaxis_title="Δx（對數價差變化）",xaxis_title="預測資訊日期",hovermode="x unified")
st.plotly_chart(fig,use_container_width=True)

st.subheader("測試期間：下一期價差水準")
spread_fig=go.Figure()
spread_fig.add_trace(go.Scatter(x=preds.date,y=preds.actual_next_spread,name="實際 xₜ₊₁",line=dict(color="#172A35",width=2)))
for name,color in [("無變化基準","#9AA9A5"),("線性回復／OU 啟發","#137C78"),("PySR 替代公式","#D16B3C")]:
    spread_fig.add_trace(go.Scatter(x=preds.date,y=preds[f"{name}_pred_next_spread"],name=name,line=dict(color=color,width=1.6,dash="dot")))
spread_fig.update_layout(height=350,margin=dict(l=8,r=8,t=18,b=8),paper_bgcolor="#F6F4EE",plot_bgcolor="#FFFFFF",legend=dict(orientation="h",y=-.22),yaxis_title="下一期對數價差 xₜ₊₁",xaxis_title="預測資訊日期",hovermode="x unified")
st.plotly_chart(spread_fig,use_container_width=True)

colf, colm=st.columns([1.2,1])
with colf:
    st.subheader("模型方程與白話解釋")
    st.markdown("**無變化基準**")
    st.markdown("<div class='formula'>Δx̂ₜ = 0　⇒　x̂ₜ₊₁ = xₜ</div>",unsafe_allow_html=True)
    st.caption("假設下一期變化為零；是所有模型必須超越的參照點。")
    d=res["ou_diag"]
    st.markdown("**線性回復／OU 啟發**")
    st.markdown(f"<div class='formula'>Δx̂ₜ = {d['a']:.5g} {d['b']:+.5g} × xₜ</div>",unsafe_allow_html=True)
    if d["supported"]:
        st.caption(f"訓練期 φ={d['phi']:.4f}，符合 0<φ<1；OU 啟發半衰期約 {d['half_life']:.1f} 個交易步。這不代表跨期間穩定。")
    else:
        st.caption(f"訓練期 b={d['b']:.5g}、φ={d['phi']:.4f}；未支持此回復形式，不解讀半衰期。")
    st.markdown("**PySR 替代公式**")
    st.markdown(f"<div class='formula'>Δx̂ₜ = {res['formula']['selected_equation']}</div>",unsafe_allow_html=True)
    st.caption(f"{res['formula']['origin']}；複雜度 {res['formula']['selected_complexity']}。特徵順序：x、前期 Δx、20 日歷史波動。")
with colm:
    st.subheader("測試 MAE 比較")
    bar=go.Figure(go.Bar(x=metrics.index,y=metrics.MAE,marker_color=["#9AA9A5","#137C78","#D16B3C"],text=[f"{x:.4f}" for x in metrics.MAE],textposition="outside"))
    bar.update_layout(height=300,margin=dict(l=8,r=8,t=22,b=8),paper_bgcolor="#F6F4EE",plot_bgcolor="#FFFFFF",yaxis_title="MAE（越低越好）")
    st.plotly_chart(bar,use_container_width=True)
    st.markdown("**測試期前後段 MAE**")
    st.dataframe(pd.DataFrame(res["segment_metrics"]),use_container_width=True)

best_nonbaseline=metrics.drop(index="無變化基準").MAE.idxmin()
if metrics.loc[best_nonbaseline,"MAE"] < metrics.loc["無變化基準","MAE"]:
    conclusion=f"在本組資料與本次測試期間，{best_nonbaseline} 的 MAE 低於無變化基準；這只表示本次下一期價差變化預測誤差較低。"
else:
    conclusion="本次測試未觀察到模型優於無變化基準的 MAE 表現。"
st.subheader("結論與限制")
st.markdown(conclusion)
st.markdown("結果不能解讀成交易獲利或因果證據。價差係數只用訓練期估計，PySR 只用訓練資料搜尋並以驗證資料選式，測試資料未參與挑選。單一股票組合與單次測試仍不足以證明長期穩定性。")
if data.synthetic:
    st.markdown("本節結果來自已知參數的模擬序列，僅用於驗證流程；真實樣本結果請在側欄選擇 TWSE 官方樣本。")
st.caption("研究範圍：本工具評估價差變化的樣本外預測；短期誤差比較不直接代表長期關係或投資績效。")

st.download_button("下載逐日預測 CSV", data=preds.to_csv(index=False).encode("utf-8-sig"), file_name="market_physics_predictions.csv", mime="text/csv")
with st.expander("實驗紀錄與候選公式"):
    st.json(res["metadata"])
    st.dataframe(res["front"].head(20),use_container_width=True)
    st.write("資料來源與實驗版本摘要：")
    st.code(json.dumps(res["metadata"],ensure_ascii=False,indent=2,default=str),language="json")
