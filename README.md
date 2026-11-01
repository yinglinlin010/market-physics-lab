# 股票也有回復力嗎？

**物理模型與 AI 公式的比較實驗**是一個本機 Streamlit 研究工作台，檢查兩檔台股的訓練期對數價差，並比較「無變化」、「線性回復／OU 啟發」與 PySR 搜尋公式對下一期價差變化的預測。這是物理啟發假說的研究，不是交易策略或獲利測試。

## 快速啟動

需要 macOS/Linux/Windows、Python 3.12 與 `uv`。安裝 Python 套件應在上台展示前完成。

```bash
uv sync --frozen --all-groups
uv run streamlit run app.py
```

開啟 Streamlit 顯示的本機頁面。預設選項是隨專案保存的 TWSE 官方 2330／2454 固定樣本，無需網路即可重現真實行情案例。預設訓練 60%／驗證 20%、PySR seed=42、20 次搜尋迭代與複雜度 16，會載入預先搜尋的公式；展示時執行資料準備、三模型預測和測試評估，不下載行情、不呼叫雲端模型，也不需要現場初始化 Julia。可切換至合成資料檢查程式如何回收已知回復參數。

為完全離線的 90 秒展示，先在有網路的開發環境完成套件同步；展示時保持「TWSE 官方樣本（離線，2026-05 至 10）」與預設設定。更改資料、種子、切分或搜尋設定後，既有結果會清除；若找不到完全相符的公式快取，程式會在本機執行 PySR。PySR 首次搜尋會透過 Juliapkg 下載並初始化 Julia 與 SymbolicRegression.jl，這一步不是預設展示流程。

## 資料

### 固定種子合成資料

以 seed 42 產生 900 個模擬交易日，含已知 AR(1) 回復價差。它用於檢查程式，不是股票實證。公式、PySR 候選、切分、預測與指標保存在 `artifacts/precomputed/synthetic-seed-42/`。

### 官方台股資料

`data/twse_2330_2454_2026-05_to_2026-10.csv` 是台積電（2330）與聯發科（2454）2026-05 至 2026-10 的共同交易日資料，從 TWSE 官方逐月歷史端點取得。政府資料開放平臺列出同類個股日收盤價資料，授權為 Open Government Data License 1.0。價格是未做股利或分割還原的日收盤價；來源、取得時間、期間與檔案 SHA-256 見 `data/README.md` 和實驗 manifest。再利用時請保留資料來源註記。

互動介面提供四種來源：專案附帶的固定 TWSE 官方樣本（離線）、固定種子合成資料、TWSE 官方歷史資料（需網路）及本機 `date,ticker,close` CSV。TWSE 查詢逐股逐月順序執行並在本機暫存；若官方端點限流、回應空白或改變格式，介面會說明錯誤。延長歷史區間也可以先從授權來源取得 CSV 再載入。價格不任意補值。

## 實驗方法

1. 以共同有效日為基礎，依時間切成訓練／驗證／測試，預設 60%／20%／20%，不隨機打散。
2. 僅用訓練期對 `ln(P_A)` 對 `ln(P_B)` 回歸，固定 `alpha`、`beta`，形成 `x_t = ln(P_A,t) - alpha - beta × ln(P_B,t)`。
3. 預測目標是 `y_t = x_(t+1)-x_t`。特徵為當期 `x_t`、已知的前一期價差變化及 20 日歷史變化波動率；不含未來資料，沒有用全資料擬合標準化器。
4. 排除標籤跨越訓練／驗證／測試邊界的樣本；最後一筆沒有下一期真值，不計分。
5. 無變化模型預測 `ŷ_t=0`；線性回復模型只用訓練資料估 `y_t=a+b x_t`。僅在 `b<0` 且 `0<φ=1+b<1` 時顯示 OU 啟發診斷與半衰期。
6. PySR 使用訓練資料、固定 seed、加減乘算子、serial deterministic 搜尋。候選公式只按驗證 MAE 選擇，同分選較簡單公式。搜尋輸入與驗證資料雜湊不含測試資料；公式以安全算術運算式保存並重新載入。
7. 三者對同一測試日期與目標計算 MAE、RMSE、相對基準 MAE 改善及測試前後段 MAE。分段值是描述結果，不用於選式。

沒有超越基準仍是有效結果。MAE/RMSE 不代表獲利、勝率或可交易性；一次短測試也不證明價差長期穩定。若修改模型設定時參考了測試結果，必須把該測試期間標記為開發資料，另找新的保留區間才可再作最終測試。

## 重現公式搜尋

預設離線展示不會重跑 PySR。要重建合成公式與評估產物：

```bash
uv run python scripts/search_formula.py
```

這會執行 PySR 搜尋並更新 `artifacts/precomputed/synthetic-seed-42/`。首次執行需要網路安裝 Julia；記錄在 `formula.json` 的版本、算子、seed、迭代數、複雜度上限、特徵順序、前處理、搜尋資料指紋及驗證選式規則可用來核對來源。PySR 的 Julia checkpoint 是版本相依產物，不是現場必要檔；可攜的候選表與已選公式是 CSV/JSON。

## 測試

```bash
uv run pytest
```

完整測試紀錄在 `docs/test-results.md`；實驗報告為 `docs/research_report.docx`，AI 開發過程摘要為 `docs/ai_summary.md`，90 秒簡報與講稿分別在 `slides/midterm_90s.pptx`、`slides/90s_script.md`。要依最新結果重製報告及簡報，可執行 `uv run python scripts/create_deliverables.py`；再執行 `uv run python scripts/polish_presentation.py`，套用學術簡報版面規則。

測試涵蓋合成 AR(1) 參數回收、測試期擾動不改變訓練估計及已選公式、共同預測日期、公式保存重載、安全算式、資料輸入錯誤、TWSE 日期與收盤欄位解析，以及設定變更時清除舊結果。

## 目錄

- `app.py`：Streamlit 介面
- `src/`：資料、特徵、模型、評估、實驗與結果狀態
- `data/`：授權的官方樣本與來源說明
- `artifacts/precomputed/`：離線展示公式和結果
- `artifacts/real_twse_2026/`：官方樣本的搜尋、候選公式、逐日預測與實驗紀錄
- `tests/`：資料邊界及公平比較測試
- `docs/`：AI 開發摘要、設計決策、交件報告
- `slides/`：90 秒投影片與講稿
- `PRODUCT.md`、`DESIGN.md`：產品與介面設計依據

## 研究限制

- 隨附真實資料只有 123 個共同交易日，測試只有 24 日；不適合推論長期關係。
- 價格未做股利或分割調整，`alpha` 與 `beta` 只在訓練期估計；關係可能漂移。
- PySR 搜尋屬統計式候選，可能過度擬合；短驗證／測試區間帶來很大不確定性。
- 未做顯著性檢定、跨期間驗證、交易成本分析或交易回測。
- 固定官方樣本與固定種子合成流程均可離線重現；新增即時行情需 TWSE 服務可用。

## 來源

- [PySR 官方文件](https://ai.damtp.cam.ac.uk/pysr/) 與 [PySR 官方 Skill](https://github.com/astroautomata/PySR/tree/master/skills/pysr)
- [TWSE OpenAPI](https://openapi.twse.com.tw/)
- [政府資料開放平臺資料集 11548](https://data.gov.tw/en/datasets/11548)
- [政府資料開放授權條款 1.0](https://data.gov.tw/license)
