from __future__ import annotations

from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PptInches, Pt as PptPt

ROOT = Path(__file__).resolve().parents[1]
DOCS = ROOT / "docs"
SLIDES = ROOT / "slides"
ASSETS = DOCS / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)
SLIDES.mkdir(parents=True, exist_ok=True)

metrics = pd.read_csv(ROOT / "artifacts/real_twse_2026/metrics.csv", index_col=0)
preds = pd.read_csv(ROOT / "artifacts/real_twse_2026/predictions.csv", parse_dates=["date"])
manifest = __import__("json").loads((ROOT / "artifacts/real_twse_2026/run_manifest.json").read_text())

plt.rcParams["font.family"] = ["PingFang TC", "Arial Unicode MS", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False
fig, ax = plt.subplots(figsize=(8, 3.7), dpi=180)
colors = ["#9AA9A5", "#137C78", "#D16B3C"]
bars = ax.bar(["無變化基準", "線性回復／OU", "PySR 公式"], metrics["MAE"], color=colors, width=.58)
ax.set_ylabel("測試 MAE（越低越好）")
ax.set_title("TWSE 官方樣本｜測試期間 24 個預測日")
ax.spines[["top", "right"]].set_visible(False)
ax.grid(axis="y", color="#DCE2DD", linewidth=.7)
ax.set_axisbelow(True)
ax.bar_label(bars, fmt="%.5f", padding=4, fontsize=9)
ax.set_ylim(0, max(metrics["MAE"]) * 1.3)
fig.tight_layout()
chart = ASSETS / "real_test_mae.png"
fig.savefig(chart, facecolor="#F6F4EE", bbox_inches="tight")
plt.close(fig)

# Editable research report.
doc = Document()
section = doc.sections[0]
section.top_margin = Inches(.7)
section.bottom_margin = Inches(.65)
section.left_margin = Inches(.8)
section.right_margin = Inches(.8)
styles = doc.styles
styles["Normal"].font.name = "PingFang TC"
styles["Normal"]._element.rPr.rFonts.set(__import__("docx").oxml.ns.qn("w:eastAsia"), "PingFang TC")
styles["Normal"].font.size = Pt(10.5)
styles["Normal"].paragraph_format.space_after = Pt(6)
qn = __import__("docx").oxml.ns.qn

def heading(text: str, level: int = 1):
    p = doc.add_paragraph(style="Heading 1" if level == 1 else "Heading 2")
    p.paragraph_format.space_before = Pt(12 if level == 1 else 7)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.name = "PingFang TC"
    r._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang TC")
    r.font.size = Pt(16 if level == 1 else 12)
    r.font.color.rgb = RGBColor.from_string("137C78" if level == 1 else "172A35")
    return p

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("股票也有回復力嗎？")
r.bold = True
r.font.name = "PingFang TC"
r._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang TC")
r.font.size = Pt(25)
r.font.color.rgb = RGBColor.from_string("172A35")
p.paragraph_format.space_after = Pt(8)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("物理模型與 AI 公式的比較實驗｜期中專題研究報告")
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(95, 114, 122)
heading("摘要")
doc.add_paragraph("本專題以兩檔台股的對數價差為研究對象，比較無變化基準、線性回復／Ornstein–Uhlenbeck（OU）啟發模型與 PySR 符號回歸公式。研究資料取自臺灣證券交易所 2026 年 5 月至 10 月的逐日收盤價，共 123 個共同交易日；依日期順序進行訓練、驗證、測試。模型預測目標為下一期對數價差變化。真實樣本的 24 個測試日中，基準 MAE 為 0.008493，OU 模型為 0.008930，PySR 為 0.010204；本次候選模型均未降低測試誤差。結果提供一個可追溯的比較案例，並展示物理啟發模型與符號回歸在相同評估規則下的差異。")
p=doc.add_paragraph(); p.add_run("關鍵詞：").bold=True; p.add_run("對數價差、均值回復、OU 模型、符號回歸、PySR、時間序列驗證")

doc.add_page_break()
heading("一、研究背景與目的")
doc.add_paragraph("相關公司可能共同受到市場景氣、產業消息與資金流動影響，股價水準因此呈現相似趨勢。比較單一股票的絕對價格，較難分辨共同波動與相對偏離；將兩檔價格取對數，再以訓練期估計兩者的線性關係，可建立一條相對價差序列。本研究關心的是：當價差偏離其訓練期關係後，下一期是否出現系統性的反向調整？此問題與共整合及誤差修正的時間序列觀點相關，但本專題只作模型比較，沒有執行正式共整合檢定，因此不將價格關係視為已獲證明的長期均衡。[1]")
doc.add_paragraph("研究目的有三項：第一，以固定時間切分建立可重現的樣本外比較；第二，將線性回復模型的方向與速度，和 PySR 搜尋出的可讀方程並列；第三，檢查較複雜的公式是否能在未參與搜尋與選式的測試期間，超越預測變化為零的基準。工具提供課堂展示與研究教學使用，讓使用者同時閱讀方程、資料切分與誤差指標。")

heading("二、理論背景與研究假說")
doc.add_paragraph("OU 過程以隨機擾動搭配朝長期平均值回復的漂移描述連續時間動態。其基本形式可寫為 dxₜ = κ(μ−xₜ)dt + σdWₜ，其中 κ 控制回復速度，μ 為平均位置，σdWₜ 表示隨機擾動。[2] 在本研究中，模型使用離散交易日的線性近似：Δxₜ₊₁ = a + b xₜ + εₜ₊₁。當 b<0 且 φ=1+b 落在 0 與 1 之間，模型的條件預測方向指向均衡位置；相應的離散半衰期為 ln(0.5)/ln(φ)。這是 OU 啟發診斷，不是以連續時間似然估計得到的 OU 參數，也不保證參數在其他區間維持不變。")
doc.add_paragraph("PySR 將符號回歸轉為候選方程搜尋，透過受限算子與複雜度控制，在資料中尋找可讀的數學表達式。[3] 本專題採用加、減、乘三種算子並固定搜尋種子；候選式只依驗證集 MAE 選擇。研究假說設定為可檢驗標準，而不是預期必然成立的結論：H1，訓練資料中的線性係數符合回復條件；H2，PySR 或 OU 的測試 MAE 低於無變化基準。測試結果可支持或不支持 H2。")

heading("三、資料與實驗設計")
doc.add_paragraph("合成資料以固定 seed=42 產生 900 個模擬交易日，具有已知 AR(1) 回復機制，只用於程式流程驗證。真實樣本使用台積電（2330）與聯發科（2454）2026 年 5 月至 10 月 TWSE 官方月資料，共 123 個共同交易日。資料是未做股利或分割還原的收盤價，CSV SHA-256 為 fbd8cf64c130780ac6bb6c9e8e1c4d4760e5c51363b00b57a8b7570bdd19cb01；取得時間 2026-11-01 13:34 UTC。來源與授權詳見 README 與 data/README.md。")
doc.add_paragraph("依日期順序切分：訓練 60%、驗證 20%、測試 20%。兩檔日收盤價形成 xₜ = ln(Pₐ,ₜ) − α − β ln(Pᵦ,ₜ)，α、β 僅以訓練期估計。目標 yₜ = xₜ₊₁ − xₜ；特徵為 xₜ、前一期 Δx 與 20 日歷史波動。切分邊界上沒有完整標籤的列會剔除，三模型使用同一組測試日。")

heading("四、模型與預測流程")
for name, desc in [
    ("無變化基準", "預測 Δx̂ₜ = 0，因此 x̂ₜ₊₁ = xₜ。"),
    ("線性回復／OU 啟發", "以訓練資料估計 Δx̂ₜ = a + b xₜ。只有 0 < φ = 1 + b < 1 時才計算 OU 啟發半衰期。"),
    ("PySR 替代公式", "在訓練資料以固定種子、加減乘算子及 serial 搜尋候選方程，再用驗證 MAE 選式；測試資料不參與搜尋與選擇。")]:
    heading(name, 2)
    doc.add_paragraph(desc)
doc.add_paragraph("本次 PySR 選出的公式為：Δx̂ₜ = x × (x × (−143.16728 × x) − dx_lag1)。搜尋版本 1.5.10、seed=42、20 iterations、maxsize=16。公式已存為可攜 JSON，與 PySR 模型的驗證預測最大差異為 4.34×10⁻¹⁹。")

heading("五、真實樣本結果與討論")
doc.add_picture(str(chart), width=Inches(6.3))
cap = doc.add_paragraph("圖 1｜TWSE 樣本的測試集 MAE。越低越好；測試日期 2026-09-21 至 2026-10-30，共 24 日。")
cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
cap.runs[0].italic = True
cap.runs[0].font.size = Pt(9)
table = doc.add_table(rows=1, cols=4)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = "Light Shading Accent 1"
for cell, text in zip(table.rows[0].cells, ["模型", "MAE", "RMSE", "相對基準改善"]): cell.text = text
for name, row in metrics.iterrows():
    cells = table.add_row().cells
    vals = [name, f"{row.MAE:.6f}", f"{row.RMSE:.6f}", "基準" if name == "無變化基準" else f"{row['相對基準改善']:.2%}"]
    for cell, text in zip(cells, vals): cell.text = text
doc.add_paragraph("測試 MAE：基準 0.008493、OU 0.008930（相對基準 −5.15%）、PySR 0.010204（−20.16%）。RMSE 亦皆高於基準。這是未勝過基準的有效結果，沒有改用測試集挑選股票或公式。")

heading("六、合成資料的流程驗證")
syn = pd.read_csv(ROOT / "artifacts/precomputed/synthetic-seed-42/metrics.csv", index_col=0)
doc.add_paragraph(f"合成資料的測試 MAE 為基準 {syn.loc['無變化基準','MAE']:.6f}、OU {syn.loc['線性回復／OU 啟發','MAE']:.6f}、PySR {syn.loc['PySR 替代公式','MAE']:.6f}。這只驗證程式可跑完；不能替代股票實證，也不應據此宣稱真實市場存在可利用回復。")

heading("七、限制、驗證狀態與後續工作")
doc.add_paragraph("本專案自動測試 10 項全數通過，涵蓋合成 AR(1) 參數回收、測試資料變更不影響訓練估計或驗證選式、三模型日期一致、公式保存重載、安全算式、無效／過短輸入、設定變更失效、TWSE 欄位解析與官方樣本來源指紋。瀏覽器已執行官方樣本的離線評估流程，檢視指標、圖表、方程、結論及下載入口。手機窄視窗完整互動與透過瀏覽器下載後逐欄核對 CSV 尚未完成，應視為未驗證。")
doc.add_paragraph("本次推論範圍受單一股票配對、123 個共同交易日與 24 個測試日限制。價格採原始收盤價，未作股利或分割還原；驗證期也短，尚未提供顯著性檢定或時間相依的不確定性區間。因此結果適用於指定資料版本與期間。預測誤差比較不代表因果證據或實際績效。")
doc.add_paragraph("後續研究可預先固定多組股票配對與多個滾動測試期間，檢查回復係數及候選公式的跨期穩定性，並比較企業行動調整資料與不同誤差估計方式。官方資料來源、取得時間、價格調整方式與完整重現步驟見 README.md 和 data/README.md。")

heading("參考資料")
for ref in [
    "[1] Engle, R. F., & Granger, C. W. J. (1987). Co-integration and Error Correction: Representation, Estimation, and Testing. Econometrica, 55(2), 251–276. https://doi.org/10.2307/1913236",
    "[2] Uhlenbeck, G. E., & Ornstein, L. S. (1930). On the Theory of the Brownian Motion. Physical Review, 36(5), 823–841. https://doi.org/10.1103/PhysRev.36.823",
    "[3] Cranmer, M. (2023). Interpretable Machine Learning for Science with PySR and SymbolicRegression.jl. arXiv:2305.01582. https://arxiv.org/abs/2305.01582",
    "[4] 臺灣證券交易所，個股日收盤價資料（STOCK_DAY）；政府資料開放平臺資料集 11548，政府資料開放授權條款 1.0。https://data.gov.tw/en/datasets/11548",
]:
    p=doc.add_paragraph(ref); p.paragraph_format.left_indent=Inches(.25); p.paragraph_format.first_line_indent=Inches(-.25)
doc.save(DOCS / "research_report.docx")

# Presentation with concise speaker notes in a companion script.
prs = Presentation()
prs.slide_width = PptInches(13.333)
prs.slide_height = PptInches(7.5)
BG = PptColor(246, 244, 238)
INK = PptColor(23, 42, 53)
TEAL = PptColor(19, 124, 120)
MUTED = PptColor(95, 114, 122)
RUST = PptColor(209, 107, 60)

def add_slide(title: str, kicker: str = "期中專題｜市場物理工作台"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill; fill.solid(); fill.fore_color.rgb = BG
    box = slide.shapes.add_textbox(PptInches(.65), PptInches(.35), PptInches(12), PptInches(.35))
    run = box.text_frame.paragraphs[0].add_run(); run.text = kicker; run.font.name="PingFang TC"; run.font.size=PptPt(13); run.font.color.rgb=TEAL
    box = slide.shapes.add_textbox(PptInches(.65), PptInches(.83), PptInches(12), PptInches(.8))
    p = box.text_frame.paragraphs[0]; p.text=title; p.font.name="PingFang TC"; p.font.size=PptPt(28); p.font.bold=True; p.font.color.rgb=INK
    return slide

def textbox(slide, x,y,w,h,text,size=18,color=INK,bold=False):
    shape=slide.shapes.add_textbox(PptInches(x),PptInches(y),PptInches(w),PptInches(h))
    tf=shape.text_frame; tf.word_wrap=True; tf.margin_left=PptInches(.05); tf.margin_right=PptInches(.05); tf.margin_top=PptInches(.03)
    for i,line in enumerate(text.split("\n")):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.text=line; p.font.name="PingFang TC"; p.font.size=PptPt(size); p.font.bold=bold; p.font.color.rgb=color; p.space_after=PptPt(9)
    return shape

def card(slide, x, y, w, h, title, body, accent=TEAL):
    shape=slide.shapes.add_shape(1,PptInches(x),PptInches(y),PptInches(w),PptInches(h))
    shape.fill.solid(); shape.fill.fore_color.rgb=PptColor(255,255,255)
    shape.line.color.rgb=PptColor(220,226,221)
    textbox(slide,x+.18,y+.13,w-.36,.42,title,17,accent,True)
    textbox(slide,x+.18,y+.66,w-.36,h-.78,body,15,INK)

s=add_slide("股票也有回復力嗎？", "物理模型與 AI 公式的比較實驗｜期中專題")
textbox(s,.85,1.9,11.7,1.35,"兩檔股票的相對價差，\n會在偏離後往回移動嗎？",30,INK,True)
textbox(s,.9,4.05,10.8,.7,"以台積電與聯發科的官方日收盤價，檢驗價差變化的樣本外預測。",19,TEAL)
card(s,.9,5.25,3.55,1.15,"比較基準","變化為零",MUTED)
card(s,4.8,5.25,3.55,1.15,"物理啟發","線性回復／OU",TEAL)
card(s,8.7,5.25,3.55,1.15,"資料搜尋","PySR 方程",RUST)

s=add_slide("研究背景：從共同波動辨認相對偏離")
textbox(s,.9,1.7,5.35,1.2,"相關公司可能共同受到市場與產業因素影響，股價水準因此同步波動。",19,INK)
card(s,6.65,1.65,5.5,1.35,"共同趨勢","市場與產業消息同時影響兩檔股票",TEAL)
textbox(s,8.95,3.08,.8,.45,"↓",26,TEAL,True)
card(s,6.65,3.55,5.5,1.35,"相對價差","以訓練期價格關係衡量相對偏離",RUST)
textbox(s,.95,3.3,5.15,1.0,"xₜ = ln(Pₐ,ₜ) − α − β ln(Pᵦ,ₜ)",22,TEAL,True)
textbox(s,.95,4.45,5.4,1.2,"研究問題：當 xₜ 偏離其訓練期位置時，下一期 Δx 是否呈現反向調整？",18,INK)
textbox(s,.95,6.35,11.3,.4,"研究動機參照共整合與誤差修正觀點；本實驗比較預測，不作正式共整合檢定。",13,MUTED)

s=add_slide("模型比較：由簡單基準到資料搜尋公式")
card(s,.8,1.75,3.8,3.55,"無變化基準","Δx̂ₜ = 0\n\n假設下一期價差維持原值，提供清楚的預測參照。",MUTED)
card(s,4.78,1.75,3.8,3.55,"線性回復／OU 啟發","Δx̂ₜ = a + b·xₜ\n\n只用訓練資料估計。b<0 且 0<φ=1+b<1 時，線性條件預測指向均衡位置。",TEAL)
card(s,8.75,1.75,3.8,3.55,"PySR 符號回歸","在訓練資料搜尋公式；以驗證 MAE 選式。\n\n算子：+、−、×\n特徵：x、前期 Δx、20 日波動",RUST)
textbox(s,.9,5.85,11.4,.65,"三模型預測同一目標：Δxₜ₊₁ = xₜ₊₁ − xₜ。測試資料只在公式定版後使用。",16,INK,True)

s=add_slide("資料與驗證：保留時間順序，避免測試洩漏")
card(s,.85,1.7,3.6,1.55,"樣本","台積電 2330 × 聯發科 2454\nTWSE 官方日收盤價，123 個共同交易日",TEAL)
card(s,4.85,1.7,3.6,1.55,"時間切分","訓練 60% → 驗證 20% → 測試 20%\n測試期共 24 個預測日",TEAL)
card(s,8.85,1.7,3.6,1.55,"固定規則","價差係數只用訓練估計\nPySR 訓練搜尋、驗證選式",TEAL)
textbox(s,1.05,4.05,11.0,1.2,"共同價格樣本：2026-05-04—2026-10-30\n測試日期：2026-09-21—2026-10-30\n來源版本、取得時間與原始收盤價定義均隨專案保存。",19,INK)
textbox(s,1.05,6.3,11.0,.4,"價差採原始收盤價；未做股利或分割還原。",13,MUTED)

s=add_slide("測試結果：兩個候選模型都未降低基準誤差")
s.shapes.add_picture(str(chart),PptInches(.65),PptInches(1.7),width=PptInches(7.5))
card(s,8.55,1.9,3.95,1.0,"無變化基準","MAE 0.008493",MUTED)
card(s,8.55,3.15,3.95,1.0,"線性回復／OU","MAE 0.008930｜相對改善 −5.15%",TEAL)
card(s,8.55,4.4,3.95,1.0,"PySR 公式","MAE 0.010204｜相對改善 −20.16%",RUST)
textbox(s,8.65,5.8,3.75,.55,"RMSE 同樣高於基準。",16,INK,True)

s=add_slide("結論與後續：回復假說仍待更廣泛檢驗")
textbox(s,.95,1.65,11.3,.9,"本次樣本中，候選模型未能改善測試預測。",26,TEAL,True)
card(s,.95,3.0,5.2,2.2,"本次可支持的結論","指定配對與測試期間內，無變化基準的 MAE 最低。",TEAL)
card(s,6.75,3.0,5.2,2.2,"下一步研究","預先固定多組配對、多段滾動測試與價格調整方式，再評估跨期穩定性。",RUST)
textbox(s,.95,6.25,11.3,.55,"短期預測誤差是模型比較結果；本研究不據此推論因果關係或實際績效。",14,MUTED)

prs.save(SLIDES / "midterm_90s.pptx")

script = """# 90 秒展示講稿

**0–8 秒｜題目** 這項專題比較物理啟發模型與 AI 搜尋公式，檢驗股票相對價差是否呈現可預測的回復變化。

**8–26 秒｜研究背景** 兩家同產業公司的股價可能共同受到市場消息影響。我們先用訓練期估計兩者的線性價格關係，再觀察對數價差偏離後，下一期是否往回移動。這是待檢驗的假說，不預設它一定成立。

**26–44 秒｜模型** 第一個模型假設價差變化為零；第二個以線性係數描述回復方向與速度；第三個用 PySR 從訓練資料搜尋候選方程，再用驗證資料選式。三者預測相同的下一期價差變化。

**44–59 秒｜資料與公平比較** 我們使用台積電與聯發科的 TWSE 官方收盤價，共 123 個共同交易日，依時間切成訓練、驗證、測試；最後 24 日只留作測試。價差係數只用訓練期估計，測試集不參與挑選。

**59–78 秒｜結果** 基準 MAE 是 0.008493，OU 是 0.008930，PySR 是 0.010204。這組測試中，兩個候選模型都沒有降低基準誤差；RMSE 的比較方向也相同。

**78–90 秒｜結論** 因此目前結果不支持這組資料上的預測改善。下一步應預先設定更多股票配對和滾動測試區間，檢驗回復參數與公式是否跨期間穩定。
"""
(SLIDES / "90s_script.md").write_text(script, encoding="utf-8")

