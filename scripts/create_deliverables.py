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
    p = doc.add_paragraph()
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
doc.add_paragraph("本報告比較無變化基準、線性回復／OU 啟發模型與 PySR 搜尋公式，檢查兩檔台股對數價差變化的樣本外預測。官方樣本的 24 個測試日中，沒有一個候選模型勝過基準；因此結果不支持本次資料上的改善主張。短測試只能提供一次有限觀察，不能推論長期關係或交易獲利。")

heading("研究問題與使用對象")
doc.add_paragraph("工具面向課堂展示與初學量化研究者：使用者選資料、設定時間比例後，可以檢視方程、測試預測、誤差和限制。研究問題是兩檔股票的訓練期對數價差是否呈現回復傾向，以及 PySR 搜尋出的公式能否改善下一期價差變化預測。")

heading("資料與實驗設計")
doc.add_paragraph("合成資料以固定 seed=42 產生 900 個模擬交易日，具有已知 AR(1) 回復機制，只用於程式流程驗證。真實樣本使用台積電（2330）與聯發科（2454）2026 年 5 月至 10 月 TWSE 官方月資料，共 123 個共同交易日。資料是未做股利或分割還原的收盤價，CSV SHA-256 為 fbd8cf64c130780ac6bb6c9e8e1c4d4760e5c51363b00b57a8b7570bdd19cb01；取得時間 2026-11-01 13:34 UTC。來源與授權詳見 README 與 data/README.md。")
doc.add_paragraph("依日期順序切分：訓練 60%、驗證 20%、測試 20%。兩檔日收盤價形成 xₜ = ln(Pₐ,ₜ) − α − β ln(Pᵦ,ₜ)，α、β 僅以訓練期估計。目標 yₜ = xₜ₊₁ − xₜ；特徵為 xₜ、前一期 Δx 與 20 日歷史波動。切分邊界上沒有完整標籤的列會剔除，三模型使用同一組測試日。")

heading("模型與公式")
for name, desc in [
    ("無變化基準", "預測 Δx̂ₜ = 0，因此 x̂ₜ₊₁ = xₜ。"),
    ("線性回復／OU 啟發", "以訓練資料估計 Δx̂ₜ = a + b xₜ。只有 0 < φ = 1 + b < 1 時才計算 OU 啟發半衰期。"),
    ("PySR 替代公式", "在訓練資料以固定種子、加減乘算子及 serial 搜尋候選方程，再用驗證 MAE 選式；測試資料不參與搜尋與選擇。")]:
    heading(name, 2)
    doc.add_paragraph(desc)
doc.add_paragraph("本次 PySR 選出的公式為：Δx̂ₜ = x × (x × (−143.16728 × x) − dx_lag1)。搜尋版本 1.5.10、seed=42、20 iterations、maxsize=16。公式已存為可攜 JSON，與 PySR 模型的驗證預測最大差異為 4.34×10⁻¹⁹。")

heading("真實樣本結果")
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

heading("合成資料流程檢查")
syn = pd.read_csv(ROOT / "artifacts/precomputed/synthetic-seed-42/metrics.csv", index_col=0)
doc.add_paragraph(f"合成資料的測試 MAE 為基準 {syn.loc['無變化基準','MAE']:.6f}、OU {syn.loc['線性回復／OU 啟發','MAE']:.6f}、PySR {syn.loc['PySR 替代公式','MAE']:.6f}。這只驗證程式可跑完；不能替代股票實證，也不應據此宣稱真實市場存在可利用回復。")

heading("測試、限制與下一步")
doc.add_paragraph("9 項自動測試全數通過，涵蓋合成 AR(1) 參數回收、測試資料變更不影響訓練估計或驗證選式、日期一致、公式保存重載、安全算式、無效／過短輸入、設定失效與 TWSE 日期解析。瀏覽器已實際執行合成資料主要流程並看見指標、圖表、公式、結論與 CSV 下載入口。手機寬度、CSV 實際下載檔內容與多個瀏覽器重啟流程未完成驗收，列為未驗證。")
doc.add_paragraph("限制包括單一配對、123 個觀測日、24 個測試日、未還原價格、短驗證集、沒有顯著性或時間相依誤差區間，也沒有交易成本或回測。下一步應事先鎖定更多股票配對與多段滾動時間窗，固定模型選擇程序後再做真正保留測試；另評估企業行動調整與統計不確定性。")
doc.add_paragraph("本工具輸出的是預測誤差，不是投資建議、因果證據或獲利證明。官方資料使用與重現方式見 README.md。")
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

s=add_slide("股票也有回復力嗎？", "物理模型與 AI 公式的比較實驗")
textbox(s,.8,2.0,6.2,2.2,"兩檔股票的對數價差\n會往長期關係回復嗎？",30,INK,True)
textbox(s,.85,4.5,10.8,.8,"比較無變化基準、線性回復模型與 PySR 搜尋公式",20,TEAL)
textbox(s,.85,6.45,11,.45,"研究問題：下一期價差變化能否比「預測不變」更準？",14,MUTED)

s=add_slide("固定時間切分，公平比較三種模型")
textbox(s,.8,1.8,5.4,3.9,"資料：2330 台積電 × 2454 聯發科\n123 個共同交易日\n\n訓練 60% → 驗證 20% → 測試 20%\n\n測試：24 個預測日",20,INK)
textbox(s,6.7,1.8,5.8,3.8,"① 無變化：Δx̂ = 0\n\n② OU 啟發：Δx̂ = a + b·x\n\n③ PySR：訓練搜尋、驗證選式\n\n公式預先搜尋，評估現場執行",19,TEAL)
textbox(s,.85,6.55,11,.4,"α、β 只用訓練資料估計；測試資料不參與選股票、調參或選公式。",14,MUTED)

s=add_slide("真實樣本裡，複雜模型沒有勝過基準")
s.shapes.add_picture(str(chart),PptInches(.65),PptInches(1.65),width=PptInches(7.5))
textbox(s,8.5,1.95,4.1,3.8,"無變化\nMAE 0.008493\n\nOU：0.008930\nPySR：0.010204",21,INK,True)
textbox(s,8.5,5.45,4.1,.85,"兩個候選模型的測試誤差都較高。",16,RUST,True)

s=add_slide("結論：目前證據不足以支持穩定回復")
textbox(s,.85,1.75,11.2,1.2,"本次測試沒有模型超越無變化基準。",27,TEAL,True)
textbox(s,.9,3.1,11.2,2.3,"只有一組股票配對、24 個測試日；價格未做股利或分割調整。\n短期預測誤差不等於交易獲利，也不是因果證據。\n下一步：事先固定多配對、多期間的滾動測試。",19,INK)
textbox(s,.9,6.3,11,.55,"合成資料只用來驗證程式流程，不作為真實股票結論。",14,MUTED)

prs.save(SLIDES / "midterm_90s.pptx")

script = """# 90 秒展示講稿\n\n**0–15 秒｜研究問題** 這個工具要問：兩檔股票的對數價差會不會回復？我比較預測不變、線性回復，以及 PySR 搜尋的公式。\n\n**15–35 秒｜方法** 左側可選合成、TWSE 或本機 CSV，設定股票和日期切分。示範先用預先搜尋的合成公式，現場只跑評估。真實資料依時間切成訓練、驗證、測試；價差係數只用訓練資料估計，PySR 用訓練搜尋、驗證選公式，測試集不參與挑選。\n\n**35–60 秒｜結果** 這組 TWSE 台積電和聯發科資料有 123 個共同交易日，測試 24 日。基準 MAE 是 0.008493；OU 是 0.008930；PySR 是 0.010204。兩個模型都沒有勝過基準。\n\n**60–78 秒｜解讀** 所以這次結果沒有支持候選模型在測試期改善預測。測試很短、只有一組配對，也沒有調整股利與分割，不能推論長期穩定性。預測誤差也不是獲利。\n\n**78–90 秒｜下一步** 下一步先固定更多配對與滾動測試日期，再用完全保留的新期間評估。這個工作台會輸出逐日預測和實驗紀錄，讓方程與結論都能追查。\n"""
(SLIDES / "90s_script.md").write_text(script, encoding="utf-8")

