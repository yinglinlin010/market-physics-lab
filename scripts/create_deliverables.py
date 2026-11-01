from __future__ import annotations

import json
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PptColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PptInches, Pt as PptPt

ROOT = Path(__file__).resolve().parents[1]
DOCS, SLIDES = ROOT / "docs", ROOT / "slides"
ASSETS = DOCS / "assets"
ASSETS.mkdir(parents=True, exist_ok=True)
SLIDES.mkdir(parents=True, exist_ok=True)
metrics = pd.read_csv(ROOT / "artifacts/real_twse_2026/metrics.csv", index_col=0)
preds = pd.read_csv(ROOT / "artifacts/real_twse_2026/predictions.csv", parse_dates=["date"])
manifest = json.loads((ROOT / "artifacts/real_twse_2026/run_manifest.json").read_text())
formula = manifest["pysr_formula"]

# Publication-style figures: restrained colors, labels, units, and source context.
plt.rcParams["font.family"] = ["PingFang TC", "Arial Unicode MS", "sans-serif"]
plt.rcParams["axes.unicode_minus"] = False
COLORS = ["#75858A", "#147C78", "#C66B43"]
labels = ["無變化基準", "線性回復／OU 啟發", "PySR 符號回歸"]
fig, axes = plt.subplots(1, 2, figsize=(10, 4), dpi=180)
for ax, column, title in zip(axes, ["MAE", "RMSE"], ["平均絕對誤差（MAE）", "均方根誤差（RMSE）"]):
    bars = ax.bar(labels, metrics[column].to_numpy(), color=COLORS, width=.62)
    ax.bar_label(bars, fmt="%.5f", padding=3, fontsize=12)
    ax.set_title(title, fontsize=12, pad=10)
    ax.set_ylabel("測試集誤差（對數價差變化）", fontsize=12)
    ax.tick_params(axis="x", labelrotation=10, labelsize=12)
    ax.tick_params(axis="y", labelsize=12)
    ax.spines[["top", "right"]].set_visible(False)
    ax.grid(axis="y", color="#DCE2DD", linewidth=.6)
    ax.set_axisbelow(True)
    ax.set_ylim(0, max(metrics[column]) * 1.28)
fig.suptitle("TWSE 官方樣本｜24 個一步預測測試目標", fontsize=12, y=1.02)
fig.tight_layout()
error_figure = ASSETS / "academic_test_errors.png"
fig.savefig(error_figure, facecolor="white", bbox_inches="tight")
plt.close(fig)

fig, ax = plt.subplots(figsize=(9, 4.1), dpi=180)
ax.plot(preds.date, preds.actual_dx, color="#172A35", lw=1.8, marker="o", ms=2.5, label="實際 Δx")
for name, color in zip(metrics.index, COLORS):
    ax.plot(preds.date, preds[f"{name}_pred_dx"], color=color, lw=1.25, ls="--", label=name)
fig.suptitle("測試期間逐日價差變化：實際值與一步預測", fontsize=12, y=.98)
ax.set_ylabel("Δx（對數價差變化）", fontsize=12)
ax.set_xlabel("預測資訊日 t（目標為下一交易日）", fontsize=12)
ax.tick_params(labelsize=12)
ax.spines[["top", "right"]].set_visible(False)
ax.grid(color="#E6E9E6", linewidth=.55)
ax.legend(frameon=False, ncol=2, fontsize=12, loc="upper center", bbox_to_anchor=(.5, 1.28))
fig.subplots_adjust(left=.09, right=.98, bottom=.18, top=.68)
prediction_figure = ASSETS / "academic_test_predictions.png"
fig.savefig(prediction_figure, facecolor="white", bbox_inches="tight")
plt.close(fig)

# Word helpers.
NAVY = "19313D"
TEAL = "147C78"
MUTED = "63747A"
PALE = "EAF2F0"

def set_run_font(run, size=10.5, bold=None, color=NAVY, italic=None):
    run.font.name = "PingFang TC"
    run._element.get_or_add_rPr().rFonts.set(qn("w:eastAsia"), "PingFang TC")
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def add_page_number(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("— ")
    set_run_font(run, 9, color=MUTED)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    paragraph._p.append(field)
    run = paragraph.add_run(" —")
    set_run_font(run, 9, color=MUTED)


def shade_cell(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)


def body(text, first_indent=True):
    p = doc.add_paragraph(style="Normal")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.first_line_indent = Inches(.28) if first_indent else Inches(0)
    p.paragraph_format.keep_together = True
    p.add_run(text)
    return p


def section_heading(text, level=1):
    p = doc.add_paragraph(style=f"Heading {level}")
    p.paragraph_format.keep_with_next = True
    p.add_run(text)
    return p


def equation(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.keep_together = True
    r = p.add_run(text)
    set_run_font(r, 11, color=TEAL)
    return p


def caption(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_run_font(r, 9, color=MUTED)
    return p


doc = Document()
section = doc.sections[0]
section.page_width, section.page_height = Inches(8.27), Inches(11.69)
section.top_margin, section.bottom_margin = Inches(.78), Inches(.75)
section.left_margin, section.right_margin = Inches(.92), Inches(.82)
section.header_distance, section.footer_distance = Inches(.35), Inches(.35)
styles = doc.styles
normal = styles["Normal"]
normal.font.name = "PingFang TC"
normal._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang TC")
normal.font.size = Pt(10.5)
normal.font.color.rgb = RGBColor.from_string(NAVY)
normal.paragraph_format.space_after = Pt(8)
normal.paragraph_format.line_spacing = 1.3
for name, size, color in [("Heading 1", 15, TEAL), ("Heading 2", 12, NAVY), ("Heading 3", 10.5, TEAL)]:
    style = styles[name]
    style.font.name = "PingFang TC"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "PingFang TC")
    style.font.size = Pt(size)
    style.font.bold = True
    style.font.color.rgb = RGBColor.from_string(color)
    style.paragraph_format.space_before = Pt(11 if name == "Heading 1" else 7)
    style.paragraph_format.space_after = Pt(4)
    style.paragraph_format.keep_with_next = True
    # Word's inherited heading definitions can carry list numbering. Keep academic
    # section headings as clean, unnumbered paragraphs controlled by our text.
    p_pr = style._element.get_or_add_pPr()
    num_pr = p_pr.find(qn("w:numPr"))
    if num_pr is not None:
        p_pr.remove(num_pr)
normal_p_pr = normal._element.get_or_add_pPr()
normal_num_pr = normal_p_pr.find(qn("w:numPr"))
if normal_num_pr is not None:
    normal_p_pr.remove(normal_num_pr)
header = section.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
header.add_run("期中專題研究報告　｜　股票價差回復模型比較")
for r in header.runs:
    set_run_font(r, 8.5, color=MUTED)
footer = section.footer.paragraphs[0]
add_page_number(footer)

# Cover page.
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(85)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("期中專題研究報告")
set_run_font(r, 13, bold=True, color=TEAL)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(38)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("股票也有回復力嗎？")
set_run_font(r, 28, bold=True, color=NAVY)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("物理模型與 AI 公式的比較實驗")
set_run_font(r, 16, color=MUTED)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(36)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
for line in ["研究對象　台積電（2330）與聯發科（2454）", "資料期間　2026 年 5 月 4 日至 10 月 30 日", "評估設計　訓練／驗證／測試時間切分", "資料版本　TWSE 官方固定樣本；SHA-256 詳列於第三節"]:
    rr = p.add_run(line + "\n")
    set_run_font(rr, 11, color=NAVY)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(38)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("物理三　｜　林映霖")
set_run_font(r, 14, bold=True, color=NAVY)
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(16)
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("研究報告｜2026 年 11 月")
set_run_font(r, 10, color=MUTED)
doc.add_page_break()

section_heading("摘要", 1)
body("本研究探討兩檔股票之訓練期對數價差是否呈現可預測的回復變化，並比較無變化基準、線性回復／Ornstein–Uhlenbeck（OU）啟發模型與 PySR 符號回歸公式。研究使用臺灣證券交易所（TWSE）台積電（2330）與聯發科（2454）2026 年 5 月至 10 月原始日收盤價，共 123 個共同交易日。對數價差係數僅以訓練資料估計；樣本依日期切分為訓練、驗證與測試，PySR 僅在訓練集搜尋候選公式，並以驗證集 MAE 選式。模型以逐日一步方式預測下一期對數價差變化。24 個測試目標的基準 MAE 為 0.008493、OU 為 0.008930、PySR 為 0.010204；兩候選模型相對基準的 MAE 改善分別為 −5.15% 與 −20.16%，RMSE 亦較高。訓練期 OU 啟發係數符合指定的線性回復條件，但此訓練內診斷未轉化為本測試期間的預測改善。本研究提供一個可重現且可追溯的比較案例；結果僅適用於此配對、資料版本與測試期間。")
p = doc.add_paragraph()
p.paragraph_format.space_before = Pt(5)
r = p.add_run("關鍵詞：")
set_run_font(r, 10.5, bold=True, color=TEAL)
r = p.add_run("對數價差、均值回復、OU 啟發模型、符號回歸、PySR、時間序列驗證")
set_run_font(r, 10.5)

section_heading("研究問題", 2)
body("RQ1：訓練期估計的線性價差變化模型，是否符合線性回復的方向條件？", first_indent=False)
body("RQ2：OU 啟發模型或 PySR 搜尋公式，能否在未參與估計與選式的測試資料上，降低無變化基準的預測誤差？", first_indent=False)
doc.add_page_break()
section_heading("一、研究背景與目的", 1)
body("相關公司的價格可能共同受到市場、產業與資金環境影響。當兩個價格序列共享變動來源時，個別價格水準本身未必能反映兩者的相對偏離；以訓練期價格關係構造的對數價差，提供另一個可檢驗的狀態變數。若價差偏離後，後續變化在條件平均上朝相反方向移動，則可描述為回復傾向。這是一項經驗假說，不能由短期價格相關或單次模型係數直接確立。")
body("本研究的目的有三：第一，在固定時間順序下建立訓練、驗證、測試程序；第二，比較基準、物理啟發線性模型與資料搜尋方程；第三，報告模型在共同測試目標上的誤差，並檢視訓練期所見回復訊號是否延伸至樣本外。研究脈絡參照共整合與誤差修正文獻，但本專題未進行單根或共整合檢定，因此不主張兩股存在正式的長期均衡關係（Engle & Granger, 1987）。")

section_heading("二、理論背景與研究假說", 1)
section_heading("2.1 對數價差與回復假說", 2)
body("對數價格的線性組合常用於描述兩個序列間的相對關係。若一組線性組合具有穩定性，可與共整合觀點連結；然而，估計一條訓練期線性關係本身並不足以證明共整合。本研究將該組合視為預測特徵，並把樣本外預測表現與正式的時間序列性質檢定分開處理。")
section_heading("2.2 OU 啟發之離散線性模型", 2)
body("Ornstein–Uhlenbeck 過程以朝長期平均位置的漂移及隨機擾動描述連續時間動態（Uhlenbeck & Ornstein, 1930）：")
equation("dXₜ = κ(μ − Xₜ)dt + σdWₜ")
body("在等距離散觀測下，可用線性條件期望作為簡化近似。程式直接估計價差變化對當期價差的線性關係：")
equation("Δxₜ₊₁ = a + b xₜ + εₜ₊₁；　xₜ₊₁ = a + φxₜ + εₜ₊₁，φ = 1 + b")
body("若 0 < φ < 1，則離散線性條件預測朝固定點 μ = a/(1−φ) 收斂，其半衰期為 ln(0.5)/ln(φ) 個交易步。本研究稱此規則為 OU 啟發診斷，而非連續時間 OU 參數的似然估計；訓練內符合條件也不代表係數跨期穩定。")
section_heading("2.3 符號回歸與 PySR", 2)
body("符號回歸在指定運算子集合與複雜度限制下搜尋可解釋的方程。PySR 將候選方程式整理為帕累托前緣；本研究限定加、減、乘運算子，固定亂數種子與序列化搜尋設定，候選式只依驗證集 MAE 選擇（Cranmer, 2023）。測試集不參與搜尋、選式或配對挑選。")
section_heading("2.4 研究假說", 2)
body("H1（訓練期參數診斷）：線性模型係數滿足 b<0 且 0<φ=1+b<1。H2（樣本外預測）：OU 啟發模型或 PySR 公式至少一者的測試 MAE 低於無變化基準。H1 是訓練資料內的參數條件；H2 是保留測試上的比較命題，兩者不應混為同一種證據。")

section_heading("三、資料與研究方法", 1)
section_heading("3.1 資料來源與樣本", 2)
body("價格資料來自 TWSE 個股日收盤價（STOCK_DAY）官方月資料，研究資料檔隨專案保存。配對為台積電（2330）與聯發科（2454），共同有效日自 2026-05-04 至 2026-10-30，共 123 個交易日。使用原始日收盤價，未調整股利及股票分割。資料取得時間為 2026-11-01 13:34:02 UTC；CSV SHA-256：fbd8cf64c130780ac6bb6c9e8e1c4d4760e5c51363b00b57a8b7570bdd19cb01。政府資料開放平臺列示相應日收盤價資料與授權條款；來源及再利用說明見專案 README 與 data/README。")
section_heading("3.2 變數建構", 2)
body("令 Pₐ,ₜ、Pᵦ,ₜ 為日期 t 的收盤價。以訓練樣本迴歸 ln(Pₐ,ₜ) 對 ln(Pᵦ,ₜ)，僅由訓練期估計截距 α 與斜率 β，再套用於整段價格序列：")
equation("xₜ = ln(Pₐ,ₜ) − α − β ln(Pᵦ,ₜ)；　Δxₜ₊₁ = xₜ₊₁ − xₜ")
body("預測特徵為 xₜ、上一期已觀察價差變化 Δxₜ，以及以前 20 期價差變化計算的移動標準差。未使用全資料標準化。預測採逐日、一步向前設定：對後續測試資訊日，前一個測試日的實際變化在該日已可觀察，故可作為 lag 與移動波動特徵；本設計不是在測試起點一次產生 24 期的固定起點預測。")
section_heading("3.3 時間切分與邊界處理", 2)
body("依日期排序，以 60%／20%／20% 切分價格序列。原始切分索引為訓練 73 日（2026-05-04 至 2026-08-14）、驗證 25 日（2026-08-17 至 2026-09-18）、測試 25 日（2026-09-21 至 2026-10-30）。建立監督式樣本時，排除滯後特徵不足及標籤跨越切分邊界的樣本；最後得到訓練 67 筆、驗證 24 筆、測試 24 筆。測試預測資訊日為 2026-09-21 至 2026-10-29，對應目標日為 2026-09-22 至 2026-10-30。三個模型均在同一組 24 個目標上評分。")
section_heading("3.4 比較模型", 2)
for title, text in [
    ("無變化基準", "預測 Δx̂ₜ₊₁=0，亦即價差水準預測維持 xₜ。"),
    ("線性回復／OU 啟發", "以訓練樣本估計 Δx̂ₜ₊₁ = a + b xₜ；僅當 0<φ=1+b<1 時報告 OU 啟發半衰期。"),
    ("PySR 替代公式", f"PySR 使用訓練資料搜尋候選方程，以驗證 MAE 選式。本次選式為 Δx̂ₜ₊₁ = {formula['selected_equation']}。")]:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(.2)
    p.paragraph_format.keep_together = True
    r = p.add_run(title + "。")
    set_run_font(r, 10.5, bold=True, color=TEAL)
    p.add_run(text)
body("PySR 設定為版本 1.5.10、seed=42、20 次迭代、maxsize=16，使用 +、−、× 運算子，以單程序及固定亂數種子搜尋。選定公式保存為 JSON，並以安全算式重載計算器核對驗證預測；最大絕對差異為 4.34×10⁻¹⁹。介面展示時使用預存公式，標示為「公式預先搜尋，評估現場執行」。")
section_heading("3.5 評估指標", 2)
body("對同一組測試目標計算平均絕對誤差（MAE）與均方根誤差（RMSE）。相對基準 MAE 改善定義為 1 − MAE_model/MAE_baseline；正值表示誤差下降，負值表示誤差上升。指標只描述本研究目標的預測誤差，不作為因果或投資績效估計。")
equation("MAE = n⁻¹ Σ |yₜ − ŷₜ|；　RMSE = √[n⁻¹ Σ (yₜ − ŷₜ)²]")

section_heading("四、實驗結果", 1)
section_heading("4.1 訓練期參數與模型診斷", 2)
body("訓練期價格關係估計為 α=6.3063、β=0.1759。線性價差變化模型估計 a=0.000821、b=−0.236768，故 φ=0.763232，符合本研究事先指定的 0<φ<1 診斷條件；相應 OU 啟發半衰期為 2.57 個交易步。此結果支持 H1 的訓練期參數條件，但不是正式共整合檢定，亦不代表樣本外預測一定改善。")
section_heading("4.2 保留測試集表現", 2)
doc.add_picture(str(error_figure), width=Inches(6.25))
caption("圖 1　測試集 MAE 與 RMSE。樣本為官方固定價格資料；每一指標以同一組 24 個一步預測目標計算。")
table = doc.add_table(rows=1, cols=5)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.autofit = True
table.style = "Light Shading Accent 1"
repeat_table_header(table.rows[0])
headers = ["模型", "MAE", "RMSE", "相對基準 MAE 改善", "驗證選式"]
for cell, value in zip(table.rows[0].cells, headers):
    cell.text = value
    shade_cell(cell, PALE)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for p in cell.paragraphs:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in p.runs: set_run_font(r, 9, bold=True, color=NAVY)
for name, row in metrics.iterrows():
    values = [name, f"{row.MAE:.6f}", f"{row.RMSE:.6f}", "—" if name == "無變化基準" else f"{row['相對基準改善']:.2%}", "—" if name != "PySR 替代公式" else "是"]
    cells = table.add_row().cells
    for cell, value in zip(cells, values):
        cell.text = value
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs: set_run_font(r, 9, color=NAVY)
caption("表 1　模型測試集預測誤差與相對基準改善。PySR 以驗證集 MAE 選定；測試集不參與選式。")
body("無變化基準的 MAE 為 0.008493，OU 啟發模型為 0.008930（相對改善 −5.15%），PySR 為 0.010204（−20.16%）。三者 RMSE 分別為 0.010186、0.010853 與 0.012943，排序亦由基準最佳。據此，H2 在本次測試期間未獲支持；此為指定測試樣本上的比較結果，非對所有股票或期間的普遍判斷。")
doc.add_picture(str(prediction_figure), width=Inches(6.2))
caption("圖 2　測試期間實際與預測價差變化。橫軸為預測資訊日 t；各點預測 t 的下一共同交易日目標。")

section_heading("五、討論", 1)
body("訓練期係數呈現回復方向，但基準模型在測試 MAE 與 RMSE 均優於候選模型。這說明「訓練期線性係數符合回復條件」與「回復模型能改善樣本外預測」是不同命題。候選模型估計出的條件方向未能在本次短期測試中轉化成較低誤差；PySR 的非線性公式亦未展現相對基準優勢。可能解釋包括樣本有限、價差關係隨時間變動、候選特徵資訊不足，或價格噪音高於短期可預測訊號；本研究未設計程序去識別這些解釋，因此它們僅是後續待驗證的可能性。")
body("實驗的設計重點是保留測試集隔離：價差係數只由訓練期估計，PySR 只在訓練資料搜尋並以驗證資料選式，測試樣本不參與選股、調參或選式。逐日一步預測則假設每一日可取得前一日已發生的變化資料。未來若採固定起點多步預測，特徵更新方式與誤差可能不同，應另行評估。")

section_heading("六、限制、結論與後續研究", 1)
body("本研究使用單一股票配對、123 個共同交易日及 24 個測試目標；價格未調整股利及股票分割。短驗證期與短測試期令結果不確定性較高，且未估計時間相依誤差區間、未執行正式共整合檢定，也未作跨配對或多時段外部驗證。故結論只適用於此固定資料版本與研究設定。")
body("在本次樣本中，線性 OU 啟發係數符合訓練期回復條件，但無變化基準在測試誤差上最佳；PySR 公式也未降低基準 MAE。下一步應先登錄股票配對、預測目標、切分方案與評估規則，再擴大到多配對及多個滾動測試區間，報告係數和方程穩定性及時間相依不確定性。另可比較調整後價格及不同預測更新方案。")

section_heading("參考文獻", 1)
references = [
    "Cranmer, M. (2023). Interpretable machine learning for science with PySR and SymbolicRegression.jl. arXiv:2305.01582. https://doi.org/10.48550/arXiv.2305.01582",
    "Engle, R. F., & Granger, C. W. J. (1987). Co-integration and error correction: Representation, estimation, and testing. Econometrica, 55(2), 251–276. https://doi.org/10.2307/1913236",
    "臺灣證券交易所（2026）。個股日收盤價資料（STOCK_DAY）[資料集]。政府資料開放平臺資料集 11548。https://data.gov.tw/en/datasets/11548",
    "Uhlenbeck, G. E., & Ornstein, L. S. (1930). On the theory of the Brownian motion. Physical Review, 36(5), 823–841. https://doi.org/10.1103/PhysRev.36.823",
]
for ref in references:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(.28)
    p.paragraph_format.first_line_indent = Inches(-.28)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_together = True
    r = p.add_run(ref)
    set_run_font(r, 9.5, color=NAVY)
doc.save(DOCS / "research_report.docx")

# Formal academic presentation: seven slides, one research narrative, cited source footer.
prs = Presentation()
prs.slide_width, prs.slide_height = PptInches(13.333), PptInches(7.5)
BG = PptColor(250, 249, 246)
INK = PptColor(25, 49, 61)
TEAL = PptColor(20, 124, 120)
MUTED_PPT = PptColor(99, 116, 122)
PALE_PPT = PptColor(232, 241, 239)
ORANGE = PptColor(198, 107, 67)
WHITE = PptColor(255, 255, 255)


def add_text(slide, x, y, w, h, text, size=18, color=INK, bold=False, align=PP_ALIGN.LEFT, font="PingFang TC"):
    # Presentation rule: all visible text must be at least 12 pt.
    size = max(12, size)
    shape = slide.shapes.add_textbox(PptInches(x), PptInches(y), PptInches(w), PptInches(h))
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = PptInches(.04)
    tf.margin_top = tf.margin_bottom = PptInches(.03)
    for i, line in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = line
        p.alignment = align
        p.space_after = PptPt(8)
        p.font.name = font
        p.font.size = PptPt(size)
        p.font.bold = bold
        p.font.color.rgb = color
    return shape


def add_slide(title, number, section_label="研究發表"):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = BG
    add_text(slide, .65, .25, 8, .32, f"股票也有回復力嗎？　|　{section_label}", 12, TEAL, True)
    add_text(slide, .65, .68, 12, .62, title, 26, INK, True)
    rule = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptInches(.65), PptInches(1.38), PptInches(12), PptInches(.025))
    rule.fill.solid(); rule.fill.fore_color.rgb = TEAL; rule.line.fill.background()
    add_text(slide, .65, 6.95, 12, .35, f"{number:02d} / 05", 12, MUTED_PPT, align=PP_ALIGN.RIGHT)
    return slide


def add_panel(slide, x, y, w, h, heading, text, accent=TEAL, body_size=16):
    box = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PptInches(x), PptInches(y), PptInches(w), PptInches(h))
    box.fill.solid(); box.fill.fore_color.rgb = WHITE
    box.line.color.rgb = PptColor(218, 226, 223)
    add_text(slide, x+.2, y+.18, w-.4, .4, heading, 16, accent, True)
    add_text(slide, x+.2, y+.68, w-.4, h-.83, text, body_size, INK)
    return box

# 1 Title
s = prs.slides.add_slide(prs.slide_layouts[6]); s.background.fill.solid(); s.background.fill.fore_color.rgb = BG
add_text(s, .85, .65, 11, .35, "期中專題研究發表", 12, TEAL, True)
add_text(s, .85, 1.72, 11.5, 1.15, "股票也有回復力嗎？", 36, INK, True)
add_text(s, .88, 3.02, 10.5, .65, "物理模型與 AI 公式的比較實驗", 23, MUTED_PPT)
add_text(s, .9, 4.18, 10.9, .75, "整理兩股的相對價差，並比較三種方法對下一交易日變化的預測。", 17, INK)
bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PptInches(.9), PptInches(5.42), PptInches(1.55), PptInches(.06)); bar.fill.solid(); bar.fill.fore_color.rgb=TEAL; bar.line.fill.background()
add_text(s, .9, 5.82, 10.8, .42, "物理三　｜　林映霖", 14, MUTED_PPT)
add_text(s, .65, 6.95, 12, .35, "01 / 05", 12, MUTED_PPT, align=PP_ALIGN.RIGHT)

# 2 Sample selection and purpose
s=add_slide("樣本選擇｜為什麼用這組股票？",2)
add_text(s,.95,1.58,11.3,.45,"研究目的：檢查兩股的相對價差，能否預測下一交易日的價差變化。",16,INK,True)
add_panel(s,.8,2.25,3.8,3.2,"樣本｜台積電 × 聯發科","使用 TWSE 日收盤價，期間為 2026 年 5 月至 10 月，共 123 個共同交易日。",TEAL,15)
add_panel(s,4.78,2.25,3.8,3.2,"選樣理由｜產業相關","兩家公司都屬半導體產業；訓練期日報酬相關係數為 0.51，顯示一定程度的共同漲跌。",ORANGE,15)
add_panel(s,8.76,2.25,3.8,3.2,"研究假設｜價差可能回復","若相對價差偏離後，隔日變化朝相反方向移動，就呈現回復傾向。這是要用資料檢驗的假設。",TEAL,15)
add_text(s,.95,5.9,11.3,.55,"共同漲跌不等於長期均衡；本研究未做共整合檢定，也不預設價差必定回復。",12,INK,True)

# 3 Explain why prices are transformed, plus the available inputs
s=add_slide("用比例整理股價，做出可預測的目標",3)
add_text(s,.95,1.48,11.3,.5,"先用比例整理兩檔股票的收盤價，算出每天的相對價差；模型才能預測這個價差隔天會怎麼變。",16,INK,True)
add_text(s,.95,2.12,11.3,.35,"資料可從四種來源取得：TWSE 離線樣本　｜　模擬資料　｜　TWSE 線上查詢　｜　本機 CSV",14,TEAL,True)
add_panel(s,.8,2.85,3.8,2.75,"輸入｜兩檔股票收盤價","選擇兩檔股票及資料期間，取得每日收盤價。",MUTED_PPT,15)
add_panel(s,4.78,2.85,3.8,2.75,"整理｜用比例算相對價差","比例用來整理兩股的關係，不是最後要預測的答案。",TEAL,15)
add_panel(s,8.76,2.85,3.8,2.75,"預測｜下一日價差變化","讓模型回答明確問題：明天的相對價差會增加還是減少、改變多少？",ORANGE,15)
add_text(s,.95,5.95,11.3,.62,"這樣轉換，是為了把股價變成模型能學習、也能用後續資料檢查的預測目標。股票、期間、時間切分與搜尋設定都可調整。",14,INK,True)

# 4 Model comparison in a clean three-row layout
s=add_slide("三種模型｜本名、白話意思和差別",4)
add_text(s,.95,1.58,11.3,.5,"它們都預測下一天的相對價差變化，但使用不同的猜法。",16,INK,True)
for y, accent, title, formal_name, explanation in [
    (2.18, MUTED_PPT, "無變化基準", "No-change baseline", "假設明天和今天一樣；提供最簡單的比較起點。"),
    (3.43, TEAL, "OU 啟發線性回復模型", "Ornstein–Uhlenbeck-inspired linear mean-reversion model", "價差偏離平常水準時，猜它會按線性規則往回一些。"),
    (4.68, ORANGE, "PySR 符號回歸", "PySR Symbolic Regression", "程式從資料試出預測公式，再用驗證資料挑選。"),
]:
    row=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,PptInches(.8),PptInches(y),PptInches(11.7),PptInches(1.05))
    row.fill.solid(); row.fill.fore_color.rgb=WHITE; row.line.color.rgb=PptColor(218,226,223)
    stripe=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,PptInches(.8),PptInches(y),PptInches(.075),PptInches(1.05))
    stripe.fill.solid(); stripe.fill.fore_color.rgb=accent; stripe.line.fill.background()
    add_text(s,1.05,y+.13,4.5,.3,title,16,accent,True)
    add_text(s,1.05,y+.49,4.5,.4,formal_name,12,INK)
    add_text(s,5.85,y+.2,6.3,.62,explanation,15,INK)
add_text(s,.95,6.02,11.2,.42,"差別：照舊不變｜依線性想法猜回復｜讓程式搜尋公式。",14,INK,True)

# 5 MAE as the final comparison tool and why it is appropriate
s=add_slide("最後比較工具｜MAE 越小越好",5)
add_text(s,.95,1.58,11.3,.5,"MAE（平均絕對誤差）看的是：每天預測和實際差多遠，再把差距取平均。",16,INK,True)
add_panel(s,.8,2.35,3.8,1.62,"① 算每天差距","預測值和實際值相差多少。",MUTED_PPT,15)
add_panel(s,4.78,2.35,3.8,1.62,"② 不分高估或低估","都看差距大小，避免正負抵銷。",TEAL,15)
add_panel(s,8.76,2.35,3.8,1.62,"③ 把差距取平均","平均後得到 MAE。",ORANGE,15)
add_text(s,.95,4.35,11.3,.62,"判讀方式：MAE 越小，代表預測平均越接近實際。",22,TEAL,True)
add_text(s,.95,5.25,11.2,.82,"為什麼可比較？三種模型預測同一目標，並在相同測試日期評分；所以 MAE 可以直接比較誰平均猜得較準。",16,INK,True)

prs.save(SLIDES / "midterm_90s.pptx")

script = """# 90 秒操作展示講稿

**0–12 秒｜樣本與目的** 本專題用台積電與聯發科的日收盤價，研究兩股的相對價差能不能預測下一交易日變化。訓練期報酬有一定同步，但這不代表價差必然回復。

**12–27 秒｜資料來源與設定** 工具可選四種來源：TWSE 離線樣本、程式生成的模擬股價、TWSE 線上資料和本機 CSV。股票期間、訓練與驗證比例、模擬種子及公式搜尋範圍都可以調整。

**27–42 秒｜股價轉成預測目標** 工具先用訓練期估計兩股價格關係，計算每天的相對價差，再把目標定為下一交易日價差變化。資料按時間分為訓練、驗證、測試，避免未來資訊進入建模。

**42–57 秒｜比較方法** 比較不變基準、平均回復模型與 PySR 搜尋公式。三種方法預測同一目標、使用相同測試日期；公式搜尋與挑選不使用測試資料。

**57–75 秒｜閱讀分析報告** 工具產生實際值與預測圖、誤差指標、公式解釋、結論、實驗紀錄和逐日 CSV。本次測試由不變基準誤差最低，另外兩種方法沒有勝過它。

**75–90 秒｜MAE 與收尾** MAE 是每天預測與實際差距取絕對值後的平均。因三種方法預測同一件事，也在相同測試日期比較，可用它看哪個方法平均更接近實際值。
"""
(SLIDES / "90s_script.md").write_text(script, encoding="utf-8")
print("Generated academic report, 5-slide presentation, and 90-second script.")
