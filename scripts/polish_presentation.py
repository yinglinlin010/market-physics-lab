"""Apply the project's restrained academic visual system to its 90-second deck."""
from __future__ import annotations

from pathlib import Path
import sys

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "slides" / "midterm_90s.pptx"
INK = RGBColor(23, 42, 53)
RULE = RGBColor(220, 226, 221)
PAPER = RGBColor(246, 244, 238)
WHITE = RGBColor(255, 255, 255)
PALE = RGBColor(233, 243, 241)


def remove_shape_effects(shape) -> None:
    """Remove inherited Office theme effects that create unintended shadows."""
    element = shape._element
    for node in list(element):
        if node.tag.endswith("}style"):
            element.remove(node)
    sp_pr = element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}spPr")
    if sp_pr is not None:
        for node in list(sp_pr):
            if node.tag.endswith("}effectLst") or node.tag.endswith("}effectRef"):
                sp_pr.remove(node)


def set_cell_fill(cell, color: RGBColor) -> None:
    cell.fill.solid()
    cell.fill.fore_color.rgb = color


def polish(input_path: Path, output_path: Path) -> None:
    prs = Presentation(input_path)
    cover = prs.slides[0]
    if not any(getattr(shape, "has_text_frame", False) and shape.text.strip() == "物理三　｜　林映霖" for shape in cover.shapes):
        identity = cover.shapes.add_textbox(Inches(.9), Inches(6.25), Inches(10.8), Inches(.35))
        identity.text_frame.margin_left = identity.text_frame.margin_right = Inches(.04)
        identity.text_frame.margin_top = identity.text_frame.margin_bottom = Inches(.03)
        paragraph = identity.text_frame.paragraphs[0]
        paragraph.text = "物理三　｜　林映霖"
        paragraph.font.name = "PingFang TC"
        paragraph.font.size = Pt(12)
        paragraph.font.color.rgb = RGBColor(95, 114, 122)

    for slide_number, slide in enumerate(prs.slides, start=1):
        shapes = list(slide.shapes)
        for i, shape in enumerate(shapes):
            remove_shape_effects(shape)
            if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                geometry = shape._element.find(
                    ".//{http://schemas.openxmlformats.org/drawingml/2006/main}prstGeom"
                )
                if geometry is not None and geometry.get("prst") == "roundRect":
                    # Academic panels retain their grouping but use a crisp, quiet edge.
                    geometry.set("prst", "rect")
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = WHITE
                    shape.line.color.rgb = RULE
                    shape.line.width = Pt(0.8)
                    # Panel text boxes immediately follow each panel in the source deck.
                    for text_shape in shapes[i + 1 : i + 3]:
                        if getattr(text_shape, "has_text_frame", False):
                            text_shape.left += Pt(2.4)
                            text_shape.width -= Pt(4.8)
            if getattr(shape, "has_table", False):
                table = shape.table
                for r, row in enumerate(table.rows):
                    for c, cell in enumerate(row.cells):
                        row_fill = PALE if r == 0 else (WHITE if r % 2 else PAPER)
                        set_cell_fill(cell, row_fill)
                        cell.margin_left = cell.margin_right = Pt(6)
                        cell.margin_top = cell.margin_bottom = Pt(5)
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.name = "PingFang TC"
                                run.font.color.rgb = INK
                                run.font.size = Pt(11 if r == 0 else 10.5)
                                run.font.bold = r == 0 or c == 0

        # Consistent text color and measurable reading sizes for source notes.
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text.strip()
            if not text:
                continue
            if slide_number == 3 and ("φ = 1 + b" in text or "H1 是訓練內參數診斷" in text):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(15)
            if slide_number == 8 and text.startswith("Engle & Granger"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10.5)
                        run.font.color.rgb = INK
            if text.startswith("註：相對改善"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                        run.font.color.rgb = RGBColor(95, 114, 122)
            if text.startswith("預測資訊日"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                        run.font.color.rgb = RGBColor(95, 114, 122)
            if text.startswith("PySR 方法參考"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
            if text.startswith("後續：事先固定"):
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(11)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)


if __name__ == "__main__":
    source = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else SOURCE
    destination = Path(sys.argv[2]).resolve() if len(sys.argv) > 2 else SOURCE
    polish(source, destination)
    print(f"Polished {len(Presentation(destination).slides)} slides: {destination}")
