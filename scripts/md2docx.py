# -*- coding: utf-8 -*-
"""
md2docx.py —— 教学材料 Markdown → Word(.docx) 批量转换

做三件事：
1. 用 pandoc 把 Markdown 转成 docx（保留标题层级、表格、列表结构）；
2. 用 python-docx 后处理，统一全局字体为宋体（含标题、表格内文字），
   标题按层级设字号，颜色归黑；
3. 给所有表格补齐边框，并把首行设为重复表头（跨页自动重复）。

落实《职业教育编写规范》第 8 条「全文统一使用宋体，包括标题」。

关于板书：教案「板书设计」章是 ASCII 树形图，在 Markdown 中是代码块，
转换后为 Source Code 样式。默认同样设宋体；若在 Word 中树形图对齐错乱，
可加 `--code-font Consolas` 改用等宽字体渲染该区域（此时校验脚本
verify_docx.py 须带同名参数，否则代码块会被误判为非宋体）。

依赖：pandoc（命令行可用）、python-docx
用法：
    python md2docx.py "a.md" "b.md"              # 原地输出同名 .docx
    python md2docx.py --out dist "a.md" "b.md"   # 输出到指定目录
    python md2docx.py --code-font Consolas "教案.md"
"""
import argparse
import shutil
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

FONT = "宋体"
# 标题层级 → 字号（pt）。铁律第 8 条：标题靠字号与加粗区分，不靠字体种类。
HEADING_SIZES = {
    "Title": 18,
    "Heading 1": 16,
    "Heading 2": 14,
    "Heading 3": 12,
    "Heading 4": 12,
    "Heading 5": 12,
    "Heading 6": 12,
    "Heading 7": 12,
    "Heading 8": 12,
}
CODE_STYLES = ("Source Code", "Preformatted Text")


def set_run_font(run, size=None, bold=None, font=FONT):
    """把单个 run 设为指定字体（默认宋体，中英文都设），可选设字号/加粗，颜色归黑。"""
    run.font.name = font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    for attr in ("w:eastAsia", "w:ascii", "w:hAnsi"):
        rFonts.set(qn(attr), font)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)


def set_table_borders(table):
    """补齐表格边框（pandoc 默认表格可能无边框，打印时会散）。"""
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)


def set_repeat_header(row):
    """首行设为重复表头（跨页时自动重复）。"""
    trPr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    trPr.append(el)


def unify(path: Path, code_font: str = FONT):
    """统一字体：正文/标题/表格内文字全部宋体，标题按层级设字号。"""
    doc = Document(str(path))

    for p in doc.paragraphs:
        style_name = p.style.name if p.style is not None else ""
        if style_name in CODE_STYLES:
            for r in p.runs:
                set_run_font(r, size=10, font=code_font)
            continue
        size = HEADING_SIZES.get(style_name)
        for r in p.runs:
            set_run_font(r, size=size, bold=(True if size else None))

    for t in doc.tables:
        set_table_borders(t)
        if len(t.rows) > 0:
            set_repeat_header(t.rows[0])
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for r in p.runs:
                        set_run_font(r)

    doc.save(str(path))


def docx_name_of(md_path: Path) -> str:
    """仅替换结尾的 .md——用 with_suffix 会把 "V1.2.md" 误截成 "V1.docx"。"""
    name = md_path.name
    return name[:-3] + ".docx" if name.lower().endswith(".md") else name + ".docx"


def convert(md_path: Path, out_dir: Path | None, code_font: str = FONT) -> Path:
    docx_path = out_dir / docx_name_of(md_path) if out_dir else md_path.with_name(docx_name_of(md_path))
    if out_dir:
        out_dir.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        ["pandoc", str(md_path), "-o", str(docx_path), "--from=gfm", "--to=docx"],
        check=True,
    )
    unify(docx_path, code_font=code_font)
    return docx_path


def main():
    ap = argparse.ArgumentParser(description="Markdown → DOCX（统一宋体）")
    ap.add_argument("files", nargs="+", help="待转换的 .md 文件")
    ap.add_argument("--out", default=None, help="输出目录（默认与源文件同目录）")
    ap.add_argument("--code-font", default=FONT, help=f"代码块（板书树形图）字体，默认 {FONT}")
    args = ap.parse_args()

    if shutil.which("pandoc") is None:
        sys.exit("未找到 pandoc，请先安装：https://pandoc.org/installing.html")

    out_dir = Path(args.out) if args.out else None
    for f in args.files:
        src = Path(f)
        if not src.exists():
            print(f"[跳过] 不存在：{src}")
            continue
        dst = convert(src, out_dir, code_font=args.code_font)
        print(f"[完成] {src.name} → {dst}")


if __name__ == "__main__":
    main()
