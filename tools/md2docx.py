import re, sys
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT

FONT = "맑은 고딕"

def set_font(run, size=None, bold=None, color=None, mono=False):
    f = "Consolas" if mono else FONT
    run.font.name = f
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    if size: run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if color: run.font.color.rgb = RGBColor.from_string(color)

def add_inline(par, text, size=10, base_bold=False):
    parts = re.split(r"(\*\*[^*]+\*\*|`[^`]+`)", text)
    for p in parts:
        if not p: continue
        if p.startswith("**") and p.endswith("**"):
            set_font(par.add_run(p[2:-2]), size, True)
        elif p.startswith("`") and p.endswith("`"):
            set_font(par.add_run(p[1:-1]), size, base_bold, "1F4E79", mono=True)
        else:
            set_font(par.add_run(p.replace("\\*", "*")), size, base_bold)

def shade(cell, hexcolor):
    tcPr = cell._element.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): hexcolor})
    tcPr.append(shd)

def flush_table(doc, rows):
    rows = [r for r in rows if not re.match(r"^\|\s*:?-{2,}", r)]
    cells = [[c.strip() for c in r.strip().strip("|").split("|")] for r in rows]
    ncol = max(len(c) for c in cells)
    t = doc.add_table(rows=len(cells), cols=ncol)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(cells):
        for j in range(ncol):
            cell = t.cell(i, j)
            cell.text = ""
            add_inline(cell.paragraphs[0], row[j] if j < len(row) else "", 9, base_bold=(i == 0))
            if i == 0: shade(cell, "D9E2F3")
    doc.add_paragraph()

def convert(src, dst):
    doc = Document()
    st = doc.styles["Normal"]; st.font.name = FONT; st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT); st.font.size = Pt(10)
    lines = open(src, encoding="utf-8").read().splitlines()
    table, code = [], None
    for line in lines:
        if code is not None:
            if line.startswith("```"):
                p = doc.add_paragraph(); set_font(p.add_run("\n".join(code)), 9, mono=True); code = None
            else: code.append(line)
            continue
        if line.startswith("```"): code = []; continue
        if line.startswith("|"):
            table.append(line); continue
        if table: flush_table(doc, table); table = []
        s = line.rstrip()
        if not s or s == "---": continue
        m = re.match(r"^(#{1,4})\s+(.*)", s)
        if m:
            lvl = len(m.group(1)); h = doc.add_heading(level=min(lvl, 3) if lvl > 1 else 0)
            set_font(h.add_run(m.group(2)), {1: 20, 2: 15, 3: 12, 4: 11}[lvl], True, "1F3864")
            continue
        if s.startswith(">"):
            p = doc.add_paragraph(); p.paragraph_format.left_indent = Pt(14)
            add_inline(p, s.lstrip("> ").strip(), 9.5); continue
        m = re.match(r"^(\s*)[-*]\s+(.*)", s)
        if m:
            p = doc.add_paragraph(style="List Bullet"); p.paragraph_format.left_indent = Pt(18 + len(m.group(1)) * 6)
            add_inline(p, m.group(2)); continue
        m = re.match(r"^(\s*)(\d+)\.\s+(.*)", s)
        if m:
            p = doc.add_paragraph(); p.paragraph_format.left_indent = Pt(14 + len(m.group(1)) * 6)
            add_inline(p, f"{m.group(2)}. {m.group(3)}"); continue
        add_inline(doc.add_paragraph(), s.strip())
    if table: flush_table(doc, table)
    doc.save(dst)

for src in sys.argv[1:]:
    convert(src, src[:-3] + ".docx"); print("saved", src[:-3] + ".docx")
