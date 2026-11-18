"""PowerPoint 산출물: 법정 교과목별 교안 8종."""
import re

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

from common import FONT

NAVY = RGBColor(0x1F, 0x38, 0x64)
ORANGE = RGBColor(0xE8, 0x6A, 0x10)
ICE = RGBColor(0xE8, 0xF1, 0xFA)
GRAY = RGBColor(0x59, 0x59, 0x59)
LIGHT = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
RED = RGBColor(0xC0, 0x00, 0x00)

W, H = Emu(12192000), Emu(6858000)  # 16:9
IN = 914400


def _tb(slide, x, y, w, h, lines, size=18, color=GRAY, bold=False, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, bullet=None, spacing=6):
    box = slide.shapes.add_textbox(int(x), int(y), int(w), int(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(lines, str):
        lines = [lines]
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.space_after = Pt(spacing)
        r = p.add_run()
        r.text = (bullet + " " if bullet else "") + line
        r.font.name = FONT
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = color
        rpr = r._r.get_or_add_rPr()
        ea = rpr.makeelement("{http://schemas.openxmlformats.org/drawingml/2006/main}ea", {"typeface": FONT})
        rpr.append(ea)
    return box


def _rect(slide, x, y, w, h, fill, shape=MSO_SHAPE.RECTANGLE):
    s = slide.shapes.add_shape(shape, int(x), int(y), int(w), int(h))
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    s.shadow.inherit = False
    if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = min(0.5, 0.12 * IN / max(1, min(int(w), int(h))))
    return s


def _autosize(lines, base, min_size=12, budget=9):
    """줄 수·길이에 따라 글자 크기를 줄여 넘침을 막는다."""
    weight = sum(max(1, len(l) / 42) for l in lines)
    size = base
    while size > min_size and weight * (size / base) > budget:
        size -= 1
        weight = sum(max(1, len(l) / (42 * base / size)) for l in lines)
    return size


class Deck:
    def __init__(self, ctx, subject_no):
        self.ctx = ctx
        self.no = subject_no
        self.subject = ctx.subjects[subject_no]
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.blank = self.prs.slide_layouts[6]
        self.page = 0

    def _base(self, heading, kicker=None):
        s = self.prs.slides.add_slide(self.blank)
        self.page += 1
        _rect(s, 0, 0, W, int(0.08 * IN), NAVY)
        _rect(s, int(0.55 * IN), int(0.45 * IN), int(0.09 * IN), int(0.55 * IN), ORANGE)
        if kicker:
            _tb(s, 0.8 * IN, 0.3 * IN, 10 * IN, 0.35 * IN, kicker, 12, ORANGE, True)
        _tb(s, 0.8 * IN, 0.55 * IN, 11.5 * IN, 0.7 * IN, heading, 28, NAVY, True)
        _tb(s, 0.55 * IN, 7.05 * IN, 11 * IN, 0.35 * IN,
            f"정기교육 {self.no}. {self.subject['short']}  |  {self.ctx.laws['regular_subjects_basis']}", 9, GRAY)
        _tb(s, 12.3 * IN, 7.05 * IN, 0.8 * IN, 0.35 * IN, str(self.page), 10, GRAY, align=PP_ALIGN.RIGHT)
        return s

    def cover(self, hours):
        s = self.prs.slides.add_slide(self.blank)
        self.page += 1
        _rect(s, 0, 0, W, H, NAVY)
        _rect(s, 0, int(4.9 * IN), W, int(2.6 * IN), RGBColor(0x17, 0x2B, 0x4D))
        _rect(s, int(0.8 * IN), int(1.2 * IN), int(0.12 * IN), int(1.9 * IN), ORANGE)
        _tb(s, 1.1 * IN, 1.0 * IN, 11 * IN, 0.5 * IN, f"{self.ctx.half_label} 제설기간제근로자 정기 안전보건교육", 16, RGBColor(0xBF, 0xD4, 0xEE), True)
        _tb(s, 1.1 * IN, 1.5 * IN, 11.3 * IN, 1.8 * IN, f"과목 {self.no}. {self.subject['short']}", 40, WHITE, True)
        _tb(s, 1.1 * IN, 3.25 * IN, 11 * IN, 1.2 * IN, "법정 교과목: " + self.subject["name"], 14, RGBColor(0xDD, 0xE6, 0xF2))
        _tb(s, 1.1 * IN, 5.2 * IN, 11 * IN, 1.6 * IN, [
            f"교육시간  {hours}시간",
            f"강사  {self.ctx.cfg['instructor_role']}",
            f"근거  {self.ctx.laws['hours']['regular_basis']} / {self.ctx.laws['regular_subjects_basis']}",
        ], 13, WHITE, spacing=4)

    def bullets(self, heading, items, kicker=None, base=22, highlight=None):
        s = self._base(heading, kicker)
        size = _autosize(items, base)
        if highlight:
            _rect(s, int(0.8 * IN), int(1.45 * IN), int(11.7 * IN), int(0.75 * IN), ICE, MSO_SHAPE.ROUNDED_RECTANGLE)
            _tb(s, 1.0 * IN, 1.5 * IN, 11.3 * IN, 0.65 * IN, highlight, 16, NAVY, True, anchor=MSO_ANCHOR.MIDDLE)
            top = 2.4
        else:
            top = 1.6
        _tb(s, 0.9 * IN, top * IN, 11.5 * IN, (6.8 - top) * IN, items, size, GRAY, bullet="▪", spacing=10)
        return s

    def two_col(self, heading, left_title, left, right_title, right, kicker=None):
        s = self._base(heading, kicker)
        for k, (t, items, color) in enumerate([(left_title, left, RED), (right_title, right, NAVY)]):
            x = 0.8 + k * 6.0
            _rect(s, int(x * IN), int(1.6 * IN), int(5.7 * IN), int(5.2 * IN), LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE)
            _rect(s, int(x * IN), int(1.6 * IN), int(5.7 * IN), int(0.6 * IN), color, MSO_SHAPE.ROUNDED_RECTANGLE)
            _tb(s, (x + 0.2) * IN, 1.62 * IN, 5.3 * IN, 0.55 * IN, t, 17, WHITE, True, anchor=MSO_ANCHOR.MIDDLE)
            size = _autosize(items, 17, 11, budget=7)
            _tb(s, (x + 0.2) * IN, 2.35 * IN, 5.3 * IN, 4.3 * IN, items, size, GRAY, bullet="•", spacing=8)
        return s

    def case(self, text):
        s = self._base("사례로 보는 위험", "CASE")
        _rect(s, int(0.8 * IN), int(1.7 * IN), int(11.7 * IN), int(3.0 * IN), ICE, MSO_SHAPE.ROUNDED_RECTANGLE)
        _tb(s, 1.2 * IN, 1.9 * IN, 11.0 * IN, 2.6 * IN, text, 20, NAVY, anchor=MSO_ANCHOR.MIDDLE)
        _tb(s, 0.9 * IN, 5.0 * IN, 11.5 * IN, 1.5 * IN, [
            "토의: 우리 작업장에서 같은 일이 생길 수 있는 곳은 어디인가?",
            "토의: 무엇을 바꾸면 막을 수 있었는가?",
        ], 16, GRAY, bullet="Q.", spacing=8)

    def table(self, heading, rows, kicker=None):
        s = self._base(heading, kicker)
        nr, nc = len(rows), len(rows[0])
        gt = s.shapes.add_table(nr, nc, int(0.8 * IN), int(1.7 * IN), int(11.7 * IN), int(0.6 * IN * nr)).table
        for i in range(nr):
            gt.rows[i].height = int((0.6 if i == 0 else 1.0) * IN)
        for i, row in enumerate(rows):
            for j, v in enumerate(row):
                c = gt.cell(i, j)
                c.text = v
                p = c.text_frame.paragraphs[0]
                for r in p.runs:
                    r.font.name = FONT
                    r.font.size = Pt(14)
                    r.font.bold = i == 0
                    r.font.color.rgb = WHITE if i == 0 else GRAY
                c.fill.solid()
                c.fill.fore_color.rgb = NAVY if i == 0 else (WHITE if i % 2 else LIGHT)
        return s

    def quiz(self, qa):
        s = self._base("확인 퀴즈", "CHECK")
        y = 1.6
        for i, item in enumerate(qa, 1):
            _rect(s, int(0.8 * IN), int(y * IN), int(11.7 * IN), int(1.45 * IN), LIGHT, MSO_SHAPE.ROUNDED_RECTANGLE)
            _tb(s, 1.0 * IN, (y + 0.1) * IN, 11.3 * IN, 0.6 * IN, f"Q{i}. {item['q']}", 18, NAVY, True)
            _tb(s, 1.0 * IN, (y + 0.7) * IN, 11.3 * IN, 0.6 * IN, f"A. {item['a']}", 16, ORANGE)
            y += 1.65
        return s

    def save(self):
        fname = re.sub(r"[\\/:*?\"<>|·,]", "", self.subject["short"]).replace(" ", "")
        path = self.ctx.path(f"03_교안_과목{self.no}_{fname}.pptx")
        self.prs.save(path)
        return path


def build_subject(ctx, no):
    c = ctx.cur[str(no)]
    d = Deck(ctx, no)
    hours = ctx.hours_by_subject()[no]
    d.cover(hours)
    d.bullets("학습 목표", c["goals"], "GOAL", base=24)
    d.bullets("법령 요점", c["law_points"], "LAW", base=20, highlight="법정 교과목: " + ctx.subjects[no]["name"])
    d.bullets("제설작업에서는 이렇게 나타납니다", c["hazards"], "HAZARD")
    d.case(c["case"])
    d.bullets("예방 수칙", c["rules"], "RULE")
    if "fire" in c:
        d.two_col("화재·폭발 사고 시 대피", "위험 상황", ["유류·가스 누출 또는 장비 과열", "제설기지 보관창고 화재", "연기 확산"],
                  "대피 요령", c["fire"], "EVACUATION")
    if "first_aid" in c:
        d.two_col("한파작업 건강장해 응급조치", "이런 증상이면", c["hazards"][:2], "이렇게 조치", c["first_aid"], "FIRST AID")
    if "worksheet" in c:
        d.table("실습: 제설작업 위험성평가", c["worksheet"], "PRACTICE")
    if "msds" in c:
        d.bullets("물질안전보건자료(MSDS) 꼭 볼 항목", c["msds"], "MSDS")
    d.quiz(c["quiz"])
    d.bullets("오늘의 핵심", [c["rules"][0], c["rules"][1], c["law_points"][0]], "SUMMARY", base=22)
    return d.save()


def build_all(ctx):
    return [build_subject(ctx, no) for no in ctx.subjects]
