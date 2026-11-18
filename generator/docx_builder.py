"""Word 산출물: 면제기준, 법정기준, 계획·결과 공문, 교육일지, 사진대지, README."""
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.shared import Pt, Cm, RGBColor

from common import FONT, BLANK, ph

NAVY = "1F3864"


# ---------- 공통 헬퍼 ----------
def _font(run, size=11, bold=False, color=None):
    run.font.name = FONT
    rpr = run._element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    return run


def new_doc(landscape=False):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), FONT)
    st.font.size = Pt(11)
    sec = doc.sections[0]
    if landscape:
        sec.orientation = WD_ORIENT.LANDSCAPE
        sec.page_width, sec.page_height = sec.page_height, sec.page_width
    for side in ("left_margin", "right_margin"):
        setattr(sec, side, Cm(2.0))
    sec.top_margin = Cm(2.0)
    sec.bottom_margin = Cm(1.8)
    return doc


def para(doc_or_cell, text="", size=11, bold=False, align=None, color=None, space_after=4, indent=0):
    p = doc_or_cell.add_paragraph()
    if text:
        _font(p.add_run(text), size, bold, color)
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif align == "right":
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    return p


def title(doc, text, size=18):
    return para(doc, text, size, True, "center", NAVY, space_after=10)


def h2(doc, text):
    return para(doc, text, 13, True, color=NAVY, space_after=4)


def _shade(cell, fill):
    tcPr = cell._element.get_or_add_tcPr()
    shd = tcPr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): fill})
    tcPr.append(shd)


def table(doc, rows, widths=None, header=True, size=10, height=None, first_col_shade=False):
    t = doc.add_table(rows=len(rows), cols=len(rows[0]))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j, val in enumerate(row):
            c = t.cell(i, j)
            c.text = ""
            c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            lines = str(val).split("\n")
            p = c.paragraphs[0]
            for k, line in enumerate(lines):
                if k:
                    p = c.add_paragraph()
                _font(p.add_run(line), size, bold=(header and i == 0) or (first_col_shade and j == 0))
            if header and i == 0:
                _shade(c, "D9E2F3")
                c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif first_col_shade and j == 0:
                _shade(c, "F2F2F2")
            if widths:
                c.width = Cm(widths[j])
        if height:
            t.rows[i].height = Cm(height)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def stamp(doc, ctx):
    p = para(doc, ctx.law_stamp, 8, color="7F7F7F", space_after=0)
    return p


def page_break(doc):
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)


def _approval_box(doc, ctx, height=1.2):
    line = ctx.P.get("approval_line") or ["담당", "", "", ""]
    heads = [x if x else BLANK for x in line]
    table(doc, [heads, [""] * len(heads)], widths=[3] * len(heads), size=9, height=height)


def _gongmun_head(doc, ctx, subject):
    para(doc, ph(ctx.P.get("org"), 6), 20, True, "center", space_after=12)
    table(doc, [
        ["수신", f"{ph(ctx.P.get('receiver'), 6)}"],
        ["(경유)", ph(ctx.P.get("via"), 4)],
        ["제목", subject],
    ], widths=[2.5, 14], header=False, size=11, first_col_shade=True)


def _gongmun_tail(doc, ctx):
    para(doc, "끝.", 11, align="right", space_after=16)
    para(doc, ph(ctx.P.get("org"), 6) + "장", 16, True, "center", space_after=16)
    _approval_box(doc, ctx)
    table(doc, [
        ["시행", f"{ph(ctx.P.get('dept'), 4)}-{ph(ctx.P.get('doc_no'), 3)} ({ph('', 5)})", "접수", ph("", 5)],
        ["주소", ph("", 10), "공개구분", "공개 / 부분공개 / 비공개"],
        ["전화", ph("", 5), "담당자", ph("", 4)],
    ], widths=[2, 7, 2.5, 5], header=False, size=9, first_col_shade=True)
    stamp(doc, ctx)


# ---------- 01 채용 시 교육 면제·감면 기준 ----------
def build_hiring_exemption(ctx):
    L = ctx.laws
    doc = new_doc()
    title(doc, "제설기간제근로자 채용 시 교육 면제·감면 기준")
    para(doc, f"{ctx.year}년 동절기 적용 / 특별교육(16시간)은 타 부서 실시", 10, align="center", color="595959", space_after=10)

    h2(doc, "1. 법정 교육시간")
    table(doc, [
        ["구분", "대상", "교육시간", "근거"],
        ["채용 시 교육", L["worker"]["type"], f"{L['hours']['hiring']}시간 이상", L["hours"]["hiring_basis"]],
        ["특별교육(참고)", "별표5 특별교육 대상작업 종사자", f"{L['hours']['special']}시간 이상", L["hours"]["special_basis"]],
    ], widths=[2.8, 5.2, 2.5, 6.5])

    h2(doc, "2. 면제·감면 조건")
    rows = [["구분", "조건", "효과", "근거", "입증서류"]]
    for e in L["hiring_exemptions"]:
        rows.append([e["id"], e["condition"], e["effect"], e["basis"], e["evidence"]])
    table(doc, rows, widths=[1.2, 5, 3.8, 3.3, 3.7], size=9)

    h2(doc, "3. 판정 시 유의사항")
    for n in L["hiring_exemption_notes"]:
        para(doc, "• " + n, 10, indent=0.4)

    h2(doc, "4. 채용 시 교육 법정 교육내용(참고)")
    para(doc, f"근거: {L['hiring_subjects_basis']}", 9, color="595959")
    for i, s in enumerate(L["hiring_subjects"], 1):
        para(doc, f"{i}. {s}", 10, indent=0.4, space_after=1)

    h2(doc, "5. 처리 절차")
    for i, s in enumerate([
        "채용 확정자 명단을 '01_채용시교육_면제판정표.xlsx'에 입력",
        "특별교육 담당 부서로부터 특별교육 이수자 명단·증빙 사본 수령 (E1)",
        "경력증명서로 동종 업종(세분류) 근무개월·이직일 확인 (E2)",
        "판정표의 판정 결과와 근거조항 확인 후 해당 인원 채용 시 교육 실시 또는 면제 처리",
        "입증서류를 교육 서류철에 함께 보관",
    ], 1):
        para(doc, f"{i}) {s}", 10, indent=0.4, space_after=1)
    para(doc, f"※ 미실시 시: {L['penalty']['text']} ({L['penalty']['basis']})", 10, True, color="C00000")
    stamp(doc, ctx)
    doc.save(ctx.path("01_채용시교육_면제감면_기준.docx"))


# ---------- 02 하반기 정기교육 법정기준 ----------
def build_regular_standard(ctx):
    L = ctx.laws
    doc = new_doc()
    title(doc, f"{ctx.half_label} 제설기간제근로자 정기교육 법정기준")
    h2(doc, "1. 교육 대상 및 시간")
    table(doc, [
        ["항목", "내용", "근거"],
        ["대상", f"{ctx.half_label}({ctx.period}) 재직 제설기간제근로자 전원\n({L['worker']['category_regular']})", "법 제29조제1항"],
        ["법정 시간", f"매반기 {ctx.legal_hours}시간 이상", L["hours"]["regular_basis"]],
        ["이번 편성", f"{ctx.total_hours}시간 (요구 {ctx.required_hours:g}시간"
                     + (", 전년도 무재해 50% 감면 적용)" if ctx.cfg.get("apply_no_accident_reduction") else ", 감면 미적용)"), "settings.json"],
        ["강사", f"{ctx.cfg['instructor_role']} (자격 근거: 제26조제3항제1호나목)", L["instructor"]["basis"]],
        ["교육형태", " / ".join(L["methods"]["allowed"]), L["methods"]["basis"]],
    ], widths=[2.5, 9, 5.5])

    h2(doc, "2. 법정 교과목 (모두 포함해야 함)")
    hb = ctx.hours_by_subject()
    rows = [["번호", "법정 교과목", "편성 시간"]]
    for no, s in ctx.subjects.items():
        rows.append([str(no), s["name"], f"{hb[no]}h"])
    table(doc, rows, widths=[1.3, 13, 2.5], size=10)
    para(doc, f"근거: {L['regular_subjects_basis']} ({L['sources']['annex5']['amended']} 개정)", 9, color="595959")
    para(doc, "※ " + L["methods"]["content_rule"], 9, color="595959")

    h2(doc, "3. 정기교육 감면·인정 사유 (해당 시 선택 적용)")
    rows = [["구분", "조건", "효과", "근거", "입증서류"]]
    for r in L["regular_reductions"]:
        rows.append([r["id"], r["condition"], r["effect"], r["basis"], r["evidence"]])
    table(doc, rows, widths=[1.2, 5, 4, 3.4, 3.4], size=9)
    para(doc, "※ " + L["instructor"]["instructor_credit_note"], 9, color="595959")

    h2(doc, "4. 원격교육 병행 시 준수사항")
    for r in L["methods"]["online_rules"]:
        para(doc, "• " + r, 10, indent=0.4, space_after=1)
    para(doc, f"※ 미실시 시: {L['penalty']['text']} ({L['penalty']['basis']})", 10, True, color="C00000")
    stamp(doc, ctx)
    doc.save(ctx.path("02_하반기_정기교육_법정기준.docx"))


# ---------- 05 교육계획보고 공문 ----------
def build_plan_gongmun(ctx):
    L = ctx.laws
    doc = new_doc()
    subj = f"{ctx.half_label} 제설기간제근로자 정기 안전보건교육 실시계획 보고"
    _gongmun_head(doc, ctx, subj)
    para(doc, f"1. 관련: 「산업안전보건법」 제29조제1항, 같은 법 시행규칙 제26조 및 별표4·별표5, 「안전보건교육규정」 제3조의2", 11)
    para(doc, f"2. 위 호와 관련하여 {ctx.half_label} 제설기간제근로자 정기 안전보건교육을 아래와 같이 실시하고자 보고합니다.", 11, space_after=8)
    para(doc, "가. 교육 개요", 11, True, indent=0.4)
    days = sorted({s.day for s in ctx.sessions})
    table(doc, [
        ["구분", "내용"],
        ["목적", "제설작업 중 산업재해 예방 및 법정 정기교육 이행"],
        ["대상", f"제설기간제근로자 {ph('', 2)}명 ({L['worker']['type']})"],
        ["일시", " / ".join(f"{d}일차 {ph('', 4)}" for d in days) + f"  (총 {ctx.total_hours}시간)"],
        ["장소", ph(ctx.P.get("place"), 6)],
        ["강사", f"{ctx.cfg['instructor_role']} {ph(ctx.P.get('instructor_name'), 3)} (시행규칙 제26조제3항제1호나목)"],
        ["교육방법", ", ".join(sorted({s.method for s in ctx.sessions}))],
        ["법정시간", f"매반기 {ctx.legal_hours}시간 이상 ({L['hours']['regular_basis']})"],
    ], widths=[3, 13.5], size=10)
    para(doc, "나. 교육과목 및 시간", 11, True, indent=0.4)
    rows = [["일차", "시간", "교육과목(별표5)", "시간", "방법"]]
    for s in ctx.sessions:
        rows.append([f"{s.day}일차", f"{s.start}~{s.end}", f"{s.subject_no}. {ctx.subjects[s.subject_no]['short']}", f"{s.hours}h", s.method])
    rows.append(["", "", "합계", f"{ctx.total_hours}h", ""])
    table(doc, rows, widths=[1.8, 3, 7.5, 1.6, 2.8], size=10)
    para(doc, f"다. 소요예산: {ph(ctx.P.get('budget'), 5)}", 11, indent=0.4)
    para(doc, "라. 채용 시 교육: 면제·감면 판정표에 따라 처리하고, 특별교육 이수자는 채용 시 교육 실시로 간주(시행규칙 제26조제1항 후단)", 11, indent=0.4, space_after=8)
    para(doc, "붙임  1. 교육시간표 및 강사자격 1부.", 11)
    para(doc, "        2. 과목별 교안 1부.", 11)
    para(doc, "        3. 채용 시 교육 면제 판정표 1부.", 11, space_after=0)
    _gongmun_tail(doc, ctx)
    doc.save(ctx.path("05_교육계획보고_공문.docx"))


# ---------- 06 교육일지 ----------
def build_journal(ctx):
    doc = new_doc()
    days = sorted({s.day for s in ctx.sessions})
    for idx, d in enumerate(days):
        if idx:
            page_break(doc)
        title(doc, f"안전보건교육 일지 ({d}일차)", 16)
        _approval_box(doc, ctx, height=0.9)
        ss = [s for s in ctx.sessions if s.day == d]
        hours = sum(s.hours for s in ss)
        table(doc, [
            ["교육구분", f"근로자 정기교육 ({ctx.half_label})", "교육일시", f"{ph('', 4)} {ss[0].start}~{ss[-1].end} ({hours}시간)"],
            ["교육장소", ph(ctx.P.get("place"), 5), "교육방법", ", ".join(sorted({s.method for s in ss}))],
            ["강사", f"{ph(ctx.P.get('instructor_name'), 3)}\n({ctx.cfg['instructor_role']})", "강사자격 근거", "시행규칙 제26조제3항제1호나목"],
            ["대상인원", ph("", 2) + "명", "참석 / 불참", f"{ph('', 2)}명 / {ph('', 2)}명"],
        ], widths=[2.6, 5.9, 2.8, 5.7], header=False, size=10, first_col_shade=True)
        rows = [["시간", "교육과목(별표5)", "주요 교육내용", "시간"]]
        for s in ss:
            cur = ctx.cur[str(s.subject_no)]
            rows.append([f"{s.start}~{s.end}", f"{s.subject_no}. {ctx.subjects[s.subject_no]['short']}",
                         ", ".join(cur["hazards"][:2]), f"{s.hours}h"])
        table(doc, rows, widths=[2.4, 4.6, 8.5, 1.5], size=8.5)
        table(doc, [
            ["불참자 및 보충교육 계획", ph("", 12)],
            ["교육 중 건의·특이사항", ph("", 12)],
            ["교육 효과 확인(퀴즈 등)", "□ 실시   □ 미실시     결과: " + ph("", 6)],
        ], widths=[4.5, 12.5], header=False, size=9.5, height=1.1, first_col_shade=True)
        table(doc, [
            ["확인", "강사", "관리감독자", "안전보건관리책임자"],
            ["성명/서명", "(인)", "(인)", "(인)"],
        ], widths=[3, 4.6, 4.7, 4.7], size=9.5, height=0.8)
        para(doc, "※ 현장교육 시 교육 사실 확인 일지 작성 의무: 안전보건교육규정 제4조제2호 / 출석부·사진대지와 함께 보관", 8, color="7F7F7F")
        stamp(doc, ctx)
    doc.save(ctx.path("06_교육일지.docx"))


# ---------- 06 교육사진대지 ----------
def build_photo_sheet(ctx):
    doc = new_doc()
    days = sorted({s.day for s in ctx.sessions})
    slots = ctx.cfg.get("photo_slots", 4)
    for idx, d in enumerate(days):
        if idx:
            page_break(doc)
        title(doc, f"안전보건교육 사진대지 ({d}일차)", 18)
        table(doc, [["교육명", f"{ctx.half_label} 제설기간제근로자 정기 안전보건교육", "교육일시", ph("", 5)],
                    ["교육장소", ph(ctx.P.get("place"), 5), "작성자", ph("", 3)]],
              widths=[2.5, 7, 2.5, 5], header=False, size=10, first_col_shade=True)
        nrow = (slots + 1) // 2
        t = doc.add_table(rows=nrow * 2, cols=2)
        t.style = "Table Grid"
        t.alignment = WD_TABLE_ALIGNMENT.CENTER
        n = 0
        for r in range(nrow):
            for c in range(2):
                n += 1
                pic = t.cell(r * 2, c)
                pic.width = Cm(8.5)
                pic.text = ""
                pp = pic.paragraphs[0]
                pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                _font(pp.add_run(f"사진 {n}\n(이곳에 사진 삽입)"), 10, color="A6A6A6")
                cap = t.cell(r * 2 + 1, c)
                cap.text = ""
                _font(cap.paragraphs[0].add_run(f"일시: {ph('', 4)}\n장소: {ph('', 4)}\n내용: {ph('', 6)}"), 9)
            t.rows[r * 2].height = Cm(6.5)
        para(doc, "")
        para(doc, "※ 강사·교육생·교육자료(화면)가 함께 보이도록 촬영, 개인정보(얼굴 등) 활용 동의 확인", 8, color="7F7F7F")
        stamp(doc, ctx)
    doc.save(ctx.path("06_교육사진대지.docx"))


# ---------- 07 교육결과보고 공문 ----------
def build_result_gongmun(ctx):
    doc = new_doc()
    subj = f"{ctx.half_label} 제설기간제근로자 정기 안전보건교육 실시결과 보고"
    _gongmun_head(doc, ctx, subj)
    para(doc, f"1. 관련: {ph(ctx.P.get('dept'), 4)}-{ph('', 3)}({ph('', 5)}) \"{ctx.half_label} 제설기간제근로자 정기 안전보건교육 실시계획 보고\"", 11)
    para(doc, "2. 위 호와 관련하여 정기 안전보건교육 실시결과를 아래와 같이 보고합니다.", 11, space_after=8)
    para(doc, "가. 실시 결과", 11, True, indent=0.4)
    table(doc, [
        ["구분", "내용"],
        ["교육일시", ph("", 8) + f" (총 {ctx.total_hours}시간)"],
        ["교육장소", ph(ctx.P.get("place"), 6)],
        ["강사", f"{ctx.cfg['instructor_role']} {ph(ctx.P.get('instructor_name'), 3)}"],
        ["대상 / 이수 / 미이수", f"{ph('', 2)}명 / {ph('', 2)}명 / {ph('', 2)}명 (이수율 {ph('', 2)}%)"],
        ["법정시간 충족", f"□ 충족(1인당 {ctx.required_hours:g}시간 이상)   □ 미충족 → 보충교육"],
    ], widths=[4, 12.5], size=10)
    para(doc, "나. 과목별 실시 시간", 11, True, indent=0.4)
    hb = ctx.hours_by_subject()
    rows = [["번호", "교육과목(별표5)", "계획", "실시"]]
    for no, s in ctx.subjects.items():
        rows.append([str(no), s["short"], f"{hb[no]}h", ph("", 1) + "h"])
    rows.append(["", "합계", f"{ctx.total_hours}h", ph("", 1) + "h"])
    table(doc, rows, widths=[1.5, 9.5, 2.5, 2.5], size=10)
    para(doc, f"다. 미이수자 보충교육 계획: {ph('', 10)} (해당 반기 내 실시)", 11, indent=0.4)
    para(doc, f"라. 채용 시 교육: 실시 {ph('', 2)}명 / 특별교육 이수로 간주 {ph('', 2)}명 / 감면(50%) {ph('', 2)}명 / 면제 {ph('', 2)}명", 11, indent=0.4, space_after=8)
    for i, s in enumerate(["교육일지 1부.", "출석부 1부.", "교육사진대지 1부.", "법령준수 체크리스트 1부.",
                           "교육결과 집계표 1부.", "채용 시 교육 면제 판정표 및 특별교육 이수 증빙 사본 각 1부."], 1):
        para(doc, ("붙임  " if i == 1 else "        ") + f"{i}. {s}", 11, space_after=0)
    _gongmun_tail(doc, ctx)
    doc.save(ctx.path("07_교육결과보고_공문.docx"))


# ---------- 00 README ----------
def build_readme(ctx, law_check: dict, files: list):
    doc = new_doc()
    title(doc, f"{ctx.year}년 제설기간제근로자 안전보건교육 패키지 안내")
    h2(doc, "1. 적용 법령 기준")
    s = ctx.laws["sources"]
    table(doc, [["법령", "시행·개정", "번호"]] + [
        [v["name"], v.get("effective", v.get("amended", "")), v.get("number", "")] for v in s.values()
    ], widths=[8, 4, 5], size=10)
    h2(doc, "2. 법령 개정 확인 결과")
    para(doc, law_check.get("status_text", "확인 결과 없음"), 11, True,
         color="C00000" if law_check.get("changed") or law_check.get("failed") else "2E7D32")
    for line in law_check.get("details", []):
        para(doc, "• " + line, 10, indent=0.4, space_after=1)
    h2(doc, "3. 담당자 할 일")
    for i, t in enumerate([
        "모든 문서의 〔　　〕 공란(기관명·수신·결재라인·일자·장소·강사 성명 등) 기입",
        "채용 확정 후 '01_채용시교육_면제판정표.xlsx'로 채용 시 교육 면제·감면 판정",
        "'05_교육계획보고_공문.docx' 상신 (교육 실시 전)",
        f"{ctx.half_label} 내(12월 31일까지) 정기교육 {ctx.required_hours:g}시간 이상 실시",
        "교육 당일 출석부 서명, 교육일지·사진대지 작성, 체크리스트 점검",
        "'07_교육결과보고_공문.docx' 상신 및 서류 일체 보관(3년 이상 권장)",
    ], 1):
        para(doc, f"{i}. {t}", 11, indent=0.4, space_after=2)
    h2(doc, "4. 산출물 목록")
    for f in files:
        para(doc, "• " + f, 10, indent=0.4, space_after=0)
    para(doc, "")
    stamp(doc, ctx)
    doc.save(ctx.path("00_README_패키지안내.docx"))


def build_all(ctx):
    build_hiring_exemption(ctx)
    build_regular_standard(ctx)
    build_plan_gongmun(ctx)
    build_journal(ctx)
    build_photo_sheet(ctx)
    build_result_gongmun(ctx)
