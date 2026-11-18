"""Excel 산출물: 면제판정표, 강사자격·시간표, 출석부, 결과 집계표, 법령준수 체크리스트."""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from common import FONT, BLANK, ph

THIN = Side(style="thin", color="808080")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="D9E2F3")
SUB_FILL = PatternFill("solid", fgColor="F2F2F2")
INPUT_FILL = PatternFill("solid", fgColor="FFF9E6")
WRAP = Alignment(wrap_text=True, vertical="center")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def f(size=10, bold=False, color=None):
    return Font(name=FONT, size=size, bold=bold, color=color)


def title(ws, text, ncol, row=1, sub=None):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row, 1, text)
    c.font = f(15, True, "1F3864")
    c.alignment = CENTER
    ws.row_dimensions[row].height = 30
    if sub:
        ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=ncol)
        c = ws.cell(row + 1, 1, sub)
        c.font = f(9, color="595959")
        c.alignment = CENTER


def header(ws, row, values, widths=None):
    for j, v in enumerate(values, 1):
        c = ws.cell(row, j, v)
        c.font = f(10, True)
        c.fill = HEAD_FILL
        c.alignment = CENTER
        c.border = BORDER
        if widths:
            ws.column_dimensions[get_column_letter(j)].width = widths[j - 1]
    ws.row_dimensions[row].height = 32


def cell(ws, r, c, v, bold=False, fill=None, align=WRAP, size=10, color=None):
    x = ws.cell(r, c, v)
    x.font = f(size, bold, color)
    x.alignment = align
    x.border = BORDER
    if fill:
        x.fill = fill
    return x


def footer(ws, row, ctx, ncol):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)
    c = ws.cell(row, 1, ctx.law_stamp)
    c.font = f(8, color="7F7F7F")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[row].height = 30


def page_setup(ws, landscape=False):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True


# ---------- 01 채용 시 교육 면제 판정표 ----------
HIRING_COLS = ["연번", "성명", "계약 시작일", "계약 종료일", "계약기간\n(개월)", "특별교육\n이수(Y/N)",
               "동종업종(세분류)\n근무개월 합계", "직전 이직일", "채용일", "동일 도급인·\n동일업무(Y/N)",
               "기타 장관인정\n교육(Y/N)", "판정", "실시할 채용 시\n교육시간(h)", "근거조항", "입증서류 확인"]
HIRING_W = [5, 10, 12, 12, 9, 10, 14, 12, 12, 13, 12, 20, 13, 30, 14]


def _hiring_formulas(r, full_h):
    half_h = full_h / 2
    judge = (f'=IF(B{r}="","",IF(UPPER(F{r})="Y","E1 특별교육 이수 간주",IF(UPPER(J{r})="Y","E3 면제",'
             f'IF(UPPER(K{r})="Y","E5 면제(인정근거 확인)",IF(AND(N(G{r})>=6,H{r}<>"",I{r}<>"",I{r}<=EDATE(H{r},12)),'
             f'"E2 50% 감면","해당없음(전부 실시)")))))')
    months = f'=IF(OR(C{r}="",D{r}=""),"",ROUND((D{r}-C{r}+1)/30.4,1))'
    hours = (f'=IF(L{r}="","",IF(LEFT(L{r},2)="E2",{half_h:g},IF(L{r}="해당없음(전부 실시)",{full_h:g},0)))')
    basis = (f'=IF(L{r}="","",IF(LEFT(L{r},2)="E1","시행규칙 제26조제1항 후단",IF(LEFT(L{r},2)="E2","법 제30조제2항, 시행규칙 제27조제5항제1호",'
             f'IF(LEFT(L{r},2)="E3","법 제30조제2항, 시행규칙 제27조제5항제3호",IF(LEFT(L{r},2)="E5","시행규칙 제27조제5항제4호",'
             f'"법 제29조제2항, 시행규칙 별표4 제1호나목3)")))))')
    return months, judge, hours, basis


def _hiring_sheet(ws, ctx, start, rows_data, nrows):
    header(ws, start, HIRING_COLS, HIRING_W)
    yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(yn)
    full_h = ctx.laws["hours"]["hiring"]
    for i in range(nrows):
        r = start + 1 + i
        data = rows_data[i] if i < len(rows_data) else {}
        months, judge, hours, basis = _hiring_formulas(r, full_h)
        vals = [i + 1, data.get("name"), data.get("start"), data.get("end"), months, data.get("special"),
                data.get("months"), data.get("left"), data.get("hired"), data.get("same"), data.get("other"),
                judge, hours, basis, None]
        for j, v in enumerate(vals, 1):
            is_input = j in (2, 3, 4, 6, 7, 8, 9, 10, 11, 15)
            c = cell(ws, r, j, v, fill=INPUT_FILL if is_input else None, align=CENTER if j != 14 else WRAP, size=9)
            if j in (3, 4, 8, 9):
                c.number_format = "yyyy-mm-dd"
        for col in ("F", "J", "K"):
            yn.add(f"{col}{r}")
    ws.freeze_panes = ws.cell(start + 1, 3)


def build_hiring_judgement(ctx):
    from datetime import date
    wb = Workbook()
    ws = wb.active
    ws.title = "판정표"
    n = ctx.cfg.get("attendance_rows", 40)
    title(ws, f"{ctx.year}년 제설기간제근로자 채용 시 교육 면제·감면 판정표", len(HIRING_COLS),
          sub="노란 칸만 입력 → 판정·시간·근거 자동 표시 / 판정 결과는 검토용이며 최종 판단은 담당자가 입증서류로 확인")
    _hiring_sheet(ws, ctx, 4, [], n)
    footer(ws, 4 + n + 2, ctx, len(HIRING_COLS))
    page_setup(ws, landscape=True)

    # 예시(테스트) 시트: 조건별 기대값 검증용
    y = ctx.year
    ex = wb.create_sheet("예시(검증용)")
    title(ex, "판정 예시 - 조건별 기대 결과", len(HIRING_COLS) + 1)
    samples = [
        dict(name="예시1", start=date(y, 11, 15), end=date(y + 1, 3, 15), special="Y", months=0, expect="E1 특별교육 이수 간주"),
        dict(name="예시2", start=date(y, 11, 15), end=date(y + 1, 3, 15), special="N", months=8, left=date(y, 3, 31), hired=date(y, 11, 15), expect="E2 50% 감면"),
        dict(name="예시3", start=date(y, 11, 15), end=date(y + 1, 3, 15), special="N", months=4, left=date(y, 3, 31), hired=date(y, 11, 15), expect="해당없음(전부 실시)"),
        dict(name="예시4", start=date(y, 11, 15), end=date(y + 1, 3, 15), special="N", months=9, left=date(y - 1, 3, 31), hired=date(y, 11, 15), expect="해당없음(전부 실시)"),
        dict(name="예시5", start=date(y, 11, 15), end=date(y + 1, 3, 15), special="N", same="Y", expect="E3 면제"),
    ]
    _hiring_sheet(ex, ctx, 3, samples, len(samples))
    c = len(HIRING_COLS) + 1
    cell(ex, 3, c, "기대 결과", True, HEAD_FILL, CENTER)
    ex.column_dimensions[get_column_letter(c)].width = 22
    for i, s in enumerate(samples):
        cell(ex, 4 + i, c, s["expect"], align=CENTER, size=9)
    notes = ["예시3: 동종업종 4개월(전년도 제설 1회) → 6개월 미달로 감면 불가",
             "예시4: 근무 9개월이나 이직 후 1년 초과 → 감면 불가"]
    for k, t in enumerate(notes):
        ex.cell(10 + k, 1, "※ " + t).font = f(9, color="C00000")

    # 기준 시트
    st = wb.create_sheet("면제기준")
    title(st, "채용 시 교육 면제·감면 기준", 5)
    header(st, 3, ["구분", "조건", "효과", "근거", "입증서류"], [7, 45, 35, 32, 32])
    for i, e in enumerate(ctx.laws["hiring_exemptions"]):
        for j, k in enumerate(["id", "condition", "effect", "basis", "evidence"], 1):
            cell(st, 4 + i, j, e[k], size=9)
        st.row_dimensions[4 + i].height = 45
    r = 4 + len(ctx.laws["hiring_exemptions"]) + 1
    for t in ctx.laws["hiring_exemption_notes"]:
        st.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        x = st.cell(r, 1, "※ " + t)
        x.font = f(9, color="C00000")
        x.alignment = WRAP
        st.row_dimensions[r].height = 28
        r += 1
    footer(st, r + 1, ctx, 5)
    wb.save(ctx.path("01_채용시교육_면제판정표.xlsx"))
    return samples


# ---------- 04 강사자격 및 교육시간표 ----------
def build_timetable(ctx):
    wb = Workbook()
    ws = wb.active
    ws.title = "교육시간표"
    cols = ["일차", "일자", "시간", "교육과목(별표5)", "시간(h)", "교육방법", "강사", "강사자격 근거", "비고"]
    title(ws, f"{ctx.half_label} 제설기간제근로자 정기 안전보건교육 시간표", len(cols),
          sub=f"법정 매반기 {ctx.legal_hours}시간 이상({ctx.laws['hours']['regular_basis']}) / 교육과목: {ctx.laws['regular_subjects_basis']}")
    header(ws, 4, cols, [6, 12, 13, 52, 8, 13, 22, 24, 18])
    r = 5
    for s in ctx.sessions:
        basis = "시행규칙 제26조제3항제1호나목" if s.method == "집체교육" else "안전보건교육규정 제3조의2제2항제5호 단서"
        vals = [f"{s.day}일차", BLANK, f"{s.start}~{s.end}", f"{s.subject_no}. {s.subject_name}", s.hours, s.method,
                s.instructor, basis, s.note]
        for j, v in enumerate(vals, 1):
            cell(ws, r, j, v, align=CENTER if j in (1, 2, 3, 5, 6) else WRAP, size=9,
                 fill=INPUT_FILL if j == 2 else None)
        ws.row_dimensions[r].height = 36
        r += 1
    cell(ws, r, 4, "합계", True, SUB_FILL, CENTER)
    cell(ws, r, 5, f"=SUM(E5:E{r - 1})", True, SUB_FILL, CENTER)
    cell(ws, r, 6, f'=IF(E{r}>={ctx.required_hours:g},"법정시간 충족","법정시간 미달")', True, SUB_FILL, CENTER)
    last = r - 1
    r += 2
    ws.cell(r, 1, f"※ 강사: {ctx.cfg['instructor_role']} 성명 {ph(ctx.P.get('instructor_name'), 3)} / 장소 {ph(ctx.P.get('place'), 5)}").font = f(10)
    footer(ws, r + 2, ctx, len(cols))
    page_setup(ws, landscape=True)

    ws2 = wb.create_sheet("과목별 시간 합계")
    title(ws2, "법정 교과목별 편성 시간 확인", 4)
    header(ws2, 3, ["번호", "법정 교과목(별표5)", "편성 시간(h)", "포함 여부"], [7, 70, 13, 13])
    for i, (no, s) in enumerate(ctx.subjects.items()):
        rr = 4 + i
        cell(ws2, rr, 1, no, align=CENTER)
        cell(ws2, rr, 2, s["name"], size=9)
        cell(ws2, rr, 3, f'=SUMIF(교육시간표!$D$5:$D${last},"{no}. *",교육시간표!$E$5:$E${last})', align=CENTER)
        cell(ws2, rr, 4, f'=IF(C{rr}>=1,"포함","누락")', align=CENTER)
        ws2.row_dimensions[rr].height = 30
    rr = 4 + len(ctx.subjects)
    cell(ws2, rr, 2, "합계", True, SUB_FILL, CENTER)
    cell(ws2, rr, 3, f"=SUM(C4:C{rr - 1})", True, SUB_FILL, CENTER)
    footer(ws2, rr + 2, ctx, 4)

    ws3 = wb.create_sheet("강사자격")
    title(ws3, "자체 실시 안전보건교육 강사 자격", 4, sub=ctx.laws["instructor"]["basis"])
    header(ws3, 4, ["호·목", "자격 요건", "본 교육 지정", "비고"], [9, 75, 13, 25])
    sel = ctx.laws["instructor"]["selected_code"]
    for i, q in enumerate(ctx.laws["instructor"]["qualifications"]):
        rr = 5 + i
        cell(ws3, rr, 1, q["code"], align=CENTER)
        cell(ws3, rr, 2, q["text"], size=9)
        cell(ws3, rr, 3, "● 지정" if q["code"] == sel else "", align=CENTER, bold=True)
        cell(ws3, rr, 4, "성명: " + BLANK if q["code"] == sel else "", size=9)
        ws3.row_dimensions[rr].height = 30
    rr = 5 + len(ctx.laws["instructor"]["qualifications"]) + 1
    for t in [ctx.laws["instructor"]["instructor_credit_note"]] + ctx.laws["methods"]["online_rules"]:
        ws3.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
        x = ws3.cell(rr, 1, "※ " + t)
        x.font = f(9)
        x.alignment = WRAP
        ws3.row_dimensions[rr].height = 30
        rr += 1
    footer(ws3, rr + 1, ctx, 4)
    wb.save(ctx.path("04_강사자격_및_교육시간표.xlsx"))


# ---------- 06 출석부 ----------
def build_attendance(ctx):
    wb = Workbook()
    wb.remove(wb.active)
    n = ctx.cfg.get("attendance_rows", 40)
    for d in sorted({s.day for s in ctx.sessions}):
        ss = [s for s in ctx.sessions if s.day == d]
        ws = wb.create_sheet(f"{d}일차")
        cols = ["연번", "소속", "성명", "생년월일", "입실 시각", "퇴실 시각", "이수 시간(h)", "서명", "비고"]
        title(ws, f"안전보건교육 출석부 ({d}일차)", len(cols))
        info = [("교육명", f"{ctx.half_label} 제설기간제근로자 정기 안전보건교육"),
                ("교육일시", f"{BLANK}  {ss[0].start}~{ss[-1].end} ({sum(s.hours for s in ss)}시간)"),
                ("교육장소", ph(ctx.P.get("place"), 5)),
                ("교육과목", " / ".join(f"{s.subject_no}. {ctx.subjects[s.subject_no]['short']}" for s in ss)),
                ("강사", f"{ctx.cfg['instructor_role']} {ph(ctx.P.get('instructor_name'), 3)} (서명)")]
        for i, (k, v) in enumerate(info):
            r = 3 + i
            cell(ws, r, 1, k, True, SUB_FILL, CENTER)
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
            ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=len(cols))
            cell(ws, r, 3, v, size=9)
            ws.row_dimensions[r].height = 22 if k != "교육과목" else 34
        hr = 3 + len(info) + 1
        header(ws, hr, cols, [6, 14, 12, 12, 10, 10, 11, 16, 14])
        for i in range(n):
            r = hr + 1 + i
            for j in range(1, len(cols) + 1):
                cell(ws, r, j, i + 1 if j == 1 else None, align=CENTER)
            ws.row_dimensions[r].height = 22
        last = hr + n
        cell(ws, last + 1, 6, "참석 인원", True, SUB_FILL, CENTER)
        cell(ws, last + 1, 7, f'=COUNTA(C{hr + 1}:C{last})', True, SUB_FILL, CENTER)
        ws.cell(last + 3, 1, "※ 입·퇴실 시각과 서명은 본인이 직접 기재. 지각·조퇴 시 실제 이수시간을 기재하고 부족분은 보충교육.").font = f(9)
        footer(ws, last + 4, ctx, len(cols))
        ws.print_title_rows = f"{hr}:{hr}"
        page_setup(ws)
    wb.save(ctx.path("06_출석부.xlsx"))


# ---------- 07 교육결과 집계표 ----------
def build_result_summary(ctx):
    wb = Workbook()
    ws = wb.active
    ws.title = "이수 집계"
    days = sorted({s.day for s in ctx.sessions})
    cols = ["연번", "소속", "성명"] + [f"{d}일차(h)" for d in days] + ["원격·보충(h)", "합계(h)", "판정", "비고"]
    title(ws, f"{ctx.half_label} 정기 안전보건교육 이수 집계표", len(cols),
          sub=f"1인당 요구시간 {ctx.required_hours:g}시간 / 출석부의 이수시간을 옮겨 적으면 판정 자동 표시")
    n = ctx.cfg.get("attendance_rows", 40)
    hr = 5
    header(ws, hr, cols, [6, 14, 12] + [11] * len(days) + [13, 10, 14, 16])
    first_h = 4
    last_h = 3 + len(days) + 1
    sum_col = last_h + 1
    for i in range(n):
        r = hr + 1 + i
        cell(ws, r, 1, i + 1, align=CENTER)
        for j in range(2, len(cols) + 1):
            cell(ws, r, j, None, align=CENTER, fill=INPUT_FILL if j < sum_col else None)
        a, b = get_column_letter(first_h), get_column_letter(last_h)
        cell(ws, r, sum_col, f'=IF(C{r}="","",SUM({a}{r}:{b}{r}))', align=CENTER)
        s = get_column_letter(sum_col)
        cell(ws, r, sum_col + 1, f'=IF(C{r}="","",IF({s}{r}>={ctx.required_hours:g},"이수","미달(보충 필요)"))', align=CENTER)
    last = hr + n
    j = get_column_letter(sum_col + 1)
    # 상단 요약: 대상 | 이수 | 미이수 | 이수율 (3행 가로 배치)
    summary = [("대상", f'=COUNTA(C{hr + 1}:C{last})'),
               ("이수", f'=COUNTIF({j}{hr + 1}:{j}{last},"이수")'),
               ("미이수", f'=COUNTIF({j}{hr + 1}:{j}{last},"미달*")'),
               ("이수율", '=IF(B3=0,"",D3/B3)')]
    for k, (lab, fx) in enumerate(summary):
        cell(ws, 3, 1 + k * 2, lab, True, SUB_FILL, CENTER)
        x = cell(ws, 3, 2 + k * 2, fx, align=CENTER)
    x.number_format = "0.0%"
    footer(ws, last + 2, ctx, len(cols))
    ws.freeze_panes = ws.cell(hr + 1, 4)
    page_setup(ws)
    wb.save(ctx.path("07_교육결과_집계표.xlsx"))


# ---------- 08 법령준수 체크리스트 ----------
def regular_check_items(ctx):
    L = ctx.laws
    items = [
        ("대상", f"{ctx.half_label} 재직 제설기간제근로자 전원을 교육 대상에 포함했는가", "법 제29조제1항"),
        ("시간", f"1인당 교육시간이 {ctx.required_hours:g}시간 이상인가 (법정 매반기 {ctx.legal_hours}시간, 감면 시 감면 후 시간)", L["hours"]["regular_basis"]),
        ("과목", "별표5 정기교육 8개 과목을 모두 포함했는가", L["regular_subjects_basis"]),
        ("과목", "1번 과목에 화재·폭발 사고 시 대피, 2번 과목에 폭염·한파작업 건강장해 응급조치를 포함했는가", f"시행규칙 별표5({L['sources']['annex5']['amended']} 개정)"),
        ("내용", "교육내용을 제설작업 환경·작업내용 등 사업장 위험성에 맞춰 정했는가", "안전보건교육규정 제3조의2제2항제1호"),
        ("강사", f"강사가 자격 요건을 충족하는가 ({ctx.cfg['instructor_role']})", "시행규칙 제26조제3항제1호나목, 안전보건교육규정 제3조의2제2항제5호"),
        ("교재", "교육내용에 적합한 교재(교안)를 사용했는가", "시행규칙 제36조제1항, 안전보건교육규정 제3조의2제2항제4호"),
        ("형태", "교육형태가 집체·현장·인터넷 원격·비대면 실시간 중 하나 또는 혼합인가", L["methods"]["basis"]),
        ("현장", "현장교육 시 자격 있는 강사가 주관하고 교육일지를 작성했는가", "안전보건교육규정 제4조"),
        ("원격", "원격교육 시 과목당 1시간 이상, 강의 동영상 50% 이상인가", "안전보건교육규정 제5조제1호"),
        ("원격", "모바일 원격교육 시 작업·운전 중 수강 금지를 공지하고 확인받았는가 (제설장비 운전원 주의)", "안전보건교육규정 제5조제3호"),
        ("위탁", "위탁 시 등록된 근로자 안전보건교육기관인가", L["outsourcing_basis"]),
        ("감면", "감면 적용 시 전년도 산업재해 미발생 입증자료를 갖췄는가", "법 제30조제1항제1호, 시행규칙 제27조제1항"),
        ("감면", "타 법령 교육·체험교육 등 인정 시 이수증·명단·일지를 갖췄는가", "시행규칙 제27조제3항, 안전보건교육규정 제9조"),
        ("기록", "출석부(서명), 교육일지, 사진대지를 작성·보관했는가", "안전보건교육규정 제4조제2호, 제9조제5항(입증서류)"),
        ("보충", "미이수자 보충교육을 해당 반기(12월 31일) 내 실시했는가", "법 제29조제1항"),
        ("인지", f"미실시 시 {L['penalty']['text']}임을 인지했는가", L["penalty"]["basis"]),
    ]
    return items


def build_checklist(ctx):
    wb = Workbook()
    ws = wb.active
    ws.title = "정기교육(회차별)"
    days = sorted({s.day for s in ctx.sessions})
    cols = ["No", "구분", "점검 항목", "근거 조항"] + [f"{d}일차" for d in days] + ["반기 종합", "확인 내용·비고"]
    title(ws, f"{ctx.half_label} 정기 안전보건교육 법령준수 체크리스트", len(cols),
          sub="각 교육 회차 종료 후 판정(적합/부적합/해당없음) 선택 → 부적합 항목은 조치 후 비고에 기재")
    header(ws, 4, cols, [5, 7, 52, 36] + [10] * len(days) + [11, 24])
    dv = DataValidation(type="list", formula1='"적합,부적합,해당없음"', allow_blank=True)
    ws.add_data_validation(dv)
    items = regular_check_items(ctx)
    for i, (cat, item, basis) in enumerate(items):
        r = 5 + i
        cell(ws, r, 1, i + 1, align=CENTER)
        cell(ws, r, 2, cat, align=CENTER)
        cell(ws, r, 3, item, size=9)
        cell(ws, r, 4, basis, size=9)
        for k in range(len(days) + 1):
            c = cell(ws, r, 5 + k, None, align=CENTER, fill=INPUT_FILL)
            dv.add(c.coordinate)
        cell(ws, r, 6 + len(days), None, size=9)
        ws.row_dimensions[r].height = 34
    r = 5 + len(items)
    cell(ws, r, 3, "부적합 건수", True, SUB_FILL, CENTER)
    for k in range(len(days) + 1):
        col = get_column_letter(5 + k)
        cell(ws, r, 5 + k, f'=COUNTIF({col}5:{col}{r - 1},"부적합")', True, SUB_FILL, CENTER)
    r += 2
    for lab in ["점검자", "확인자(관리감독자)", "안전보건관리책임자"]:
        cell(ws, r, 2, lab, True, SUB_FILL, CENTER)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        cell(ws, r, 4, "성명:                    (서명)")
        r += 1
    footer(ws, r + 1, ctx, len(cols))
    ws.freeze_panes = "E5"
    page_setup(ws, landscape=True)

    ws2 = wb.create_sheet("채용시교육 면제 확인")
    cols2 = ["구분", "확인 항목", "효과", "근거 조항", "필요 입증서류", "확인 결과", "비고"]
    title(ws2, "채용 시 교육 면제·감면 법령준수 확인", len(cols2), sub="판정표(01)의 면제·감면 인원별로 입증서류 확보 여부 확인")
    header(ws2, 4, cols2, [7, 42, 30, 30, 30, 11, 20])
    dv2 = DataValidation(type="list", formula1='"적합,부적합,해당없음"', allow_blank=True)
    ws2.add_data_validation(dv2)
    for i, e in enumerate(ctx.laws["hiring_exemptions"]):
        r = 5 + i
        for j, k in enumerate(["id", "condition", "effect", "basis", "evidence"], 1):
            cell(ws2, r, j, e[k], size=9, align=CENTER if j == 1 else WRAP)
        c = cell(ws2, r, 6, None, align=CENTER, fill=INPUT_FILL)
        dv2.add(c.coordinate)
        cell(ws2, r, 7, None)
        ws2.row_dimensions[r].height = 44
    r = 5 + len(ctx.laws["hiring_exemptions"])
    extra = [("공통", "면제·감면 미해당자에게 채용 시 교육 8시간 이상을 실시했는가", "", ctx.laws["hours"]["hiring_basis"], "출석부·교육일지"),
             ("공통", "E2 판정 시 동종업종 6개월 이상과 이직 후 1년 이내를 모두 확인했는가 (전년도 제설 1회만으로는 미충족)", "", "시행규칙 제27조제5항제1호", "경력증명서·이직 확인서류")]
    for cat, item, eff, basis, ev in extra:
        for j, v in enumerate([cat, item, eff, basis, ev], 1):
            cell(ws2, r, j, v, size=9, align=CENTER if j == 1 else WRAP)
        c = cell(ws2, r, 6, None, align=CENTER, fill=INPUT_FILL)
        dv2.add(c.coordinate)
        cell(ws2, r, 7, None)
        ws2.row_dimensions[r].height = 40
        r += 1
    footer(ws2, r + 1, ctx, len(cols2))
    page_setup(ws2, landscape=True)
    wb.save(ctx.path("08_법령준수_체크리스트.xlsx"))
    return items


def build_all(ctx):
    samples = build_hiring_judgement(ctx)
    build_timetable(ctx)
    build_attendance(ctx)
    build_result_summary(ctx)
    build_checklist(ctx)
    return samples
