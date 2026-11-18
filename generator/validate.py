"""산출물 자체검증: 법정 시간·과목·근거 표기·파일 무결성·면제 판정 로직."""
from datetime import date, timedelta

from docx import Document
from openpyxl import load_workbook
from pptx import Presentation

EXPECTED_STATIC = [
    "00_README_패키지안내.docx",
    "01_채용시교육_면제감면_기준.docx",
    "01_채용시교육_면제판정표.xlsx",
    "02_하반기_정기교육_법정기준.docx",
    "04_강사자격_및_교육시간표.xlsx",
    "05_교육계획보고_공문.docx",
    "06_교육일지.docx",
    "06_출석부.xlsx",
    "06_교육사진대지.docx",
    "07_교육결과보고_공문.docx",
    "07_교육결과_집계표.xlsx",
    "08_법령준수_체크리스트.xlsx",
]


def judge_py(s, full_h):
    """엑셀 판정 수식과 같은 규칙의 파이썬 구현(수식 회귀검사용)."""
    if (s.get("special") or "").upper() == "Y":
        return "E1 특별교육 이수 간주", 0
    if (s.get("same") or "").upper() == "Y":
        return "E3 면제", 0
    if (s.get("other") or "").upper() == "Y":
        return "E5 면제(인정근거 확인)", 0
    left, hired = s.get("left"), s.get("hired")
    if (s.get("months") or 0) >= 6 and left and hired:
        try:
            limit = left.replace(year=left.year + 1)
        except ValueError:  # 2/29
            limit = left + timedelta(days=365)
        if hired <= limit:
            return "E2 50% 감면", full_h / 2
    return "해당없음(전부 실시)", full_h


def _docx_text(p):
    d = Document(p)
    parts = [x.text for x in d.paragraphs]
    for t in d.tables:
        for row in t.rows:
            parts.extend(c.text for c in row.cells)
    return "\n".join(parts)


def run(ctx, samples):
    errors, notes = [], []
    L = ctx.laws

    # 1) 법정 시간·과목
    if len(ctx.subjects) != 8:
        errors.append(f"정기교육 법정 과목 수가 8이 아님: {len(ctx.subjects)}")
    missing_cur = [no for no in ctx.subjects if str(no) not in ctx.cur]
    if missing_cur:
        errors.append(f"교안 원고 누락 과목: {missing_cur}")
    hb = ctx.hours_by_subject()
    zero = [no for no, h in hb.items() if h < 1]
    if zero:
        errors.append(f"시간표에 편성되지 않은 법정 과목: {zero}")
    if ctx.total_hours < ctx.required_hours:
        errors.append(f"총 교육시간 {ctx.total_hours}h < 요구 {ctx.required_hours}h")
    online = set(ctx.cfg.get("online_subjects", []))
    for s in ctx.sessions:
        if s.subject_no in online and s.hours < 1:
            errors.append(f"원격 과목 {s.subject_no} 1시간 미만(안전보건교육규정 제5조제1호)")
    notes.append(f"교육시간 {ctx.total_hours}h / 요구 {ctx.required_hours:g}h, 8과목 모두 편성")

    # 2) 파일 존재·열람
    expected = list(EXPECTED_STATIC) + [f"03_교안_과목{no}_" for no in ctx.subjects]
    present = [p.name for p in ctx.out.iterdir()]
    for e in expected:
        if not any(n.startswith(e) for n in present):
            errors.append(f"산출물 누락: {e}")
    stamp_key = L["sources"]["rule"]["number"]
    for p in ctx.out.iterdir():
        try:
            if p.suffix == ".docx":
                txt = _docx_text(p)
                if stamp_key not in txt:
                    errors.append(f"{p.name}: 적용 법령 기준 표기 누락")
            elif p.suffix == ".xlsx":
                wb = load_workbook(p)
                if not any(stamp_key in str(c.value or "") for ws in wb for row in ws.iter_rows() for c in row):
                    errors.append(f"{p.name}: 적용 법령 기준 표기 누락")
            elif p.suffix == ".pptx":
                prs = Presentation(p)
                if len(prs.slides) < 8:
                    errors.append(f"{p.name}: 슬라이드 {len(prs.slides)}장(<8)")
                cover = " ".join(sh.text_frame.text for sh in prs.slides[0].shapes if sh.has_text_frame)
                if "별표5" not in cover:
                    errors.append(f"{p.name}: 표지에 근거조항 누락")
        except Exception as ex:  # noqa: BLE001
            errors.append(f"{p.name}: 열기 실패 {ex}")

    # 3) 근거 조항 표기(핵심 문서)
    must = {
        "05_교육계획보고_공문.docx": ["제29조제1항", "제26조", "별표4", "별표5"],
        "02_하반기_정기교육_법정기준.docx": ["별표4", "별표5", "제26조제3항", "제175조제5항제1호"],
        "01_채용시교육_면제감면_기준.docx": ["제27조제5항제1호", "제26조제1항 후단", "제27조제5항제3호"],
    }
    for name, keys in must.items():
        p = ctx.out / name
        if p.exists():
            txt = _docx_text(p)
            for k in keys:
                if k not in txt:
                    errors.append(f"{name}: 근거 '{k}' 누락")
    wb = load_workbook(ctx.out / "08_법령준수_체크리스트.xlsx")
    ws = wb["정기교육(회차별)"]
    blank_basis = [r[0].value for r in ws.iter_rows(min_row=5) if r[2].value and r[0].value and isinstance(r[0].value, int) and not r[3].value]
    if blank_basis:
        errors.append(f"체크리스트 근거 미표기 항목: {blank_basis}")

    # 4) 면제 판정 로직 회귀검사
    full_h = L["hours"]["hiring"]
    wb = load_workbook(ctx.out / "01_채용시교육_면제판정표.xlsx")
    ex = wb["예시(검증용)"]
    expect_col = ex.max_column
    for i, s in enumerate(samples):
        got, _ = judge_py(s, full_h)
        exp = ex.cell(4 + i, expect_col).value
        if got != exp:
            errors.append(f"면제 판정 예시{i + 1}: 기대 '{exp}' / 로직 '{got}'")
        f = ex.cell(4 + i, 12).value
        if not (isinstance(f, str) and f.startswith("=IF(")):
            errors.append(f"면제 판정 예시{i + 1}: 판정 수식 없음")
    notes.append(f"면제 판정 예시 {len(samples)}건 기대값 일치 확인")
    return errors, notes
