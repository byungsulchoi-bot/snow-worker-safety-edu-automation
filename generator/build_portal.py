"""다운로드 자료실(웹앱) 생성: output/{연도}의 산출물을 portal/에 모아 index.html을 만든다.

사용법
  python generator/build_portal.py --year 2026 [--law-check logs/law_check_2026.json] [--basis-date 2026-10-01]

사전 조건
  logs/drive_{연도}.json  산출물을 Google Drive '제설근로자 안전교육 자료실/{연도}' 폴더에 올린 뒤
                          {"root_folder_id", "year_folder_id", "files": {파일명: Drive 파일 ID}} 로 저장

결과
  portal/index.html     게시할 페이지(자료 목록 내장, 파일은 Drive 링크)
  portal/manifest.json  연도별 게시 이력(누적)
"""
import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import ROOT, Ctx  # noqa: E402

PORTAL = ROOT / "portal"
MIME = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "zip": "application/zip",
}
GROUPS = [
    {"no": "안내", "title": "패키지 안내 (적용 법령·할 일)", "prefixes": ["00_"], "basis": ""},
    {"no": "1", "title": "채용 시 교육 면제 조건", "prefixes": ["01_"], "basis": "법 제30조제2항, 시행규칙 제26조제1항 후단·제27조제5항"},
    {"no": "2", "title": "하반기 정기교육 법정시간·교과목", "prefixes": ["02_"], "basis": "법 제29조제1항, 시행규칙 제26조제1항, 별표4·별표5"},
    {"no": "3", "title": "법정 교과목별 교안 PPT", "prefixes": ["03_"], "basis": "시행규칙 별표5, 제36조제1항"},
    {"no": "4", "title": "강사자격 및 교육시간표", "prefixes": ["04_"], "basis": "시행규칙 제26조제3항"},
    {"no": "5", "title": "교육계획보고 공문", "prefixes": ["05_"], "basis": ""},
    {"no": "6", "title": "교육일지 · 출석부 · 교육사진 양식", "prefixes": ["06_"], "basis": "안전보건교육규정 제4조제2호"},
    {"no": "7", "title": "교육결과보고 공문 및 첨부", "prefixes": ["07_"], "basis": ""},
    {"no": "8", "title": "법령준수 체크리스트 (근거조항 표기)", "prefixes": ["08_"], "basis": ""},
]


def build(year, law_check_path, basis_date):
    ctx = Ctx(year)
    src = ROOT / "output" / str(year)
    zip_src = ROOT / "output" / f"{year}_제설근로자_안전보건교육_패키지.zip"
    if not src.exists() or not zip_src.exists():
        raise SystemExit(f"먼저 generate.py build --year {year} 를 실행하세요")

    law = {"changed": False, "failed": False, "status_text": "", "details": []}
    if law_check_path:
        p = Path(law_check_path)
        law = json.loads((p if p.is_absolute() else ROOT / p).read_text(encoding="utf-8"))

    # Artifact는 Office·ZIP 파일을 직접 게시할 수 없으므로 Google Drive에 올린 파일 ID로 링크한다.
    drive_path = ROOT / "logs" / f"drive_{year}.json"
    if not drive_path.exists():
        raise SystemExit(f"{drive_path} 가 없습니다. 산출물을 Drive에 올리고 파일명→ID 매핑을 저장하세요")
    drive = json.loads(drive_path.read_text(encoding="utf-8"))
    ids = drive["files"]

    def link(name):
        fid = ids.get(name)
        if not fid:
            raise SystemExit(f"Drive 매핑에 없는 파일: {name}")
        return {"view": f"https://drive.google.com/file/d/{fid}/view",
                "download": f"https://drive.google.com/uc?export=download&id={fid}"}

    files = []
    for f in sorted(src.iterdir()):
        files.append({"filename": f.name, "bytes": f.stat().st_size, **link(f.name)})
    zip_entry = {"filename": zip_src.name, "bytes": zip_src.stat().st_size, **link(zip_src.name)}
    folder_url = f"https://drive.google.com/drive/folders/{drive['year_folder_id']}"

    hb = ctx.hours_by_subject()
    entry = {
        "year": year,
        "basis_date": basis_date or f"{year}-10-01",
        "posted_label": f"{(basis_date or f'{year}-10-01')} 기준 게시",
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "half": ctx.half_label,
        "required_hours": f"{ctx.required_hours:g}",
        "total_hours": ctx.total_hours,
        "instructor": ctx.cfg["instructor_role"].split("(")[0],
        "law_stamp": ctx.law_stamp,
        "law": {k: law.get(k) for k in ("changed", "failed", "status_text", "details")},
        "subjects": [{"no": no, "name": s["name"], "hours": hb[no]} for no, s in ctx.subjects.items()],
        "exemptions": [{k: e[k] for k in ("id", "condition", "effect", "basis")} for e in ctx.laws["hiring_exemptions"]],
        "files": files,
        "zip": zip_entry,
        "folder": folder_url,
    }

    PORTAL.mkdir(exist_ok=True)
    mpath = PORTAL / "manifest.json"
    manifest = json.loads(mpath.read_text(encoding="utf-8")) if mpath.exists() else {"years": []}
    manifest["years"] = [y for y in manifest["years"] if y["year"] != year] + [entry]
    manifest["years"].sort(key=lambda y: y["year"], reverse=True)
    manifest["groups"] = GROUPS
    mpath.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")

    tpl = (ROOT / "generator" / "portal_template.html").read_text(encoding="utf-8")
    data = json.dumps(manifest, ensure_ascii=False).replace("</", "<\\/")
    (PORTAL / "index.html").write_text(tpl.replace("/*__DATA__*/null", data), encoding="utf-8")
    print(json.dumps({"ok": True, "year": year, "index": str(PORTAL / "index.html"),
                      "files": len(files) + 1, "years": [y["year"] for y in manifest["years"]]}, ensure_ascii=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=datetime.now().year)
    ap.add_argument("--law-check")
    ap.add_argument("--basis-date")
    a = ap.parse_args()
    build(a.year, a.law_check, a.basis_date)
