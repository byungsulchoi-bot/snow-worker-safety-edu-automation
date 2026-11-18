"""다운로드 자료실(웹앱) 생성.

두 가지 게시 대상을 지원한다.
  --target site   (기본) Vercel 등 정적 호스팅용. 산출물을 site/files/{연도}/ 에 복사하고 사이트가 직접 내려받게 한다.
                  site/ 는 Git에 커밋되며 GitHub에 푸시하면 Vercel이 자동 재배포한다.
  --target portal claude.ai Artifact용. Office·ZIP 파일을 직접 게시할 수 없으므로 Google Drive 링크를 쓴다.
                  사전 조건: logs/drive_{연도}.json = {"root_folder_id", "year_folder_id", "files": {파일명: Drive ID}}

사용법
  python generator/build_portal.py --year 2026 --law-check logs/law_check_2026.json [--basis-date 2026-10-01] [--target site|portal]
"""
import argparse
import json
import shutil
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from common import ROOT, Ctx  # noqa: E402

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


def _drive_links(year):
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
    return link, f"https://drive.google.com/drive/folders/{drive['year_folder_id']}"


def build(year, law_check_path, basis_date, target):
    ctx = Ctx(year)
    out_dir = ROOT / ("site" if target == "site" else "portal")
    src = ROOT / "output" / str(year)
    zip_src = ROOT / "output" / f"{year}_제설근로자_안전보건교육_패키지.zip"
    if not src.exists() or not zip_src.exists():
        raise SystemExit(f"먼저 generate.py build --year {year} 를 실행하세요")

    law = {"changed": False, "failed": False, "status_text": "", "details": []}
    if law_check_path:
        p = Path(law_check_path)
        law = json.loads((p if p.is_absolute() else ROOT / p).read_text(encoding="utf-8"))

    out_dir.mkdir(exist_ok=True)
    srcs = sorted(src.iterdir())
    files, folder_url = [], ""
    if target == "site":
        dest = out_dir / "files" / str(year)
        if dest.exists():
            shutil.rmtree(dest)
        dest.mkdir(parents=True)
        for i, f in enumerate(srcs, 1):
            rel = f"files/{year}/{i:02d}{f.suffix}"   # URL은 ASCII, 저장 파일명은 download 속성으로 한글 유지
            shutil.copy2(f, out_dir / rel)
            files.append({"filename": f.name, "bytes": f.stat().st_size, "download": rel})
        zrel = f"files/{year}/package.zip"
        shutil.copy2(zip_src, out_dir / zrel)
        zip_entry = {"filename": zip_src.name, "bytes": zip_src.stat().st_size, "download": zrel}
    else:
        link, folder_url = _drive_links(year)
        files = [{"filename": f.name, "bytes": f.stat().st_size, **link(f.name)} for f in srcs]
        zip_entry = {"filename": zip_src.name, "bytes": zip_src.stat().st_size, **link(zip_src.name)}

    hb = ctx.hours_by_subject()
    bd = basis_date or f"{year}-10-01"
    entry = {
        "year": year,
        "basis_date": bd,
        "posted_label": f"{bd} 기준 게시",
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

    mpath = out_dir / "manifest.json"
    manifest = json.loads(mpath.read_text(encoding="utf-8")) if mpath.exists() else {"years": []}
    manifest["years"] = [y for y in manifest["years"] if y["year"] != year] + [entry]
    manifest["years"].sort(key=lambda y: y["year"], reverse=True)
    manifest["groups"] = GROUPS
    manifest["mode"] = target
    mpath.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")

    tpl = (ROOT / "generator" / "portal_template.html").read_text(encoding="utf-8")
    data = json.dumps(manifest, ensure_ascii=False).replace("</", "<\\/")
    html = tpl.replace("/*__DATA__*/null", data)
    if target == "site":
        # 일반 웹 호스팅은 문서 골격이 필요하다(Artifact는 게시 시 자동으로 감싼다).
        cut = html.index('<div class="tape"')
        html = ('<!doctype html>\n<html lang="ko">\n<head>\n<meta charset="utf-8">\n'
                '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">\n'
                '<style>html{-webkit-text-size-adjust:100%}body{margin:0}[hidden]{display:none!important}</style>\n'
                + html[:cut] + '</head>\n<body>\n' + html[cut:] + '\n</body>\n</html>\n')
    (out_dir / "index.html").write_text(html, encoding="utf-8")
    print(json.dumps({"ok": True, "target": target, "year": year, "index": str(out_dir / "index.html"),
                      "files": len(files) + 1, "years": [y["year"] for y in manifest["years"]]}, ensure_ascii=False))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--year", type=int, default=datetime.now().year)
    ap.add_argument("--law-check")
    ap.add_argument("--basis-date")
    ap.add_argument("--target", choices=["site", "portal"], default="site")
    a = ap.parse_args()
    build(a.year, a.law_check, a.basis_date, a.target)
