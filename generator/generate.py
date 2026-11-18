"""제설기간제근로자 안전보건교육 패키지 생성기.

사용법
  python generator/generate.py build --year 2027 [--law-check logs/law_check_2027.json]
  python generator/generate.py check-sent --year 2027
  python generator/generate.py mark-sent --year 2027 --message-id <Gmail id>

build 결과는 마지막 줄에 JSON 한 줄로 출력한다(예약 작업이 읽음).
"""
import argparse
import json
import sys
import zipfile
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from common import ROOT, Ctx  # noqa: E402
import docx_builder  # noqa: E402
import pptx_builder  # noqa: E402
import validate  # noqa: E402
import xlsx_builder  # noqa: E402

SENT_LOG = ROOT / "logs" / "sent_log.json"
MAX_ZIP = 25 * 1024 * 1024


def _load_sent():
    if SENT_LOG.exists():
        return json.loads(SENT_LOG.read_text(encoding="utf-8"))
    return {}


def cmd_build(year, law_check_path):
    ctx = Ctx(year)
    for old in ctx.out.glob("*"):
        if old.is_file():
            old.unlink()
    law_check = {"status_text": "법령 개정 확인을 실행하지 않음(수동 생성)", "changed": False, "failed": False, "details": []}
    if law_check_path:
        p = Path(law_check_path)
        p = p if p.is_absolute() else ROOT / p
        law_check = json.loads(p.read_text(encoding="utf-8"))

    docx_builder.build_all(ctx)
    samples = xlsx_builder.build_all(ctx)
    pptx_builder.build_all(ctx)
    files = sorted(p.name for p in ctx.out.iterdir())
    docx_builder.build_readme(ctx, law_check, sorted(files + ["00_README_패키지안내.docx"]))

    errors, notes = validate.run(ctx, samples)

    zip_path = ROOT / "output" / f"{year}_제설근로자_안전보건교육_패키지.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(ctx.out.iterdir()):
            z.write(p, arcname=f"{year}_제설근로자_안전보건교육/{p.name}")
    size = zip_path.stat().st_size
    if size > MAX_ZIP:
        errors.append(f"ZIP {size / 1e6:.1f}MB > 25MB (Gmail 첨부 한도)")

    result = {
        "year": year,
        "ok": not errors,
        "errors": errors,
        "notes": notes,
        "files": sorted(p.name for p in ctx.out.iterdir()),
        "zip": str(zip_path),
        "zip_bytes": size,
        "law_stamp": ctx.law_stamp,
        "total_hours": ctx.total_hours,
        "required_hours": ctx.required_hours,
    }
    print(json.dumps(result, ensure_ascii=False))
    return 0 if not errors else 1


def cmd_check_sent(year):
    rec = _load_sent().get(str(year))
    print(json.dumps({"year": year, "sent": bool(rec), "record": rec}, ensure_ascii=False))
    return 0


def cmd_mark_sent(year, message_id):
    log = _load_sent()
    log[str(year)] = {"sent_at": datetime.now().isoformat(timespec="seconds"), "message_id": message_id}
    SENT_LOG.parent.mkdir(parents=True, exist_ok=True)
    SENT_LOG.write_text(json.dumps(log, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps({"year": year, "marked": True}, ensure_ascii=False))
    return 0


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("--year", type=int, default=datetime.now().year)
    b.add_argument("--law-check")
    c = sub.add_parser("check-sent")
    c.add_argument("--year", type=int, default=datetime.now().year)
    m = sub.add_parser("mark-sent")
    m.add_argument("--year", type=int, default=datetime.now().year)
    m.add_argument("--message-id", required=True)
    a = ap.parse_args()
    if a.cmd == "build":
        return cmd_build(a.year, a.law_check)
    if a.cmd == "check-sent":
        return cmd_check_sent(a.year)
    return cmd_mark_sent(a.year, a.message_id)


if __name__ == "__main__":
    sys.exit(main())
