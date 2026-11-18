"""공통: 데이터 로드, 연도·시간 계산, 자리표시."""
import json
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
BLANK = "〔　　　　〕"
FONT = "맑은 고딕"


def load_json(rel):
    with open(ROOT / rel, encoding="utf-8") as f:
        return json.load(f)


def ph(value, width=4):
    """설정값이 비어 있으면 공란 자리표시를 돌려준다."""
    return value if value else "〔" + "　" * width + "〕"


@dataclass
class Session:
    day: int
    start: str
    end: str
    subject_no: int
    subject_name: str
    hours: int
    method: str
    instructor: str
    note: str = ""


class Ctx:
    def __init__(self, year: int):
        self.year = year
        self.laws = load_json("data/laws.json")
        cfg_name = "config/settings.json" if (ROOT / "config/settings.json").exists() else "config/settings.example.json"
        self.cfg = load_json(cfg_name)
        self.cur = load_json("data/curriculum.json")["subjects"]
        self.P = self.cfg["placeholders"]
        self.subjects = {s["no"]: s for s in self.laws["regular_subjects"]}
        self.out = ROOT / "output" / str(year)
        self.out.mkdir(parents=True, exist_ok=True)
        self.sessions = self._build_sessions()

    # ---- 법정 시간 ----
    @property
    def legal_hours(self):
        return self.laws["hours"]["regular_per_half"]

    @property
    def required_hours(self):
        """감면 적용 시 법정시간의 50%(시행규칙 제27조제1항), 아니면 법정시간."""
        h = self.legal_hours
        return h / 2 if self.cfg.get("apply_no_accident_reduction") else h

    @property
    def half_label(self):
        return f"{self.year}년 하반기"

    @property
    def period(self):
        return f"{self.year}. 7. 1. ~ {self.year}. 12. 31."

    @property
    def law_stamp(self):
        s = self.laws["sources"]
        return (f"적용 법령 기준: {s['act']['name']} [{s['act']['effective']} {s['act']['number']}], "
                f"{s['rule']['name']} [{s['rule']['effective']} {s['rule']['number']}] 별표4·5({s['annex5']['amended']} 개정), "
                f"{s['notice']['name']} [{s['notice']['effective']} {s['notice']['number']}] / 법령 확인일 {self.laws['checked_at']}")

    def _build_sessions(self):
        online = set(self.cfg.get("online_subjects", []))
        lunch_s = datetime.strptime(self.cfg["lunch"]["start"], "%H:%M")
        lunch_e = datetime.strptime(self.cfg["lunch"]["end"], "%H:%M")
        cursor = {}
        res = []
        for item in self.cfg["schedule"]:
            d = item["day"]
            t = cursor.get(d, datetime.strptime(self.cfg["day_start"], "%H:%M"))
            if t >= lunch_s and t < lunch_e:
                t = lunch_e
            end = t + timedelta(hours=item["hours"])
            note = ""
            if t < lunch_s < end:  # 점심시간을 걸치면 그만큼 종료 시각을 늦춤
                end = end + (lunch_e - lunch_s)
                note = f"점심 {self.cfg['lunch']['start']}~{self.cfg['lunch']['end']} 제외"
            no = item["subject"]
            is_online = no in online
            res.append(Session(
                day=d, start=t.strftime("%H:%M"), end=end.strftime("%H:%M"),
                subject_no=no, subject_name=self.subjects[no]["name"], hours=item["hours"],
                method="인터넷 원격교육" if is_online else "집체교육",
                instructor=("위탁·원격 과정 " + BLANK) if is_online else self.cfg["instructor_role"],
                note=note,
            ))
            cursor[d] = end
        return res

    def hours_by_subject(self):
        h = {no: 0 for no in self.subjects}
        for s in self.sessions:
            h[s.subject_no] += s.hours
        return h

    @property
    def total_hours(self):
        return sum(s.hours for s in self.sessions)

    def path(self, name):
        return self.out / name
