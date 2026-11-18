# 제설기간제근로자 안전보건교육 업무자동화

매년 10월 1·15·25일 09:00(Claude 데스크톱 앱 예약 작업 `snow-worker-safety-edu-package`)에
법령 개정 확인 → 교육 패키지 20종 생성·검증 → 본인 Gmail 발송(연 1회)을 수행합니다.

## 웹 자료실
- 주소: https://claude.ai/artifact/3kUse4Bwks34GdSJruRhEa (비공개, 다른 사람에게는 페이지의 공유 메뉴로 공유)
- 매년 10월 1일 실행 때 산출물을 Google Drive `제설근로자 안전교육 자료실/{연도}` 폴더에 올리고, 자료실을 같은 주소로 갱신합니다(지난 연도 자료도 연도 선택으로 계속 열람).
- 수동 갱신: Drive 업로드 후 `logs/drive_{연도}.json` 작성 → `python generator/build_portal.py --year {연도} --law-check logs/law_check_{연도}.json` → `portal/index.html`을 같은 주소로 재게시

## 처음 한 번 할 일
0. 새로 내려받은(clone) 경우: `python -m pip install -r requirements.txt` 후 `config/settings.example.json`을 `config/settings.json`으로 복사
   (settings.json, 발송 이력, Drive 파일 ID는 개인정보라 Git에 올리지 않습니다)
1. `config/settings.json`의 `mail_to`에 받을 Gmail 주소 입력 (비어 있으면 임시보관함에 저장만 함)
2. 필요 시 `placeholders`(기관명·부서·수신·장소·강사 성명·결재라인) 입력 → 공문에 자동 반영, 비우면 〔　〕 공란

## 수동 실행
```bash
python generator/generate.py build --year 2026 --law-check logs/law_check_2026.json
```
결과: `output/2026/` 폴더와 `output/2026_제설근로자_안전보건교육_패키지.zip`

## 자주 바꾸는 설정 (`config/settings.json`)
| 키 | 의미 |
|---|---|
| `online_subjects` | 원격교육으로 돌릴 과목 번호. 예: `[6,7,8]` (과목당 1시간 이상) |
| `apply_no_accident_reduction` | 전년도 무재해 시 정기교육 50% 감면 적용(시행규칙 제27조제1항) |
| `schedule` | 일차·과목·시간 편성 (8과목 모두 1시간 이상, 합계 12시간 이상이어야 검증 통과) |

## 구성
- `data/laws.json` 법정 기준 단일 출처(시간·과목·강사·면제·근거조항)
- `data/law_baseline.json` 법령 개정 비교 기준값
- `data/curriculum.json` 과목별 교안 원고(제설작업 특화)
- `generator/` 생성기(docx·xlsx·pptx)와 검증기
- `logs/` 법령 확인 결과, 발송 이력(`sent_log.json`)

필요 패키지: `python -m pip install python-docx openpyxl python-pptx`
