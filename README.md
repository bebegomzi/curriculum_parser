# 고등학교 교육과정 파서·선택과목 시뮬레이터

학교 교육과정 편성표를 CSV로 구조화하고, 학생이 3개년 선택과목과 학점을 확인할 수 있게 만든 정적 웹 앱입니다.

## 지원 학교

| 학교 | 기준 | 원본 형식 | 웹 주소 |
|---|---:|---|---|
| 반곡고 | 2026학년도 입학생 | XLSX | 저장소 루트 `/curriculum_parser/` |
| 공주여고 | 2027학년도 입학생 | HWPX | `/curriculum_parser/?school=gjghs` |
| 웅상고 | 2026학년도 입학생 수정안 | XLSX | `/curriculum_parser/?school=woongsang` |

기존 공주여고 호환 주소와 `/curriculum_parser/woongsang/`, `/curriculum_parser/wshs.html`은 해당 학교 페이지로 자동 이동합니다.

## 구조

```text
curriculum_parser/
├─ index.html                         # 모든 학교가 공유하는 시뮬레이터
├─ course_metadata.js                 # 과목 설명·성적 산출 메타데이터
├─ admission_recommendations.js       # 대학 권장과목 데이터
├─ convert_curriculum_to_llm_csv.py   # 범용 XLSX 파서
├─ convert_hwpx_curriculum_to_llm_csv.py # 공주여고 양식 HWPX 파서
├─ 반곡고/                            # 원본 XLSX와 검수 CSV
├─ 공주여고/                          # 원본 HWPX와 검수 CSV, 호환 리디렉션
├─ gjghs/                             # 영문 호환 리디렉션
├─ woongsang/                         # 웅상고 원본·보완표·검수 CSV·리디렉션
├─ references/                        # 파서·권장과목 작성 참고 원자료
└─ tests/                             # 검수 CSV 기반 회귀 테스트
```

학교별로 달라지는 웹 설정은 `index.html`의 `SCHOOL_CONFIGS`에 모아 두었습니다. 새 학교는 학교 폴더와 CSV를 추가한 뒤 이 설정에 제목·CSV 경로를 등록하면 됩니다.

## 설치

Python 3.10 이상을 권장합니다.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 파싱

### XLSX

시트와 출력 접두사는 자동 감지할 수 있으며, 양식이 모호할 때에는 명시하는 편이 안전합니다.

```powershell
python convert_curriculum_to_llm_csv.py `
  "반곡고/2026학년도 반곡고등학교 입학생 3개년 교육과정 편성표.xlsx" `
  --sheet 반곡고 `
  --outdir 반곡고 `
  --prefix bangok_2026_curriculum
```

웅상고 원본은 과목유형·기본학점 수식이 포함되지 않은 외부 시트 `[1]숨김`을 참조합니다. `woongsang/course_overrides.csv`는 오류가 난 두 필드만 보완하며, 과목명·운영 학점·선택군·합계는 원본 편성표에서 직접 읽습니다.

```powershell
python convert_curriculum_to_llm_csv.py `
  "woongsang/2026학년도 입학생 3개년 교육과정 편성 수정(안)_웅상고.xlsx" `
  --sheet Sheet1 `
  --outdir woongsang `
  --prefix woongsang_2026_curriculum `
  --overrides woongsang/course_overrides.csv
```

### HWPX

현재 HWPX 파서는 공주여고 2027 편성표의 표 구조에 맞춘 어댑터입니다. 다른 학교 HWPX는 표 좌표를 먼저 확인한 뒤 어댑터를 확장해야 합니다.

```powershell
python convert_hwpx_curriculum_to_llm_csv.py `
  "공주여고/2027학년도 교육과정 편제(수정안)-변경 후.hwpx" `
  --outdir 공주여고 `
  --prefix gjghs_2027_curriculum
```

두 파서는 다음 파일을 만듭니다.

- `*_courses_llm.csv`: 과목 1개당 1행
- `*_selection_groups_llm.csv`: 학생 선택군별 요약
- `*_summary_llm.csv`: 학기 합계·창의적 체험활동 요약
- `*_readme.md`: 원본과 변환 규칙, 생성 행 수

## 검증

검수된 반곡고·공주여고·웅상고 CSV와 새 파싱 결과를 비교하고 웹 설정의 CSV 경로도 검사합니다.

```powershell
python -m unittest discover -s tests -v
```

웹 화면은 `fetch`로 CSV를 읽으므로 파일을 직접 열지 말고 로컬 HTTP 서버를 사용합니다.

```powershell
python -m http.server 8000
```

- 반곡고: `http://localhost:8000/`
- 공주여고: `http://localhost:8000/?school=gjghs`
- 웅상고: `http://localhost:8000/?school=woongsang`

## 새 학교 추가 순서

1. 학교명 폴더에 원본 편성표를 보관합니다.
2. XLSX 범용 파서로 먼저 변환하고, 양식 차이가 있으면 감지 규칙을 보완합니다.
3. 생성 CSV의 과목 수, 선택군 범위·선택 수·과목당 학점, 학기 합계를 원본과 대조합니다.
4. 검수된 CSV와 원본을 함께 커밋하고 회귀 테스트를 추가합니다.
5. `SCHOOL_CONFIGS`에 학교 키와 CSV 경로를 등록합니다.

선택군의 병합 셀에는 선택군 전체 학점이 들어갈 수 있습니다. 파서는 원본 총학점을 첫 행에만 보존하고, 개별 과목의 정규화 학점에는 `택N (학점)`에서 읽은 과목당 학점을 기록합니다.

## 제작자

- 세종특별자치시교육청 국어 교사 정현민
- Copyright (c) 2026 Hyunmin Jeong. All rights reserved.
