"""
교육과정 편성표(.xlsx)를 LLM이 읽기 쉬운 CSV 3종으로 변환합니다.

전제
- '반곡고' 시트의 헤더/열 배열은 동일하고, 과목/요약 테이블의 길이만 달라진다는 전제입니다.
- 병합 셀은 좌상단 값을 병합 범위 전체에 전파합니다.
- 각 학기 칸의 숫자는 학점/주당 시수/이수 단위로 해석합니다.

출력
1) *_courses_llm.csv           : 과목별 평면표
2) *_selection_groups_llm.csv  : 학생 선택군 요약
3) *_summary_llm.csv           : 하단 합계/창체 요약
4) *_readme.md                 : 해석 규칙

사용 예
python convert_bangok_curriculum_to_llm_csv.py "input.xlsx" --sheet "반곡고" --outdir "./out" --prefix "bangok_2026_curriculum"
"""

from __future__ import annotations

import argparse
import csv
import re
from copy import copy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ===== 고정 레이아웃 설정 =====

DEFAULT_SHEET = "반곡고"
HEADER_LAST_ROW = 7
DATA_START_ROW = 8

# 1-indexed Excel columns
COL = {
    "교과군": 1,                 # A
    "과목명": 2,                 # B
    "학교지정": 3,               # C
    "학년선택": 4,               # D
    "공통": 5,                   # E
    "일반": 6,                   # F
    "진로": 7,                   # G
    "융합": 8,                   # H
    "비고_전문_고시외_표6": 9,   # I
    "학생선택_1학기": 10,        # J
    "학생선택_2학기": 11,        # K
    "기본학점": 12,              # L
    "1-1": 13,                   # M
    "1-2": 14,                   # N
    "2-1": 15,                   # O
    "2-2": 16,                   # P
    "3-1": 17,                   # Q
    "3-2": 18,                   # R
    "편성학점": 19,              # S
    "교과군별이수학점": 20,      # T
    "필수이수학점": 21,          # U
    "과학중점": 22,              # V
    "정보중점": 23,              # W
    "사회중점": 24,              # X
    "인문중점": 25,              # Y
    "캠공개설여부": 26,          # Z
}

SEMESTER_COLS = {
    "1-1": COL["1-1"],
    "1-2": COL["1-2"],
    "2-1": COL["2-1"],
    "2-2": COL["2-2"],
    "3-1": COL["3-1"],
    "3-2": COL["3-2"],
}

COURSE_FIELDNAMES = [
    "원본행",
    "교과군",
    "과목명",
    "학교지정",
    "학년선택",
    "과목유형",
    "비고_전문_고시외_표6",
    "학생선택_원문",
    "선택군_코드",
    "선택군_학년학기",
    "선택군_선택수",
    "선택군_과목당학점",
    "선택군_총학점",
    "선택군_행범위",
    "기본학점",
    "1-1_원본",
    "1-2_원본",
    "2-1_원본",
    "2-2_원본",
    "3-1_원본",
    "3-2_원본",
    "1-1_학점",
    "1-2_학점",
    "2-1_학점",
    "2-2_학점",
    "3-1_학점",
    "3-2_학점",
    "편성학점",
    "교과군별이수학점",
    "필수이수학점",
    "과학중점",
    "정보중점",
    "사회중점",
    "인문중점",
    "캠공개설여부",
    "비고",
]

SELECTION_FIELDNAMES = [
    "선택군_코드",
    "선택군_학년학기",
    "선택군_원문",
    "선택군_선택수",
    "선택군_과목당학점",
    "선택군_총학점",
    "선택군_행범위",
    "대상과목수",
    "대상과목목록",
]

SUMMARY_FIELDNAMES = [
    "원본행",
    "항목",
    "세부항목",
    "1-1",
    "1-2",
    "2-1",
    "2-2",
    "3-1",
    "3-2",
    "편성",
    "교과군별이수학점",
    "필수이수학점",
    "비고",
]


@dataclass
class SelectionGroup:
    code: str
    semester: str
    raw: str
    choose_count: str
    credit_per_course: str
    total_credit: str
    row_start: int
    row_end: int
    subjects: List[str]

    @property
    def row_range(self) -> str:
        return f"{self.row_start}:{self.row_end}"

    def as_row(self) -> Dict[str, str]:
        return {
            "선택군_코드": self.code,
            "선택군_학년학기": self.semester,
            "선택군_원문": self.raw,
            "선택군_선택수": self.choose_count,
            "선택군_과목당학점": self.credit_per_course,
            "선택군_총학점": self.total_credit,
            "선택군_행범위": self.row_range,
            "대상과목수": str(len(self.subjects)),
            "대상과목목록": "; ".join(self.subjects),
        }


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def true_marker(value: Any) -> str:
    """○ 또는 이미 True로 볼 수 있는 값만 True, 나머지는 빈칸."""
    text = clean(value)
    return "True" if text in {"○", "TRUE", "True", "true", "1"} else ""


def number_text(value: Any) -> str:
    """CSV에 넣을 숫자 텍스트. 비숫자 메모는 그대로 보존."""
    return clean(value)


def extract_first_number(value: Any) -> str:
    """학기 칸에서 첫 숫자만 학점으로 추출. 예: '3\n(공동)' -> '3'."""
    text = clean(value)
    if not text:
        return ""
    m = re.search(r"\d+(?:\.\d+)?", text)
    if not m:
        return ""
    n = float(m.group())
    return str(int(n)) if n.is_integer() else str(n)


def copy_ws_with_filled_merges(ws: Worksheet) -> Worksheet:
    """
    원본 시트를 메모리상에서 복사한 뒤 병합 셀을 해제하고 값을 전파합니다.
    원본 파일은 변경하지 않습니다.
    """
    new_ws = copy(ws)
    for merged in list(new_ws.merged_cells.ranges):
        value = new_ws.cell(merged.min_row, merged.min_col).value
        new_ws.unmerge_cells(str(merged))
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                new_ws.cell(row, col).value = value
    return new_ws


def find_summary_start_row(ws: Worksheet) -> int:
    """
    과목 목록이 끝나고 하단 요약이 시작되는 행을 찾습니다.
    현재 양식에서는 A:L 병합 합계행 또는 A/B열 과목명 소멸 지점을 기준으로 잡습니다.
    """
    for r in range(DATA_START_ROW, ws.max_row + 1):
        a = clean(ws.cell(r, COL["교과군"]).value)
        b = clean(ws.cell(r, COL["과목명"]).value)
        # A:L 병합 합계행은 병합 전파 후 A, B가 같은 값으로 채워질 수 있음.
        joined = normalize_space(" ".join([a, b]))
        if "창의적" in joined or "총 이수" in joined or "총계" in joined or "합계" in joined:
            return r
        # 과목명 없이 학기/편성 쪽 숫자만 나오기 시작하면 요약 영역으로 간주
        if r > DATA_START_ROW and not b:
            right_values = [clean(ws.cell(r, c).value) for c in range(COL["1-1"], COL["필수이수학점"] + 1)]
            if any(right_values):
                return r
    return ws.max_row + 1


def course_type(ws: Worksheet, row: int) -> str:
    """공통/일반/진로/융합 중 주 과목유형을 반환합니다."""
    for name in ["공통", "일반", "진로", "융합"]:
        if true_marker(ws.cell(row, COL[name]).value):
            return name
    return ""


def split_note_suffix(raw: Any) -> str:
    """
    <표6>/전문/고시외 칸은 과목유형이 아니라 별도 비고성 suffix입니다.
    예: '전문', '고시외', '표6'
    """
    return clean(raw)


def parse_selection_meta(raw: str) -> Tuple[str, str, str]:
    """
    선택군 원문에서 택N과 과목당 학점을 추출합니다.
    예: '택4 (3)' -> ('4', '3', '12')
    """
    text = clean(raw)
    choose = ""
    credit = ""

    m_choose = re.search(r"택\s*(\d+)", text)
    if m_choose:
        choose = m_choose.group(1)

    m_credit = re.search(r"\((\d+(?:\.\d+)?)\)", text)
    if m_credit:
        credit = m_credit.group(1)

    total = ""
    if choose and credit:
        total_num = float(choose) * float(credit)
        total = str(int(total_num)) if total_num.is_integer() else str(total_num)

    return choose, credit, total


def compact_selection_label(raw: str) -> str:
    """선택군 코드용 압축 원문. 첫 줄의 '택4 (3)'를 '택4(3)'으로 변환."""
    first_line = clean(raw).splitlines()[0] if clean(raw) else ""
    return re.sub(r"\s+", "", first_line)


def detect_selection_group_for_row(ws: Worksheet, row: int) -> Tuple[str, str, str, str, str, str, str]:
    """
    해당 과목 행의 학생선택 그룹을 추정합니다.
    J열 = 1학기 선택군, K열 = 2학기 선택군이지만,
    같은 '택4 (3)'이라도 실제 운영 학년-학기에 따라 별도 코드로 분리합니다.

    반환:
    (원문, 코드, 학년학기, 선택수, 과목당학점, 총학점, 행범위)
    """
    sem_values = {sem: clean(ws.cell(row, col).value) for sem, col in SEMESTER_COLS.items()}

    # 1학기 학점이 있는 경우 J열 선택군, 2학기 학점이 있는 경우 K열 선택군을 우선 사용
    candidates: List[Tuple[str, str]] = []
    for sem in ["1-1", "2-1", "3-1"]:
        if sem_values.get(sem):
            candidates.append((sem, clean(ws.cell(row, COL["학생선택_1학기"]).value)))
    for sem in ["1-2", "2-2", "3-2"]:
        if sem_values.get(sem):
            candidates.append((sem, clean(ws.cell(row, COL["학생선택_2학기"]).value)))

    # 선택군 원문이 실제로 있는 후보만 사용
    candidates = [(sem, raw) for sem, raw in candidates if raw]
    if not candidates:
        return "", "", "", "", "", "", ""

    # 여러 후보가 있으면 첫 번째. 현재 양식에서는 선택 과목이 한 학기에만 편성되는 구조를 전제.
    semester, raw = candidates[0]
    choose, credit, total = parse_selection_meta(raw)
    label = compact_selection_label(raw)
    code = f"{semester}-{label}" if label else ""

    # 병합 전파된 선택군 값이 이어지는 범위를 위/아래로 확장
    sel_col = COL["학생선택_1학기"] if semester.endswith("-1") else COL["학생선택_2학기"]
    row_start = row
    row_end = row
    while row_start > DATA_START_ROW and clean(ws.cell(row_start - 1, sel_col).value) == raw:
        row_start -= 1
    while row_end < ws.max_row and clean(ws.cell(row_end + 1, sel_col).value) == raw:
        row_end += 1
    row_range = f"{row_start}:{row_end}"

    return raw, code, semester, choose, credit, total, row_range


def build_selection_groups(course_rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    groups: Dict[str, SelectionGroup] = {}
    for row in course_rows:
        code = row.get("선택군_코드", "")
        if not code:
            continue
        start, end = row["선택군_행범위"].split(":")
        if code not in groups:
            groups[code] = SelectionGroup(
                code=code,
                semester=row["선택군_학년학기"],
                raw=row["학생선택_원문"],
                choose_count=row["선택군_선택수"],
                credit_per_course=row["선택군_과목당학점"],
                total_credit=row["선택군_총학점"],
                row_start=int(start),
                row_end=int(end),
                subjects=[],
            )
        groups[code].subjects.append(row["과목명"])
        groups[code].row_start = min(groups[code].row_start, int(start))
        groups[code].row_end = max(groups[code].row_end, int(end))

    return [groups[k].as_row() for k in sorted(groups.keys())]


def build_course_rows(ws: Worksheet, summary_start: int) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []

    for r in range(DATA_START_ROW, summary_start):
        subject = clean(ws.cell(r, COL["과목명"]).value)
        if not subject:
            continue

        sel_raw, sel_code, sel_sem, sel_choose, sel_credit, sel_total, sel_range = detect_selection_group_for_row(ws, r)

        row: Dict[str, str] = {
            "원본행": str(r),
            "교과군": clean(ws.cell(r, COL["교과군"]).value),
            "과목명": subject,
            "학교지정": true_marker(ws.cell(r, COL["학교지정"]).value),
            "학년선택": true_marker(ws.cell(r, COL["학년선택"]).value),
            "과목유형": course_type(ws, r),
            "비고_전문_고시외_표6": split_note_suffix(ws.cell(r, COL["비고_전문_고시외_표6"]).value),
            "학생선택_원문": sel_raw,
            "선택군_코드": sel_code,
            "선택군_학년학기": sel_sem,
            "선택군_선택수": sel_choose,
            "선택군_과목당학점": sel_credit,
            "선택군_총학점": sel_total,
            "선택군_행범위": sel_range,
            "기본학점": number_text(ws.cell(r, COL["기본학점"]).value),
            "편성학점": number_text(ws.cell(r, COL["편성학점"]).value),
            "교과군별이수학점": number_text(ws.cell(r, COL["교과군별이수학점"]).value),
            "필수이수학점": number_text(ws.cell(r, COL["필수이수학점"]).value),
            "과학중점": true_marker(ws.cell(r, COL["과학중점"]).value),
            "정보중점": true_marker(ws.cell(r, COL["정보중점"]).value),
            "사회중점": true_marker(ws.cell(r, COL["사회중점"]).value),
            "인문중점": true_marker(ws.cell(r, COL["인문중점"]).value),
            "캠공개설여부": true_marker(ws.cell(r, COL["캠공개설여부"]).value),
            "비고": "",
        }

        for sem, col in SEMESTER_COLS.items():
            original = number_text(ws.cell(r, col).value)
            row[f"{sem}_원본"] = original
            row[f"{sem}_학점"] = extract_first_number(original)

        rows.append(row)

    return rows


def build_summary_rows(ws: Worksheet, summary_start: int) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    for r in range(summary_start, ws.max_row + 1):
        values = [clean(ws.cell(r, c).value) for c in range(1, COL["필수이수학점"] + 1)]
        if not any(values):
            continue

        # 완전히 제목/공백성 행은 제외하지 않고, 숫자 또는 항목이 있으면 보존
        row = {
            "원본행": str(r),
            "항목": clean(ws.cell(r, COL["교과군"]).value),
            "세부항목": clean(ws.cell(r, COL["과목명"]).value),
            "1-1": number_text(ws.cell(r, COL["1-1"]).value),
            "1-2": number_text(ws.cell(r, COL["1-2"]).value),
            "2-1": number_text(ws.cell(r, COL["2-1"]).value),
            "2-2": number_text(ws.cell(r, COL["2-2"]).value),
            "3-1": number_text(ws.cell(r, COL["3-1"]).value),
            "3-2": number_text(ws.cell(r, COL["3-2"]).value),
            "편성": number_text(ws.cell(r, COL["편성학점"]).value),
            "교과군별이수학점": number_text(ws.cell(r, COL["교과군별이수학점"]).value),
            "필수이수학점": number_text(ws.cell(r, COL["필수이수학점"]).value),
            "비고": "",
        }
        if any(v for k, v in row.items() if k != "원본행"):
            rows.append(row)
    return rows


def write_csv(path: Path, fieldnames: List[str], rows: Iterable[Dict[str, str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def write_readme(path: Path, source_file: str, sheet: str, course_count: int, group_count: int, summary_count: int) -> None:
    text = f"""# 교육과정 편성표 LLM용 CSV 해석 규칙

## 원본
- 파일: `{source_file}`
- 시트: `{sheet}`

## 생성 파일
- `*_courses_llm.csv`: 과목 1개를 1행으로 정리한 평면표
- `*_selection_groups_llm.csv`: 학생 선택군 단위 요약표
- `*_summary_llm.csv`: 하단 합계/창의적 체험활동 등 요약 영역

## 변환 원칙
1. 병합 셀은 좌상단 값을 병합 범위 전체에 전파했다.
2. `1-1`, `1-2`, `2-1`, `2-2`, `3-1`, `3-2`의 숫자는 해당 학기 편성 학점이며, 이 양식에서는 주당 시수/이수 단위와 같은 의미로 해석한다.
3. `○` 표시는 `True`로 변환하고, 해당하지 않는 칸은 빈칸으로 둔다.
4. `과목유형`은 `공통`, `일반`, `진로`, `융합` 중 하나로 정리한다.
5. `전문`, `고시외`, `표6` 등은 과목유형이 아니라 `비고_전문_고시외_표6`에 별도로 둔다.
6. 학생 선택군은 같은 원문이라도 운영 학년-학기가 다르면 별도 선택군으로 본다. 예: `2-1-택4(3)`과 `2-2-택4(3)`은 다른 선택군이다.
7. `과학중점`, `정보중점`, `사회중점`, `인문중점`은 별도 boolean 열로 분리한다.

## 생성 결과 개수
- 과목 행: {course_count}
- 선택군: {group_count}
- 요약 행: {summary_count}
"""
    path.write_text(text, encoding="utf-8")


def convert(input_path: Path, sheet_name: str, outdir: Path, prefix: str) -> Dict[str, Path]:
    wb = load_workbook(input_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"시트 '{sheet_name}'을 찾을 수 없습니다. 사용 가능 시트: {wb.sheetnames}")

    ws = copy_ws_with_filled_merges(wb[sheet_name])
    summary_start = find_summary_start_row(ws)

    course_rows = build_course_rows(ws, summary_start)
    selection_rows = build_selection_groups(course_rows)
    summary_rows = build_summary_rows(ws, summary_start)

    outdir.mkdir(parents=True, exist_ok=True)
    courses_path = outdir / f"{prefix}_courses_llm.csv"
    groups_path = outdir / f"{prefix}_selection_groups_llm.csv"
    summary_path = outdir / f"{prefix}_summary_llm.csv"
    readme_path = outdir / f"{prefix}_readme.md"

    write_csv(courses_path, COURSE_FIELDNAMES, course_rows)
    write_csv(groups_path, SELECTION_FIELDNAMES, selection_rows)
    write_csv(summary_path, SUMMARY_FIELDNAMES, summary_rows)
    write_readme(readme_path, input_path.name, sheet_name, len(course_rows), len(selection_rows), len(summary_rows))

    return {
        "courses": courses_path,
        "selection_groups": groups_path,
        "summary": summary_path,
        "readme": readme_path,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="교육과정 편성표 xlsx를 LLM용 CSV로 변환합니다.")
    parser.add_argument("input", type=Path, help="입력 xlsx 파일 경로")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help=f"변환할 시트명. 기본값: {DEFAULT_SHEET}")
    parser.add_argument("--outdir", type=Path, default=Path("."), help="출력 폴더. 기본값: 현재 폴더")
    parser.add_argument("--prefix", default="curriculum", help="출력 파일명 prefix. 기본값: curriculum")
    args = parser.parse_args()

    paths = convert(args.input, args.sheet, args.outdir, args.prefix)
    print("변환 완료")
    for key, path in paths.items():
        print(f"- {key}: {path}")


if __name__ == "__main__":
    main()
