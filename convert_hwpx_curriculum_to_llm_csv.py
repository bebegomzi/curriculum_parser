"""
공주여자고등학교 HWPX 교육과정 편성표를 LLM용 CSV 3종으로 변환합니다.
"""

from __future__ import annotations

import argparse
import csv
import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


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
    return str(value).strip().replace("\u3000", " ")


def normalize_space(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def parse_hwpx_table(tbl) -> List[List[str]]:
    """HWPX의 좌표 지정 방식 셀들을 분석하여 병합 영역이 모두 복원된 2차원 리스트(Grid)로 반환합니다."""
    max_row = 0
    max_col = 0
    cells_data = []
    
    for tc in tbl.iter():
        if not tc.tag.endswith('tc'):
            continue
        
        cell_addr = tc.find('.//{*}cellAddr')
        cell_span = tc.find('.//{*}cellSpan')
        if cell_addr is None or cell_span is None:
            continue
            
        r_addr = int(cell_addr.attrib['rowAddr'])
        c_addr = int(cell_addr.attrib['colAddr'])
        r_span = int(cell_span.attrib['rowSpan'])
        c_span = int(cell_span.attrib['colSpan'])
        
        # 텍스트 추출
        text_parts = []
        for elem in tc.iter():
            if elem.tag.endswith('t') and elem.text:
                text_parts.append(elem.text)
        text = "".join(text_parts).strip()
        
        cells_data.append((r_addr, c_addr, r_span, c_span, text))
        max_row = max(max_row, r_addr + r_span)
        max_col = max(max_col, c_addr + c_span)
        
    grid = [["" for _ in range(max_col)] for _ in range(max_row)]
    for r_addr, c_addr, r_span, c_span, text in cells_data:
        for r in range(r_addr, r_addr + r_span):
            for c in range(c_addr, c_addr + c_span):
                grid[r][c] = text
                
    return grid


def extract_selection_meta(raw: str) -> str:
    text = clean(raw)
    choose = ""
    
    m_choose = re.search(r"택\s*(\d+)", text)
    if m_choose:
        choose = m_choose.group(1)
        
    return choose


def active_semesters(row: List[str], sem_cols: Dict[str, int]) -> List[str]:
    """선택군이 실제로 운영되는 모든 학기를 표의 개설 학점 열에서 찾습니다."""
    return [sem for sem, col_idx in sem_cols.items() if clean(row[col_idx])]


def main():
    parser = argparse.ArgumentParser(description="공주여고 HWPX 교육과정 편성표를 LLM용 CSV로 변환합니다.")
    parser.add_argument("input", type=Path, help="입력 HWPX 파일 경로")
    parser.add_argument("--outdir", type=Path, default=None, help="출력 폴더 (기본값: 입력 파일과 동일한 폴더)")
    parser.add_argument("--prefix", default="gjghs_2027_curriculum", help="출력 파일명 prefix. 기본값: gjghs_2027_curriculum")
    args = parser.parse_args()

    outdir = args.outdir if args.outdir else args.input.parent

    print(f"Loading HWPX: {args.input}")
    z = zipfile.ZipFile(args.input)
    tree = ET.parse(z.open("Contents/section0.xml"))
    root = tree.getroot()

    # 테이블 검색
    tbls = [e for e in root.iter() if e.tag.endswith('tbl')]
    if len(tbls) < 2:
        raise ValueError(f"정상적인 2개 이상의 테이블을 찾지 못했습니다. 발견 개수: {len(tbls)}")

    # Grid 파싱
    grid_designated = parse_hwpx_table(tbls[0]) # Table 0: 지정 및 일부 선택
    grid_selection = parse_hwpx_table(tbls[1])   # Table 1: 주로 학생 선택 및 창체/요약

    course_rows: List[Dict[str, str]] = []
    summary_rows: List[Dict[str, str]] = []
    
    # 학기 및 인덱스 매핑 정의 (Table 0/1 공통 그리드 구조 기준)
    # 0: 구분, 1: 교과(군), 2: 과목유형, 3: 과목(선택군), 4: 과목명
    # 5: 기본학점, 6: 편성학점
    # 7: 1-1, 8: 1-2, 9: 2-1, 10: 2-2, 11: 3-1, 12: 3-2
    # 13: 비고, 14: 편성학점 합
    SEMESTERS = ["1-1", "1-2", "2-1", "2-2", "3-1", "3-2"]
    SEM_COLS = {sem: 7 + idx for idx, sem in enumerate(SEMESTERS)}

    def is_summary_or_changement_row(row: List[str]) -> bool:
        joined = " ".join(row).strip()
        keywords = ["소계", "총계", "창의적", "동아리", "진로 활동", "자율·자치", "체험활동", "이수 학점"]
        return any(k in joined for k in keywords)

    # 1. 학교 지정 과목 파싱. 소계 행을 기준으로 범위를 찾아 개정안의 행 증감에 대응합니다.
    designated_end = next(
        r for r, row in enumerate(grid_designated)
        if "지정과목 이수 학점 소계" in " ".join(row)
    )
    for r in range(3, designated_end):
        row = grid_designated[r]
        subject = clean(row[3])
        if not subject:
            continue
            
        course_row = {
            "원본행": f"Table0_Row{r}",
            "교과군": clean(row[1]),
            "과목명": subject,
            "학교지정": "True",
            "학년선택": "",
            "과목유형": clean(row[2]),
            "비고_전문_고시외_표6": "",
            "학생선택_원문": "",
            "선택군_코드": "",
            "선택군_학년학기": "",
            "선택군_선택수": "",
            "선택군_과목당학점": "",
            "선택군_총학점": "",
            "선택군_행범위": "",
            "기본학점": clean(row[5]),
            "편성학점": clean(row[6]),
            "교과군별이수학점": "",
            "필수이수학점": "",
            "과학중점": "",
            "정보중점": "",
            "사회중점": "",
            "인문중점": "",
            "캠공개설여부": "",
            "비고": clean(row[13]),
        }
        for sem, col_idx in SEM_COLS.items():
            val = clean(row[col_idx])
            course_row[f"{sem}_원본"] = val
            course_row[f"{sem}_학점"] = val
            
        course_rows.append(course_row)

    # 2. 학생 선택 과목 파싱. 각 표의 합계 전까지를 자동으로 수집합니다.
    selection_sources = []
    for r in range(designated_end + 1, len(grid_designated)):
        row = grid_designated[r]
        if not is_summary_or_changement_row(row):
            selection_sources.append((f"Table0_Row{r}", row))
            
    # Table 1 선택과목 추가
    selection_end = next(
        r for r, row in enumerate(grid_selection)
        if "선택과목 이수학점 소계" in " ".join(row)
    )
    for r in range(0, selection_end):
        row = grid_selection[r]
        if not is_summary_or_changement_row(row):
            selection_sources.append((f"Table1_Row{r}", row))

    for origin_tag, row in selection_sources:
        subject = clean(row[4])
        if not subject:
            continue
            
        group_name = clean(row[1]) # e.g. "선택군1", "선택군3"
        raw_select = clean(row[13]) # e.g. "택2"
        credit = clean(row[6]) # 편성학점
        
        # 선택군 9처럼 여러 학기에 걸친 군도 누락하지 않습니다.
        active_sems = active_semesters(row, SEM_COLS)
        active_sem = "/".join(active_sems)
                
        choose_count = extract_selection_meta(raw_select)
        
        sel_code = ""
        total_credit = ""
        if group_name and active_sem:
            sel_code = f"{active_sem}-{group_name}-{raw_select}({credit})"
            # 표의 편성학점 합이 선택군 전체 학점을 가장 정확히 나타냅니다.
            total_credit = clean(row[14])

        course_row = {
            "원본행": origin_tag,
            "교과군": clean(row[1]),
            "과목명": subject,
            "학교지정": "",
            "학년선택": "True",
            "과목유형": clean(row[2]),
            "비고_전문_고시외_표6": "",
            "학생선택_원문": raw_select,
            "선택군_코드": sel_code,
            "선택군_학년학기": active_sem,
            "선택군_선택수": choose_count,
            "선택군_과목당학점": credit,
            "선택군_총학점": total_credit,
            "선택군_행범위": "", # 아래에서 별도 후처리로 채움
            "기본학점": clean(row[5]),
            "편성학점": credit,
            "교과군별이수학점": "",
            "필수이수학점": "",
            "과학중점": "",
            "정보중점": "",
            "사회중점": "",
            "인문중점": "",
            "캠공개설여부": "",
            "비고": "",
        }
        for sem, col_idx in SEM_COLS.items():
            val = clean(row[col_idx])
            course_row[f"{sem}_원본"] = val
            # 과목별 학기학점엔 전체 총이수학점이 아닌 과목당 학점 기록
            course_row[f"{sem}_학점"] = credit if val else ""
            
        course_rows.append(course_row)

    # 선택군의 행범위(row_range) 계산 및 주입
    # 같은 선택군 코드를 가진 과목들의 목록을 모으고 행범위를 구합니다.
    group_ranges: Dict[str, List[int]] = {}
    for idx, r in enumerate(course_rows):
        code = r["선택군_코드"]
        if code:
            if code not in group_ranges:
                group_ranges[code] = []
            group_ranges[code].append(idx)
            
    for code, idxs in group_ranges.items():
        # CSV의 행인덱스는 헤더 제외한 데이터 행 기준으로 1-indexed 또는 absolute mapping
        # 여기서는 리스트 내 인덱스로 행범위를 간단하게 기록합니다. (r_start:r_end)
        r_start = idxs[0] + 4 # 엑셀/HWP 테이블 줄 번호에 맞춤형 가산
        r_end = idxs[-1] + 4
        range_str = f"{r_start}:{r_end}"
        for idx in idxs:
            course_rows[idx]["선택군_행범위"] = range_str

    # 3. 하단 요약 및 창체 영역 파싱
    # Table 0 지정과목 소계와 Table 1의 선택과목 소계 이후 요약 행
    summary_sources = [
        (f"Table0_Row{designated_end}", grid_designated[designated_end]),
    ]
    for r in range(selection_end, len(grid_selection)):
        summary_sources.append((f"Table1_Row{r}", grid_selection[r]))
        
    for tag, row in summary_sources:
        # 1-1~3-2 semester mapping
        s_row = {
            "원본행": tag,
            "항목": clean(row[0]),
            "세부항목": clean(row[1]) if len(row) > 1 else "",
            "1-1": clean(row[7]) if len(row) > 7 else "",
            "1-2": clean(row[8]) if len(row) > 8 else "",
            "2-1": clean(row[9]) if len(row) > 9 else "",
            "2-2": clean(row[10]) if len(row) > 10 else "",
            "3-1": clean(row[11]) if len(row) > 11 else "",
            "3-2": clean(row[12]) if len(row) > 12 else "",
            "편성": clean(row[14]) if len(row) > 14 else (clean(row[13]) if len(row) > 13 else ""),
            "교과군별이수학점": "",
            "필수이수학점": "",
            "비고": "",
        }
        summary_rows.append(s_row)

    # 4. 학생 선택군 목록 빌드
    selection_groups = build_selection_groups(course_rows)

    # 5. CSV 작성
    outdir.mkdir(parents=True, exist_ok=True)
    courses_path = outdir / f"{args.prefix}_courses_llm.csv"
    groups_path = outdir / f"{args.prefix}_selection_groups_llm.csv"
    summary_path = outdir / f"{args.prefix}_summary_llm.csv"
    readme_path = outdir / f"{args.prefix}_readme.md"

    write_csv(courses_path, COURSE_FIELDNAMES, course_rows)
    write_csv(groups_path, SELECTION_FIELDNAMES, selection_groups)
    write_csv(summary_path, SUMMARY_FIELDNAMES, summary_rows)
    write_readme(readme_path, args.input.name, "Contents/section0.xml", len(course_rows), len(selection_groups), len(summary_rows))

    print("변환이 성공적으로 완료되었습니다!")
    print(f"- 과목: {courses_path}")
    print(f"- 선택군: {groups_path}")
    print(f"- 요약/창체: {summary_path}")
    print(f"- 안내문: {readme_path}")


def build_selection_groups(course_rows: List[Dict[str, str]]) -> List[Dict[str, str]]:
    groups: Dict[str, SelectionGroup] = {}
    for idx, row in enumerate(course_rows):
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
        
    return [groups[k].as_row() for k in sorted(groups.keys())]


def write_csv(path: Path, fieldnames: List[str], rows: Iterable[Dict[str, str]]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def write_readme(path: Path, source_file: str, sheet: str, course_count: int, group_count: int, summary_count: int) -> None:
    text = f"""# 교육과정 편성표 HWPX -> LLM용 CSV 해석 규칙

## 원본
- 파일: `{source_file}`
- 경로: `{sheet}`

## 생성 파일
- `*_courses_llm.csv`: 과목 1개를 1행으로 정리한 평면표
- `*_selection_groups_llm.csv`: 학생 선택군 단위 요약표
- `*_summary_llm.csv`: 하단 합계/창의적 체험활동 등 요약 영역
- `*_readme.md`: 해석 규칙 설명 (본 파일)

## 변환 원칙
1. HWPX 내부의 XML 좌표 속성(`cellAddr`, `cellSpan`)을 사용하여 모든 병합 영역을 완벽하게 좌표계에 맵핑 및 복원했다.
2. `1-1` ~ `3-2` 의 학기는 해당 학기 편성 학점이다.
3. 선택군에 속한 개별 과목의 학점 컬럼에는 해당 선택군의 과목당 학점(예: 3 또는 4)을 기록했다.

## 생성 결과 개수
- 과목 행: {course_count}
- 선택군: {group_count}
- 요약 행: {summary_count}
"""
    path.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
