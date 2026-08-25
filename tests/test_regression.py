"""검수된 학교별 CSV를 기준으로 파서 회귀를 검사합니다."""

from __future__ import annotations

import csv
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def read_csv_canonically(path: Path) -> list[dict[str, str]]:
    """운영체제별 셀 내부 줄바꿈 차이를 제외하고 CSV를 읽습니다."""
    with path.open(encoding="utf-8-sig", newline="") as file:
        return [
            {key: value.replace("\r\n", "\n") for key, value in row.items()}
            for row in csv.DictReader(file)
        ]


class ParserRegressionTests(unittest.TestCase):
    maxDiff = None

    def run_parser(self, *arguments: str) -> Path:
        output_dir = Path(tempfile.mkdtemp(prefix="curriculum-parser-test-"))
        self.addCleanup(shutil.rmtree, output_dir, ignore_errors=True)
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        subprocess.run(
            [sys.executable, *arguments, "--outdir", str(output_dir)],
            cwd=ROOT,
            env=env,
            check=True,
            capture_output=True,
            text=True,
            encoding="utf-8",
        )
        return output_dir

    def assert_generated_csvs_match(
        self,
        output_dir: Path,
        expected_dir: Path,
        prefix: str,
    ) -> None:
        for suffix in ("courses_llm.csv", "selection_groups_llm.csv", "summary_llm.csv"):
            name = f"{prefix}_{suffix}"
            self.assertEqual(
                read_csv_canonically(expected_dir / name),
                read_csv_canonically(output_dir / name),
                name,
            )
        readme_name = f"{prefix}_readme.md"
        self.assertEqual(
            (expected_dir / readme_name).read_text(encoding="utf-8").replace("\r\n", "\n"),
            (output_dir / readme_name).read_text(encoding="utf-8").replace("\r\n", "\n"),
            readme_name,
        )

    def test_xlsx_parser_reproduces_bangok_reviewed_csvs(self) -> None:
        prefix = "bangok_2026_curriculum"
        output_dir = self.run_parser(
            "convert_curriculum_to_llm_csv.py",
            "반곡고/2026학년도 반곡고등학교 입학생 3개년 교육과정 편성표.xlsx",
            "--sheet",
            "반곡고",
            "--prefix",
            prefix,
        )
        self.assert_generated_csvs_match(output_dir, ROOT / "반곡고", prefix)

    def test_hwpx_parser_reproduces_gongju_reviewed_csvs(self) -> None:
        prefix = "gjghs_2027_curriculum"
        output_dir = self.run_parser(
            "convert_hwpx_curriculum_to_llm_csv.py",
            "공주여고/2027학년도 교육과정 편제(수정안)-변경 후.hwpx",
            "--prefix",
            prefix,
        )
        self.assert_generated_csvs_match(output_dir, ROOT / "공주여고", prefix)

    def test_xlsx_parser_reproduces_woongsang_reviewed_csvs(self) -> None:
        prefix = "woongsang_2026_curriculum"
        output_dir = self.run_parser(
            "convert_curriculum_to_llm_csv.py",
            "woongsang/2026학년도 입학생 3개년 교육과정 편성 수정(안)_웅상고.xlsx",
            "--sheet",
            "Sheet1",
            "--prefix",
            prefix,
            "--overrides",
            "woongsang/course_overrides.csv",
        )
        expected_dir = ROOT / "woongsang"
        self.assert_generated_csvs_match(output_dir, expected_dir, prefix)

        courses = read_csv_canonically(output_dir / f"{prefix}_courses_llm.csv")
        self.assertEqual(94, len(courses))
        self.assertEqual(94, len({row["과목명"] for row in courses}))
        self.assertIn("공통국어1", {row["과목명"] for row in courses})
        self.assertFalse(any("#" in value for row in courses for value in row.values()))

        groups = read_csv_canonically(output_dir / f"{prefix}_selection_groups_llm.csv")
        semester_totals: dict[str, int] = {}
        for group in groups:
            semester = group["선택군_학년학기"]
            semester_totals[semester] = semester_totals.get(semester, 0) + int(group["선택군_총학점"])
        self.assertEqual({"2-1": 15, "2-2": 15, "3-1": 14, "3-2": 26}, semester_totals)

    def test_web_school_config_points_to_existing_csvs(self) -> None:
        html = (ROOT / "index.html").read_text(encoding="utf-8")
        expected_paths = (
            "반곡고/bangok_2026_curriculum_courses_llm.csv",
            "반곡고/bangok_2026_curriculum_summary_llm.csv",
            "공주여고/gjghs_2027_curriculum_courses_llm.csv",
            "공주여고/gjghs_2027_curriculum_summary_llm.csv",
            "woongsang/woongsang_2026_curriculum_courses_llm.csv",
            "woongsang/woongsang_2026_curriculum_summary_llm.csv",
        )
        for relative_path in expected_paths:
            self.assertIn(relative_path, html)
            self.assertTrue((ROOT / relative_path).is_file(), relative_path)


if __name__ == "__main__":
    unittest.main()
