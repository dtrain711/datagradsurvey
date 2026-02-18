#!/usr/bin/env python3
"""Analyze exit survey Excel workbook and rank courses by student preference."""

from __future__ import annotations

import argparse
import csv
import json
import re
import textwrap
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
import xml.etree.ElementTree as ET

NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


@dataclass
class CourseStats:
    most_beneficial: int = 0
    neutral: int = 0
    least_beneficial: int = 0
    did_not_take: int = 0

    @property
    def considered(self) -> int:
        return self.most_beneficial + self.neutral + self.least_beneficial

    @property
    def weighted_score(self) -> int:
        return (2 * self.most_beneficial) + self.neutral - (2 * self.least_beneficial)

    @property
    def favorability_pct(self) -> float:
        if self.considered == 0:
            return 0.0
        return 100 * (self.most_beneficial + self.neutral) / self.considered


def _col_num(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - 64)
    return value


def read_workbook_rows(xlsx_path: Path) -> tuple[list[str], list[dict[str, str]], str]:
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("a:si", NS):
                chunks = [t.text or "" for t in si.findall(".//a:t", NS)]
                shared_strings.append("".join(chunks))

        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        first_sheet = workbook.find("a:sheets/a:sheet", NS)
        if first_sheet is None:
            raise ValueError("Workbook has no sheets")
        sheet_name = first_sheet.attrib["name"]
        rel_id = first_sheet.attrib[f"{{{NS['r']}}}id"]

        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        target = None
        for rel in rels:
            if rel.attrib.get("Id") == rel_id:
                target = rel.attrib["Target"]
                break
        if not target:
            raise ValueError(f"Could not resolve sheet relationship {rel_id}")
        sheet_path = target if target.startswith("xl/") else f"xl/{target}"

        sheet_xml = ET.fromstring(zf.read(sheet_path))

    grid: dict[int, dict[int, str]] = defaultdict(dict)
    for row in sheet_xml.findall(".//a:sheetData/a:row", NS):
        row_idx = int(row.attrib["r"])
        for c in row.findall("a:c", NS):
            cell_ref = c.attrib.get("r", "")
            col_idx = _col_num(cell_ref)
            cell_type = c.attrib.get("t")
            value = ""
            v = c.find("a:v", NS)

            if cell_type == "s" and v is not None and v.text is not None:
                value = shared_strings[int(v.text)]
            elif cell_type == "inlineStr":
                t = c.find("a:is/a:t", NS)
                value = t.text if t is not None and t.text is not None else ""
            elif v is not None and v.text is not None:
                value = v.text

            grid[row_idx][col_idx] = value.strip()

    if 1 not in grid:
        raise ValueError("Missing header row")

    max_col = max(grid[1].keys())
    headers = [grid[1].get(col, "") for col in range(1, max_col + 1)]

    rows: list[dict[str, str]] = []
    for row_idx in sorted(k for k in grid.keys() if k > 1):
        data = {}
        has_data = False
        for col in range(1, max_col + 1):
            header = headers[col - 1]
            if not header:
                continue
            value = grid[row_idx].get(col, "")
            if value:
                has_data = True
            data[header] = value
        if has_data:
            rows.append(data)

    return headers, rows, sheet_name


def split_courses(cell_value: str) -> list[str]:
    if not cell_value:
        return []
    return [c.strip() for c in cell_value.split(",") if c.strip()]


def build_ranking(headers: list[str], rows: list[dict[str, str]]) -> tuple[list[tuple[str, CourseStats]], dict[str, str]]:
    group_columns = {}
    for h in headers:
        normalized = " ".join(h.split())
        if "Groups - Most Beneficial" in normalized:
            group_columns["most_beneficial"] = h
        elif "Groups - Neutral" in normalized:
            group_columns["neutral"] = h
        elif "Groups - Least Beneficial" in normalized:
            group_columns["least_beneficial"] = h
        elif "Groups - Did not take" in normalized:
            group_columns["did_not_take"] = h

    required = {"most_beneficial", "neutral", "least_beneficial", "did_not_take"}
    missing = required.difference(group_columns)
    if missing:
        raise ValueError(f"Missing expected columns: {', '.join(sorted(missing))}")

    stats: dict[str, CourseStats] = defaultdict(CourseStats)

    for row in rows:
        for category, col_name in group_columns.items():
            courses = split_courses(row.get(col_name, ""))
            for course in courses:
                current = stats[course]
                setattr(current, category, getattr(current, category) + 1)

    ranking = sorted(
        stats.items(),
        key=lambda item: (item[1].weighted_score, item[1].most_beneficial, item[1].favorability_pct),
        reverse=True,
    )
    return ranking, group_columns


def write_csv(path: Path, ranking: list[tuple[str, CourseStats]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(
            [
                "rank",
                "course",
                "weighted_score",
                "most_beneficial",
                "neutral",
                "least_beneficial",
                "did_not_take",
                "considered",
                "favorability_pct",
            ]
        )
        for idx, (course, s) in enumerate(ranking, start=1):
            writer.writerow(
                [
                    idx,
                    course,
                    s.weighted_score,
                    s.most_beneficial,
                    s.neutral,
                    s.least_beneficial,
                    s.did_not_take,
                    s.considered,
                    f"{s.favorability_pct:.2f}",
                ]
            )


def _escape_xml(text: str) -> str:
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def write_svg(path: Path, ranking: list[tuple[str, CourseStats]], year: str, top_n: int = 10) -> None:
    top = ranking[:top_n]
    if not top:
        path.write_text("<svg xmlns='http://www.w3.org/2000/svg' width='800' height='120'></svg>", encoding="utf-8")
        return

    width = 1000
    row_height = 48
    chart_left = 370
    chart_right = 930
    top_margin = 70
    max_score = max(s.weighted_score for _, s in top)
    min_score = min(s.weighted_score for _, s in top)
    scale_span = max(max_score - min(0, min_score), 1)

    height = top_margin + row_height * len(top) + 40
    lines = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}' viewBox='0 0 {width} {height}'>",
        "<style>text{font-family:Arial,Helvetica,sans-serif;} .title{font-size:24px;font-weight:700;} .label{font-size:15px;} .score{font-size:14px;font-weight:700;} .subtitle{font-size:13px;fill:#444;} </style>",
        f"<text x='40' y='36' class='title'>Top {len(top)} Ranked Courses ({_escape_xml(year)})</text>",
        "<text x='40' y='56' class='subtitle'>Weighted score = (2 × most beneficial) + neutral − (2 × least beneficial)</text>",
    ]

    axis_x = chart_left
    lines.append(f"<line x1='{axis_x}' y1='{top_margin-20}' x2='{axis_x}' y2='{height-20}' stroke='#666' stroke-width='1'/>")

    for idx, (course, s) in enumerate(top):
        y = top_margin + idx * row_height
        usable = chart_right - chart_left
        bar_width = int((s.weighted_score / scale_span) * usable)
        x = chart_left
        color = "#2a6fdb" if s.weighted_score >= 0 else "#d9534f"
        wrapped = textwrap.wrap(course, width=44)[:2]
        lines.append(f"<rect x='{x}' y='{y-14}' width='{max(bar_width,1)}' height='24' fill='{color}' rx='4' ry='4' />")
        lines.append(f"<text x='40' y='{y-2}' class='label'>{_escape_xml(wrapped[0])}</text>")
        if len(wrapped) > 1:
            lines.append(f"<text x='40' y='{y+14}' class='label'>{_escape_xml(wrapped[1])}</text>")
        lines.append(f"<text x='{x + max(bar_width,1) + 8}' y='{y+2}' class='score'>{s.weighted_score}</text>")

    lines.append("</svg>")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_summary(path: Path, ranking: list[tuple[str, CourseStats]], year: str, sheet_name: str, response_count: int) -> None:
    top5 = ranking[:5]
    lines = [
        f"# {year} Exit Survey Course Ranking",
        "",
        f"Source sheet: `{sheet_name}`",
        f"Responses analyzed: **{response_count}**",
        "",
        "Scoring rule: `weighted_score = (2 × most_beneficial) + neutral − (2 × least_beneficial)`",
        "",
        "## Top 5 courses",
        "",
        "| Rank | Course | Weighted score | Most beneficial | Neutral | Least beneficial | Favorability |",
        "|---:|---|---:|---:|---:|---:|---:|",
    ]
    for idx, (course, s) in enumerate(top5, start=1):
        lines.append(
            f"| {idx} | {course} | {s.weighted_score} | {s.most_beneficial} | {s.neutral} | {s.least_beneficial} | {s.favorability_pct:.1f}% |"
        )

    lines.extend([
        "",
        "## Figure",
        "",
        "![Course ranking chart](ranking.svg)",
    ])

    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", default="2023 Grad Program Exit Survey Data Project One.xlsx", help="Input Excel workbook path")
    parser.add_argument("--output-dir", default="outputs/2023", help="Output directory")
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    year_match = re.search(r"(20\d{2})", input_path.name)
    year = year_match.group(1) if year_match else "Unknown Year"

    headers, rows, sheet_name = read_workbook_rows(input_path)

    filtered_rows = [
        r for r in rows
        if r.get("Response ID", "").startswith("R_") and r.get("Finished", "") in {"1", "True", "TRUE"}
    ]

    ranking, group_columns = build_ranking(headers, filtered_rows)

    write_csv(output_dir / "course_ranking.csv", ranking)
    write_svg(output_dir / "ranking.svg", ranking, year=year, top_n=min(10, len(ranking)))
    write_summary(output_dir / "README.md", ranking, year=year, sheet_name=sheet_name, response_count=len(filtered_rows))

    metadata = {
        "input_file": str(input_path),
        "sheet_name": sheet_name,
        "year": year,
        "responses_analyzed": len(filtered_rows),
        "group_columns": group_columns,
        "courses_ranked": len(ranking),
    }
    (output_dir / "metadata.json").write_text(json.dumps(metadata, indent=2) + "\n", encoding="utf-8")

    if ranking:
        top_course, top_stats = ranking[0]
        print(f"Top ranked course: {top_course} (weighted score: {top_stats.weighted_score})")
    print(f"Wrote outputs to: {output_dir}")


if __name__ == "__main__":
    main()
