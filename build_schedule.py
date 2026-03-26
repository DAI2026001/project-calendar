#!/usr/bin/env python3
import datetime as dt
import json
import sys
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
BLUE_FILL_HEX = "FF8CDDFA"
FUZHU_SOURCE = Path("/Users/sdai/Downloads/富竹项目.xlsx")


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iterfind(".//a:t", NS))
        for item in root.findall("a:si", NS)
    ]


def read_workbook_sheets(archive: zipfile.ZipFile) -> list[tuple[str, ET.Element]]:
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = workbook.findall("a:sheets/a:sheet", NS)
    if not sheets:
        raise ValueError("Excel 中没有可读取的工作表")

    parsed = []
    for sheet in sheets:
        rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = "xl/" + rel_map[rid]
        parsed.append((sheet.attrib.get("name", "未命名工作表"), ET.fromstring(archive.read(target))))
    return parsed


def get_blue_style_ids(archive: zipfile.ZipFile) -> set[int]:
    styles = ET.fromstring(archive.read("xl/styles.xml"))
    fills = styles.find("a:fills", NS)
    cell_xfs = styles.find("a:cellXfs", NS)
    if fills is None or cell_xfs is None:
        return set()

    fill_list = fills.findall("a:fill", NS)
    blue_styles = set()
    for index, xf in enumerate(cell_xfs.findall("a:xf", NS)):
        fill_id = int(xf.attrib.get("fillId", "0"))
        if fill_id >= len(fill_list):
            continue
        fill_text = ET.tostring(fill_list[fill_id], encoding="unicode")
        if BLUE_FILL_HEX in fill_text:
            blue_styles.add(index)
    return blue_styles


def cell_text(shared_strings: list[str], cell: ET.Element) -> str:
    value_node = cell.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v")
    value = "" if value_node is None else value_node.text or ""
    if cell.attrib.get("t") == "s" and value:
        return shared_strings[int(value)]
    return value


def normalize_number(raw: str) -> str:
    if raw in ("", None):
        return ""

    try:
        number = float(raw)
    except (TypeError, ValueError):
        return str(raw).strip()

    if number.is_integer():
        return str(int(number))
    return f"{number:.2f}".rstrip("0").rstrip(".")


def normalize_date(raw: str) -> dict[str, str]:
    if raw in ("", None):
        return {
            "datetime": "9999-12-31T23:59",
            "dateLabel": "待定",
            "dayKey": "待定",
            "timeLabel": "待定",
        }

    text = str(raw).strip()
    try:
        date = dt.datetime(1899, 12, 30) + dt.timedelta(days=float(text))
        return {
            "datetime": date.strftime("%Y-%m-%dT%H:%M"),
            "dateLabel": date.strftime("%Y-%m-%d %H:%M"),
            "dayKey": date.strftime("%Y-%m-%d"),
            "timeLabel": date.strftime("%H:%M"),
        }
    except (TypeError, ValueError):
        cleaned = text.replace(" ", "")
        time_label = "待定"
        if ":" in cleaned:
            time_label = cleaned[-8:] if len(cleaned) >= 8 else cleaned
        return {
            "datetime": "9999-12-31T23:59",
            "dateLabel": text,
            "dayKey": "待定",
            "timeLabel": time_label,
        }


def normalize_sheet_tag(sheet_name: str) -> str:
    if "水利" in sheet_name:
        return "水利"
    if "市政" in sheet_name or "房建" in sheet_name:
        return "市政房建"
    return sheet_name


def parse_sheet(path: Path, workbook_tag: str = "辉阳") -> list[dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        sheets = read_workbook_sheets(archive)
        blue_style_ids = get_blue_style_ids(archive)

    projects = []
    for sheet_name, sheet in sheets:
        rows = sheet.findall(".//a:sheetData/a:row", NS)
        sheet_tag = normalize_sheet_tag(sheet_name)
        for row in rows[1:]:
            cells = row.findall("a:c", NS)
            if not cells:
                continue

            style_ids = [int(cell.attrib.get("s", "0")) for cell in cells]
            if not any(style_id in blue_style_ids for style_id in style_ids):
                continue

            row_map = {}
            for cell in cells:
                ref = cell.attrib.get("r", "")
                column = "".join(ch for ch in ref if ch.isalpha())
                row_map[column] = cell_text(shared_strings, cell)

            if not row_map.get("C"):
                continue

            date_info = normalize_date(row_map.get("D"))
            amount = normalize_number(row_map.get("E"))
            amount_wan = ""
            if amount:
                try:
                    amount_wan = f"{float(amount) / 10000:.2f}".rstrip("0").rstrip(".")
                except ValueError:
                    amount_wan = ""

            projects.append(
                {
                    "id": normalize_number(row_map.get("A")),
                    "workbookTag": workbook_tag,
                    "sheetName": sheet_name,
                    "sheetTag": sheet_tag,
                    "region": (row_map.get("B") or "").strip(),
                    "name": (row_map.get("C") or "").replace("\u200d", "").strip(),
                    "datetime": date_info["datetime"],
                    "dateLabel": date_info["dateLabel"],
                    "dayKey": date_info["dayKey"],
                    "timeLabel": date_info["timeLabel"],
                    "amount": amount,
                    "amountWan": amount_wan,
                    "qualification": (row_map.get("F") or "").strip(),
                    "evaluation": (row_map.get("G") or "").strip(),
                    "deposit": normalize_number(row_map.get("H")),
                    "staff": (row_map.get("I") or "").strip(),
                    "kValue": (row_map.get("J") or "").strip(),
                    "billCount": normalize_number(row_map.get("K")),
                    "danger": (row_map.get("L") or "").strip(),
                    "confinedSpace": (row_map.get("M") or "").strip(),
                    "signup": ((row_map.get("N") or "") or (row_map.get("O") or "")).strip(),
                }
            )

    projects.sort(key=lambda item: (item["datetime"], item["region"], item["name"]))
    return projects


def project_merge_key(item: dict[str, str]) -> str:
    return item["name"].strip()


def merge_text(left: str, right: str) -> str:
    return left if left not in ("", None) else right


def merge_project_records(base: dict[str, str], extra: dict[str, str]) -> dict[str, str]:
    merged = dict(base)
    for key in [
        "region",
        "datetime",
        "dateLabel",
        "dayKey",
        "timeLabel",
        "amount",
        "amountWan",
        "qualification",
        "evaluation",
        "deposit",
        "staff",
        "kValue",
        "billCount",
        "danger",
        "confinedSpace",
        "signup",
    ]:
        merged[key] = merge_text(merged.get(key, ""), extra.get(key, ""))
    merged["id"] = merge_text(merged.get("id", ""), extra.get("id", ""))
    return merged


def combine_sources(primary_projects: list[dict[str, str]], fuzhu_projects: list[dict[str, str]]) -> list[dict[str, str]]:
    water_projects = [item for item in primary_projects if item["sheetTag"] == "水利"]
    huiyang_projects = [item for item in primary_projects if item["sheetTag"] == "市政房建"]

    merged_by_key: dict[str, dict[str, str]] = {}
    for item in huiyang_projects:
        record = dict(item)
        record["sourceTag"] = "辉阳"
        merged_by_key[project_merge_key(item)] = record

    for item in fuzhu_projects:
        key = project_merge_key(item)
        if key in merged_by_key:
            merged = merge_project_records(merged_by_key[key], item)
            merged["sourceTag"] = "两家"
            merged["sheetTag"] = "市政房建"
            merged["sheetName"] = "投标项目市政房建 / 富竹项目"
            merged["workbookTag"] = "辉阳,富竹"
            merged_by_key[key] = merged
        else:
            record = dict(item)
            record["sourceTag"] = "富竹"
            merged_by_key[key] = record

    for item in water_projects:
        record = dict(item)
        record["sourceTag"] = "水利"
        merged_by_key[f"water::{project_merge_key(item)}::{item.get('datetime', '')}"] = record

    combined = list(merged_by_key.values())
    combined.sort(key=lambda item: (item["datetime"], item["sourceTag"], item["region"], item["name"]))
    return combined


def main() -> int:
    if len(sys.argv) != 3:
        print("用法: python3 build_schedule.py <辉阳xlsx路径> <输出js路径>")
        return 1

    source = Path(sys.argv[1]).expanduser()
    output = Path(sys.argv[2]).expanduser()

    projects = parse_sheet(source, workbook_tag="辉阳")
    if FUZHU_SOURCE.exists():
        fuzhu_projects = parse_sheet(FUZHU_SOURCE, workbook_tag="富竹")
        projects = combine_sources(projects, fuzhu_projects)
    payload = "window.SCHEDULE_DATA = " + json.dumps(
        projects,
        ensure_ascii=False,
        separators=(",", ":"),
    ) + ";\n"
    output.write_text(payload, encoding="utf-8")
    print(f"已生成 {output} ，共 {len(projects)} 条项目数据")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
