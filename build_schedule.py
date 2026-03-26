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


def read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iterfind(".//a:t", NS))
        for item in root.findall("a:si", NS)
    ]


def read_first_sheet(archive: zipfile.ZipFile) -> ET.Element:
    rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}

    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheet = workbook.find("a:sheets/a:sheet", NS)
    if sheet is None:
        raise ValueError("Excel 中没有可读取的工作表")

    rid = sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
    target = "xl/" + rel_map[rid]
    return ET.fromstring(archive.read(target))


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


def parse_sheet(path: Path) -> list[dict[str, str]]:
    with zipfile.ZipFile(path) as archive:
        shared_strings = read_shared_strings(archive)
        sheet = read_first_sheet(archive)
        rows = sheet.findall(".//a:sheetData/a:row", NS)

    projects = []
    for row in rows[1:]:
        row_map = {}
        for cell in row.findall("a:c", NS):
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


def main() -> int:
    if len(sys.argv) != 3:
        print("用法: python3 build_schedule.py <xlsx路径> <输出js路径>")
        return 1

    source = Path(sys.argv[1]).expanduser()
    output = Path(sys.argv[2]).expanduser()

    projects = parse_sheet(source)
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
