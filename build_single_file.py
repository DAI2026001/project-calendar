#!/usr/bin/env python3
from pathlib import Path
import sys


SCRIPT_TAG = '<script src="./schedule-data.js"></script>'


def main() -> int:
    if len(sys.argv) != 4:
        print("用法: python3 build_single_file.py <index.html> <schedule-data.js> <输出html>")
        return 1

    index_path = Path(sys.argv[1]).expanduser()
    data_path = Path(sys.argv[2]).expanduser()
    output_path = Path(sys.argv[3]).expanduser()

    html = index_path.read_text(encoding="utf-8")
    data_js = data_path.read_text(encoding="utf-8").strip()

    if SCRIPT_TAG not in html:
        raise ValueError("index.html 中未找到 schedule-data.js 的脚本引用")

    standalone_html = html.replace(SCRIPT_TAG, f"<script>\n{data_js}\n</script>", 1)
    output_path.write_text(standalone_html, encoding="utf-8")
    print(f"已生成单文件 H5：{output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
