#!/bin/zsh
set -e

SOURCE_FILE="/Users/sdai/Downloads/辉阳集团项目.xlsx"
OUTPUT_FILE="/Users/sdai/document/代码/condex/schedule-data.js"
SCRIPT_FILE="/Users/sdai/document/代码/condex/build_schedule.py"
SINGLE_FILE_SCRIPT="/Users/sdai/document/代码/condex/build_single_file.py"
SINGLE_FILE_OUTPUT="/Users/sdai/document/代码/condex/项目日程表.html"
PUBLIC_SINGLE_FILE_OUTPUT="/Users/sdai/document/代码/condex/schedule.html"

if [[ ! -f "$SOURCE_FILE" ]]; then
  echo "未找到 Excel 文件：$SOURCE_FILE"
  echo "请先从腾讯文档导出并覆盖该文件。"
  exit 1
fi

python3 "$SCRIPT_FILE" "$SOURCE_FILE" "$OUTPUT_FILE"
python3 "$SINGLE_FILE_SCRIPT" "/Users/sdai/document/代码/condex/index.html" "$OUTPUT_FILE" "$SINGLE_FILE_OUTPUT"
python3 "$SINGLE_FILE_SCRIPT" "/Users/sdai/document/代码/condex/index.html" "$OUTPUT_FILE" "$PUBLIC_SINGLE_FILE_OUTPUT"
echo ""
echo "数据已更新完成。"
echo "单文件 H5：$SINGLE_FILE_OUTPUT"
echo "公网部署版：$PUBLIC_SINGLE_FILE_OUTPUT"
echo "如需发布到网站，再执行："
echo "cd /Users/sdai/document/代码/condex && git add schedule-data.js 项目日程表.html schedule.html && git commit -m 'Update bid schedule data' && git push"
