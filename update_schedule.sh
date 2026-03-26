#!/bin/zsh
set -e

SOURCE_FILE="/Users/sdai/Downloads/辉阳集团项目.xlsx"
OUTPUT_FILE="/Users/sdai/document/代码/condex/schedule-data.js"
SCRIPT_FILE="/Users/sdai/document/代码/condex/build_schedule.py"

if [[ ! -f "$SOURCE_FILE" ]]; then
  echo "未找到 Excel 文件：$SOURCE_FILE"
  echo "请先从腾讯文档导出并覆盖该文件。"
  exit 1
fi

python3 "$SCRIPT_FILE" "$SOURCE_FILE" "$OUTPUT_FILE"
echo ""
echo "数据已更新完成。"
echo "如需发布到网站，再执行："
echo "cd /Users/sdai/document/代码/condex && git add schedule-data.js && git commit -m 'Update bid schedule data' && git push"
