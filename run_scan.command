#!/bin/zsh
set -u

SOURCE="$0"
while [ -L "$SOURCE" ]; do
  DIR="$(cd -P "$(dirname "$SOURCE")" && pwd)"
  TARGET="$(readlink "$SOURCE")"
  if [[ "$TARGET" == /* ]]; then
    SOURCE="$TARGET"
  else
    SOURCE="$DIR/$TARGET"
  fi
done

cd -P "$(dirname "$SOURCE")"

echo "实习僧上海校招新增岗位扫描"
echo "项目目录：$(pwd)"
echo "开始时间：$(date '+%Y-%m-%d %H:%M:%S')"
echo ""

python3 run_daily_scan.py
STATUS=$?
echo ""
if [ "$STATUS" -eq 0 ]; then
  echo "扫描完成。"
  echo "结果已写入 outputs 目录；最新 Excel 副本为：outputs/上海校招新增岗位筛选_latest.xlsx"
  echo ""
  echo "10秒后自动关闭终端；按 Ctrl-C 可取消关闭。"
  sleep 10
  osascript -e 'tell application "Terminal" to close front window' >/dev/null 2>&1
else
  echo "扫描失败，退出码：$STATUS。按任意键关闭..."
  read -k1
fi
exit "$STATUS"
