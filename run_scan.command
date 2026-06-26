#!/bin/zsh
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
python3 run_daily_scan.py --refresh --open-urls
STATUS=$?
echo ""
if [ "$STATUS" -eq 0 ]; then
  echo "✅ 扫描完成，3秒后自动关闭终端..."
  sleep 3
  osascript -e 'tell application "Terminal" to close front window' 2>/dev/null
else
  echo "❌ 扫描失败，退出码：$STATUS。按任意键关闭..."
  read -k1
fi
exit "$STATUS"
