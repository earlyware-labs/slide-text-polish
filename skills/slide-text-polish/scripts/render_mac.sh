#!/bin/bash
# Mac 版 PowerPoint で PDF に書き出し、ページを PNG にする（見た目は実機、改行の数値判定は推定）
# 使い方: bash render_mac.sh deck.pptx 出力フォルダ   ※ pdftoppm は brew install poppler
set -euo pipefail
src="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
out="$2"; mkdir -p "$out"; out="$(cd "$out" && pwd)"
pdf="$out/deck.pdf"
osascript <<EOF
tell application "Microsoft PowerPoint"
  open POSIX file "$src"
  save active presentation in POSIX file "$pdf" as save as PDF
  close active presentation saving no
end tell
EOF
pdftoppm -r 80 -png "$pdf" "$out/slide"
echo "書き出し完了: $out"
