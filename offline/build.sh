#!/usr/bin/env bash
# 오프라인 점검 도구를 HTML 파일 1개로 합친다. 사용: ./offline/build.sh 1.0
set -euo pipefail
cd "$(dirname "$0")"
VER="${1:?버전을 넣어 주세요 (예: 1.0)}"
OUT="dist/법인카드점검_v${VER}.html"
mkdir -p dist
{
  sed "s/오프라인 v[0-9.]*</오프라인 v${VER}</; s/<br>v[0-9.]* · /<br>v${VER} · /" src/head.html
  printf '<script>\n'; cat src/vendor/xlsx.full.min.js; printf '\n</script>\n<script>\n'
  cat src/ref.js; grep -v '^if (typeof module' src/engine.js
  printf '</script>\n<script>\n'; cat src/app.js; printf '</script>\n</body>\n</html>\n'
} > "$OUT"
if grep -qiE '<script[^>]+src=|<link[^>]+href=' "$OUT"; then echo "외부 파일 참조가 있습니다: $OUT" >&2; exit 1; fi
echo "만들었습니다: offline/$OUT ($(wc -c < "$OUT") bytes)"
