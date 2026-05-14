#!/usr/bin/env sh
set -eu

PORT="${PORT:-8080}"

if command -v npx >/dev/null 2>&1; then
  npx --yes serve . --listen "$PORT"
else
  python3 -m http.server "$PORT"
fi
