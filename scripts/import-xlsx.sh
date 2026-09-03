#!/usr/bin/env bash
set -euo pipefail
if [[ $# -ne 1 ]]; then
  echo "用法: npm run source:import -- /path/to/file-or-directory" >&2
  exit 2
fi
mkdir -p work
python3 scripts/xlsx_to_sql.py "$1" work/source-import.sql
WRANGLER_LOG_PATH=.wrangler/logs npx wrangler d1 execute bulletin-source-local --local --config wrangler.source.json --file work/source-import.sql > work/source-import.log
tail -n 8 work/source-import.log
rm -f work/source-import.sql
rm -f work/source-import.log
