#!/usr/bin/env bash
# Full suite. UI and OOXML need jsdom (devDependency or JSDOM_PATH or /tmp/dom).
# App tests need the server on 127.0.0.1:4173. CI=true refuses to skip either.
set -euo pipefail
cd "$(dirname "$0")"

echo "== enterprise =="
node enterprise.test.mjs
echo "== formula =="
node formula.test.mjs

have_jsdom=0
if node --input-type=module -e "import { loadJsdom } from './load-jsdom.mjs'; await loadJsdom();"; then
  have_jsdom=1
fi

if [ "$have_jsdom" != 1 ]; then
  if [ "${CI:-}" = "true" ]; then
    echo "jsdom is required in CI (npm ci, or set JSDOM_PATH)" >&2
    exit 1
  fi
  echo "skipping ooxml/app/offline — jsdom not installed"
  exit 0
fi

echo "== ooxml =="
node ooxml.test.mjs

echo "== ooxml python =="
if python3 -c "import openpyxl, docx" >/dev/null 2>&1; then
  python3 validate_ooxml.py
elif [ "${CI:-}" = "true" ]; then
  echo "openpyxl and python-docx are required in CI" >&2
  exit 1
else
  echo "skipping python OOXML validators (pip install openpyxl python-docx)"
fi

echo "== app (server mode) =="
if curl -sf http://127.0.0.1:4173/healthz >/dev/null 2>&1; then
  node app.test.mjs
elif [ "${CI:-}" = "true" ]; then
  echo "app tests require the server on 127.0.0.1:4173" >&2
  exit 1
else
  echo "skipping app tests — start the server with: node server.js"
fi

echo "== offline =="
node offline.test.mjs
echo "ALL SUITES PASSED"
