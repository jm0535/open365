#!/usr/bin/env bash
# Parse every first-party script and refuse runtime npm dependencies.
set -euo pipefail
cd "$(dirname "$0")/.."

node --check server.js

while IFS= read -r -d '' f; do
  node --check "$f"
done < <(find public -name '*.js' -print0)

while IFS= read -r -d '' f; do
  node --check "$f"
done < <(find test scripts -name '*.mjs' -print0)

node --input-type=module -e "import './test/load-jsdom.mjs'"

node -e '
const p = require("./package.json");
const deps = Object.keys(p.dependencies || {});
if (deps.length) {
  console.error("runtime dependencies must stay empty, found:", deps.join(", "));
  process.exit(1);
}
if (!p.scripts || !p.scripts.ci || !p.scripts.test) {
  console.error("package.json is missing ci/test scripts");
  process.exit(1);
}
'

echo "syntax ok ($(find public -name '*.js' | wc -l | tr -d ' ') app scripts, runtime deps empty)"
