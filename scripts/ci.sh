#!/usr/bin/env bash
# The same gate GitHub Actions runs. Reuses a server already on :4173
# (the preview) and only starts one when the port is free. Never deletes ./data.
set -euo pipefail
cd "$(dirname "$0")/.."
export CI="${CI:-true}"

bash scripts/syntax.sh

started=0
pid=""
cleanup() {
  if [ "$started" = 1 ] && [ -n "$pid" ]; then
    kill "$pid" 2>/dev/null || true
    wait "$pid" 2>/dev/null || true
  fi
}
trap cleanup EXIT

if ! curl -sf http://127.0.0.1:4173/healthz >/dev/null 2>&1; then
  node server.js > /tmp/open365-ci-server.log 2>&1 &
  pid=$!
  started=1
  ready=0
  for _ in $(seq 1 50); do
    if curl -sf http://127.0.0.1:4173/healthz >/dev/null 2>&1; then
      ready=1
      break
    fi
    if ! kill -0 "$pid" 2>/dev/null; then
      echo "server exited during startup" >&2
      cat /tmp/open365-ci-server.log >&2 || true
      exit 1
    fi
    sleep 0.2
  done
  if [ "$ready" != 1 ]; then
    echo "server did not become healthy on :4173" >&2
    cat /tmp/open365-ci-server.log >&2 || true
    exit 1
  fi
fi

bash test/run.sh
echo "CI gate passed"
