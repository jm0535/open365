#!/usr/bin/env bash
# Build the runtime image and require /healthz plus the app shell.
# GitHub-hosted runners have Docker. This sandbox does not — the script
# fails closed when CI=true and otherwise explains why it skipped.
set -euo pipefail
cd "$(dirname "$0")/.."

if ! command -v docker >/dev/null 2>&1; then
  echo "docker is not installed; image smoke cannot run here" >&2
  if [ "${CI:-}" = "true" ]; then
    exit 1
  fi
  exit 0
fi

docker build -t open365:ci .
name="open365-ci-$$"
docker rm -f "$name" >/dev/null 2>&1 || true
docker run -d --name "$name" -p 127.0.0.1:4174:4173 open365:ci
cleanup() { docker rm -f "$name" >/dev/null 2>&1 || true; }
trap cleanup EXIT

ready=0
for _ in $(seq 1 40); do
  if curl -sf http://127.0.0.1:4174/healthz >/dev/null 2>&1; then
    ready=1
    break
  fi
  sleep 0.5
done
if [ "$ready" != 1 ]; then
  echo "container did not become healthy" >&2
  docker logs "$name" >&2 || true
  exit 1
fi

curl -sf http://127.0.0.1:4174/healthz | grep -q '"ok":true'
code=$(curl -s -o /dev/null -w '%{http_code}' http://127.0.0.1:4174/)
test "$code" = "200"
curl -sf http://127.0.0.1:4174/api/me | grep -q '"ok":true'
echo "image smoke passed"
