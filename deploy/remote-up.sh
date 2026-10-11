#!/bin/sh
# Runs on the deploy host, in the directory that holds docker-compose.yml.
set -eu
cd "$(dirname "$0")"
if [ -z "${OPEN365_IMAGE:-}" ]; then
  echo "OPEN365_IMAGE is required" >&2
  exit 1
fi

docker compose pull
docker compose up -d --remove-orphans

i=0
while [ "$i" -lt 30 ]; do
  cid=$(docker compose ps -q open365 || true)
  status=$(docker inspect --format '{{if .State.Health}}{{.State.Health.Status}}{{else}}none{{end}}' "$cid" 2>/dev/null || echo starting)
  if [ "$status" = "healthy" ]; then
    echo "open365 healthy ($OPEN365_IMAGE)"
    exit 0
  fi
  i=$((i + 1))
  sleep 2
done

echo "open365 did not become healthy" >&2
docker compose logs --tail 80 >&2 || true
exit 1
