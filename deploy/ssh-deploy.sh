#!/usr/bin/env bash
# Used by .github/workflows/deploy.yml. Reads secrets from the environment
# and never prints them.
set -euo pipefail
: "${DEPLOY_HOST:?Set the DEPLOY_HOST secret}"
: "${DEPLOY_USER:?Set the DEPLOY_USER secret}"
: "${DEPLOY_SSH_KEY:?Set the DEPLOY_SSH_KEY secret}"
: "${OPEN365_IMAGE:?Set OPEN365_IMAGE}"

dest="${DEPLOY_PATH:-/opt/open365}"
case "$dest$OPEN365_IMAGE$DEPLOY_HOST$DEPLOY_USER" in
  *[\'\"\;\|\&\$\(\)\`]*)
    echo "refusing path, image, host, or user with shell metacharacters" >&2
    exit 1
    ;;
esac

root="$(cd "$(dirname "$0")/.." && pwd)"
mkdir -p "${HOME}/.ssh"
chmod 700 "${HOME}/.ssh"
printf '%s\n' "$DEPLOY_SSH_KEY" > "${HOME}/.ssh/id_deploy"
chmod 600 "${HOME}/.ssh/id_deploy"
ssh-keyscan -H "$DEPLOY_HOST" >> "${HOME}/.ssh/known_hosts" 2>/dev/null

ssh_opts=(-i "${HOME}/.ssh/id_deploy" -o IdentitiesOnly=yes -o StrictHostKeyChecking=yes)
ssh "${ssh_opts[@]}" "$DEPLOY_USER@$DEPLOY_HOST" "mkdir -p '$dest'"
scp "${ssh_opts[@]}" "$root/deploy/docker-compose.yml" "$root/deploy/remote-up.sh" "$DEPLOY_USER@$DEPLOY_HOST:$dest/"
ssh "${ssh_opts[@]}" "$DEPLOY_USER@$DEPLOY_HOST" "cd '$dest' && OPEN365_IMAGE='$OPEN365_IMAGE' sh remote-up.sh"
