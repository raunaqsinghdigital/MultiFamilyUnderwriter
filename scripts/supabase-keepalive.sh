#!/bin/bash
# Supabase keepalive — MultiFamilyUnderwriter (project jpyoondbepffknsrbkoe)
#
# Free-plan Supabase projects are paused automatically after 7 consecutive days
# with no API activity. This app has real but sporadic usage, so Supabase has been
# mailing a "scheduled to be paused" warning weekly. One small API round-trip per
# day resets that 7-day clock.
#
# Runs unattended via launchd (com.maplesyrupmoney.supabase-keepalive), daily at
# 09:20 local. Costs: 2 HTTPS requests/day and a few KB of egress — nothing
# measurable against the Free plan's 5 GB egress / 500 MB database allowance.
# Consumes zero agent tokens: no model is in the loop.
#
# Anon key resolution, in order:
#   1. $SUPABASE_ANON_KEY — how launchd supplies it (see the plist). Required there,
#      because macOS TCC forbids a LaunchAgent from reading anything under
#      ~/Documents; that is also why the scheduled job runs an installed copy of
#      this file from ~/Library/Application Support rather than this repo path.
#   2. static/app.js — used when you run this by hand from the repo. That file is
#      the key's single source of truth and already ships it to every browser.
# The anon key is a public client credential, not a secret; the service-role key
# is never used here.
#
# After editing this file, re-install the copy launchd runs:
#   install -m 755 scripts/supabase-keepalive.sh \
#     "$HOME/Library/Application Support/msm-supabase-keepalive/supabase-keepalive.sh"

set -uo pipefail

PROJECT_REF="jpyoondbepffknsrbkoe"
BASE="https://${PROJECT_REF}.supabase.co"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
APP_JS="${REPO_ROOT}/static/app.js"
LOG="${HOME}/Library/Logs/supabase-keepalive.log"
STAMP="$(date -u '+%Y-%m-%dT%H:%M:%SZ')"

mkdir -p "$(dirname "$LOG")"

log() { echo "${STAMP} $*" >>"$LOG"; }

fail() {
  log "FAIL $*"
  osascript -e "display notification \"$*\" with title \"Supabase keepalive failed\" subtitle \"MultiFamilyUnderwriter\"" >/dev/null 2>&1
  exit 1
}

ANON_KEY="${SUPABASE_ANON_KEY:-}"
if [ -z "$ANON_KEY" ] && [ -r "$APP_JS" ]; then
  ANON_KEY="$(grep -A2 'SUPABASE_ANON_KEY' "$APP_JS" | grep -o 'eyJ[A-Za-z0-9._-]*' | head -1)"
fi
[ -n "$ANON_KEY" ] || fail "no anon key: \$SUPABASE_ANON_KEY unset and $APP_JS unreadable"

# 1. PostgREST -> Postgres. RLS returns [] for anon, but the query really executes,
#    which is what makes this count as database activity rather than a cached 200.
rest_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 30 \
  -H "apikey: ${ANON_KEY}" -H "Authorization: Bearer ${ANON_KEY}" \
  "${BASE}/rest/v1/saved_properties?select=id&limit=1")"

# 2. GoTrue. Second independent service on the project, so a single service being
#    down does not look like the whole project going quiet.
auth_code="$(curl -sS -o /dev/null -w '%{http_code}' --max-time 30 \
  -H "apikey: ${ANON_KEY}" \
  "${BASE}/auth/v1/settings")"

if [ "$rest_code" = "200" ] && [ "$auth_code" = "200" ]; then
  log "OK rest=${rest_code} auth=${auth_code}"
else
  fail "rest=${rest_code} auth=${auth_code}"
fi

# Keep the log from growing without bound (~1 line/day, but be safe over years).
if [ "$(wc -l <"$LOG")" -gt 800 ]; then
  tail -n 400 "$LOG" >"${LOG}.tmp" && mv "${LOG}.tmp" "$LOG"
fi
