#!/usr/bin/env bash
set -euo pipefail
cd /home/fdelavega02/.openclaw/workspace
openclaw gateway stop
printf '\nOpenClaw stopped.\n\n'
openclaw gateway status || true
printf '\nYou can close this window once the status updates.\n'
