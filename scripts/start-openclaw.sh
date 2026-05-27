#!/usr/bin/env bash
set -euo pipefail
cd /home/fdelavega02/.openclaw/workspace
openclaw gateway start
printf '\nOpenClaw started.\n\n'
openclaw gateway status || true
printf '\nYou can close this window once the status looks good.\n'
