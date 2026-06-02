#!/usr/bin/env bash
set -euo pipefail
openclaw gateway start
printf '\nOpenClaw started.\n\n'
openclaw gateway status || true
printf '\nYou can close this window once the status looks good.\n'
