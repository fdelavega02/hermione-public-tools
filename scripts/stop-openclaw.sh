#!/usr/bin/env bash
set -euo pipefail
openclaw gateway stop
printf '\nOpenClaw stopped.\n\n'
openclaw gateway status || true
printf '\nYou can close this window once the status updates.\n'
