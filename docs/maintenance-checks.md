# Maintenance Checks

After an automation update, verify the boring pieces before trusting the next scheduled run.

This note is intentionally generic. It does not include private schedules, account IDs, channels, inbox rules, payloads, credentials, local paths, or Francisco-specific operational details.

## Scheduler Updates

When scheduled jobs migrate or the runtime changes:

- Confirm the scheduler is running.
- List enabled and disabled jobs.
- Check each important job's schedule, timezone, runtime/model, tool access, and delivery mode.
- Run a low-risk job manually when possible.
- Review recent failures and confirm whether they are historical or still active.
- Patch disabled legacy jobs if they might be re-enabled later.

## Negative Scans

Some checks are supposed to find nothing, such as sensitive-pattern scans. Wrap those commands so automation logs show the real outcome:

- Matches found: report the findings and fail the check.
- No matches found: treat as success.
- Tool error: treat as a real failure.

```bash
set +e
rg -n 'PRIVATE_PATTERN|SECRET_PATTERN' .
status=$?
set -e

if [ "$status" -eq 0 ]; then
  echo "scan found matches"
  exit 1
fi

if [ "$status" -eq 1 ]; then
  echo "scan passed: no matches"
  exit 0
fi

echo "scan failed: rg exited with ${status}"
exit "$status"
```

Public examples should use placeholders only. Keep real private identifiers, alert terms, account names, paths, and delivery targets out of public docs.

## Recovered Failures

If an early automation step fails and the run later recovers, name that recovery in the final note. For example: a stale remote can reject a push, but a fetch, fast-forward, recheck, and successful push means the final outcome is recovered rather than still failed.

Keep public writeups generic. Do not include private remotes, branch names, local worktree paths, delivery targets, raw logs, account identifiers, or workflow payloads.

## Scheduled Delivery

When a scheduled job already has a configured delivery target, make the final job response contain the user-facing result. Avoid sending the important content through an ambiguous inner route and returning only a short completion note to the scheduler.

Public notes should describe the delivery shape only. Do not include private aliases, account IDs, channel IDs, local session keys, or raw job payloads.
