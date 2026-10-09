# Local operational scripts

## `validate_daily_candidates.py`

Read-only by default validator for Hermy and Herms's weekday project handoff
artifacts. It checks that each artifact has today's exact date, exactly one
concrete `Candidate:`, a source date/path, no blocked/quiet/maintenance wording,
and no near-duplicate candidate in the local audit history.

```bash
python3 scripts/validate_daily_candidates.py --date 2026-10-07
```

Use `--record` only after a successful run to add accepted candidate fingerprints
to `state/daily-candidate-provenance.json`. The 08:15 dashboard-prep job should
run the read-only validation first and record results only after it accepts both
handoffs.
