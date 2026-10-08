# Local operational scripts

## `validate_daily_candidates.py`

Read-only by default validator for two weekday project-handoff artifacts. It
checks that each artifact has today's exact date, exactly one
concrete `Candidate:`, a source date/path, no blocked/quiet/maintenance wording,
and no near-duplicate candidate in the local audit history.

```bash
python3 scripts/validate_daily_candidates.py \
  --date 2026-10-07 \
  --first-handoff /path/to/first-handoff.md \
  --second-handoff /path/to/second-handoff.md
```

Use `--record` only after a successful run to add accepted candidate fingerprints
to `state/daily-candidate-provenance.json`. Run the read-only validation first
and record results only after both handoffs are accepted. The state directory is
gitignored because it records local audit history.
