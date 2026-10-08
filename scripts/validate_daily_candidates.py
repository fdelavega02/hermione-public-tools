#!/usr/bin/env python3
"""Validate two weekday daily-project handoffs.

This is intentionally local-only.  It reads the durable handoff artifacts and,
when asked, appends accepted candidate fingerprints to a local audit file.  It
does not contact Discord, OpenClaw, or any model provider.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import sys
from datetime import date
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
HISTORY_PATH = ROOT / "state" / "daily-candidate-provenance.json"
FORBIDDEN = re.compile(
    r"\b(blocker|no[- ]candidate|quiet status|maintenance|dream artifact|generic support|recycled|stale)\b",
    re.IGNORECASE,
)
DATE_RE = re.compile(r"\b20\d{2}-\d{2}-\d{2}\b")
PATH_RE = re.compile(r"(?:`[^`]+`|(?:[\w.-]+/){1,}[\w.-]+)")


def normalize(candidate: str) -> str:
    text = candidate.casefold()
    text = DATE_RE.sub("<date>", text)
    text = re.sub(r"`[^`]+`", "<path>", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def load_history() -> list[dict[str, str]]:
    if not HISTORY_PATH.exists():
        return []
    try:
        parsed = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(f"invalid history file: {exc}") from exc
    if not isinstance(parsed, list):
        raise ValueError("history file must contain a JSON list")
    return parsed


def validate(agent: str, path: Path, expected_date: str, history: list[dict[str, str]]) -> dict:
    errors: list[str] = []
    candidate = ""
    if not path.exists():
        return {"agent": agent, "path": str(path), "ok": False, "errors": ["handoff file is missing"]}

    lines = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if len(lines) != 2:
        errors.append("handoff must contain exactly two non-empty lines")
    if not lines or lines[0] != f"Dashboard date: {expected_date}":
        errors.append(f"first line must be exactly: Dashboard date: {expected_date}")
    if len(lines) < 2 or not lines[1].startswith("Candidate: "):
        errors.append("second line must start with Candidate: followed by a concrete proposal")
    else:
        candidate = lines[1][len("Candidate: ") :].strip()
        if len(candidate) < 40:
            errors.append("candidate is too short to describe a concrete proposal")
        if FORBIDDEN.search(candidate):
            errors.append("candidate contains forbidden blocker/quiet/maintenance/stale wording")
        if expected_date not in candidate and not DATE_RE.search(candidate):
            errors.append("candidate must cite a source date")
        if not PATH_RE.search(candidate):
            errors.append("candidate must cite a source path or file")

        normalized = normalize(candidate)
        for prior in history:
            # Revalidating the same accepted artifact is normal.  Freshness is
            # about reuse on a later day, not an idempotent second check.
            if prior.get("agent") != agent or prior.get("date") == expected_date:
                continue
            ratio = difflib.SequenceMatcher(None, normalized, prior.get("normalized", "")).ratio()
            if ratio >= 0.90:
                errors.append(
                    f"candidate closely duplicates {prior.get('date', 'an earlier')} handoff "
                    f"(similarity {ratio:.0%})"
                )
                break
    return {
        "agent": agent,
        "path": str(path),
        "ok": not errors,
        "errors": errors,
        "candidate": candidate,
        "normalized": normalize(candidate) if candidate else "",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate two daily-project handoffs.")
    parser.add_argument("--date", default=date.today().isoformat(), help="expected handoff date (YYYY-MM-DD)")
    parser.add_argument("--first-handoff", required=True, type=Path, help="path to the first handoff artifact")
    parser.add_argument("--second-handoff", required=True, type=Path, help="path to the second handoff artifact")
    parser.add_argument("--record", action="store_true", help="append successful results to the local audit history")
    parser.add_argument("--json", action="store_true", help="emit the complete report as JSON")
    args = parser.parse_args()

    try:
        date.fromisoformat(args.date)
        history = load_history()
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    handoffs = {"Handoff A": args.first_handoff, "Handoff B": args.second_handoff}
    results = [validate(agent, path, args.date, history) for agent, path in handoffs.items()]
    report = {"date": args.date, "ok": all(item["ok"] for item in results), "results": results}

    if args.record and report["ok"]:
        existing = {(row.get("agent"), row.get("date"), row.get("sha256")) for row in history}
        for item in results:
            digest = hashlib.sha256(item["normalized"].encode()).hexdigest()
            record = {"agent": item["agent"], "date": args.date, "sha256": digest, "normalized": item["normalized"]}
            if (record["agent"], record["date"], record["sha256"]) not in existing:
                history.append(record)
        HISTORY_PATH.parent.mkdir(parents=True, exist_ok=True)
        HISTORY_PATH.write_text(json.dumps(history, indent=2) + "\n", encoding="utf-8")

    if args.json:
        print(json.dumps(report, indent=2))
    else:
        for item in results:
            status = "PASS" if item["ok"] else "FAIL"
            print(f"{status} {item['agent']}: {item['path']}")
            for error in item["errors"]:
                print(f"  - {error}")
        print("OVERALL " + ("PASS" if report["ok"] else "FAIL"))
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
