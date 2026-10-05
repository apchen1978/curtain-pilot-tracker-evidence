# Pilot Tracker — Evidence

Rerunnable evidence for two portfolio works:

- **Pilot 追蹤器** — one tracker from lead to quote to follow-up
- **Pilot 模擬套件** — three simulation rounds (baseline / data model / quote versioning)

## Contents

- work/ — build / simulate / inspect / verify scripts (zero dependencies, Node)
- outputs/ — XLSX trackers + .inspect.ndjson evidence files (per round)
- simulation-004/ — lead-004 simulation (data, plan, log, script)

## Verification

Checks pass (lead → follow-up pipeline, quote versioning, data-model fixes).
Rerun locally:

node work/simulate_pilot_001.mjs
node work/simulate_pilot_002.mjs
node work/simulate_pilot_003.mjs
node work/inspect_curtain_tracker.mjs

## Note

Simulated pilot data only — no real prospect or client records.
