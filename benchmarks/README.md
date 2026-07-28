# Performance baselines

`npm run performance:check` keeps the production `2.x` bundle at or below
`500 KiB` and runs the generated `100,000`-row sparse XLSX workload.

The XLSX result records:

- package-to-first-window time;
- p95 virtual-scroll frame calculation time across 240 positions;
- sampled peak JavaScript heap growth;
- maximum estimated mounted DOM node count; and
- whether an in-flight worksheet load was cancelled without leaving a cached
  result.

The latest machine result is written to
`benchmark-results/xlsx-100k.latest.json`. It is intentionally ignored because
timings and heap readings are environment-specific; the committed
`xlsx-100k-budget.json` is the release gate.
