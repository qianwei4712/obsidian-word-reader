# Performance baselines

`npm run performance:check` runs DOCX, PPTX, and XLSX workloads. Every format
is measured five times in a new Node.js process, then checked against the
format-specific budget in this directory. Timing and heap gates use p95;
machine results also retain the median and every raw sample.

The shared metrics are:

- OOXML package loading and content parsing;
- time until the first readable content window;
- search and navigation latency;
- p95 scroll-window calculation or render time;
- sampled peak JavaScript heap growth;
- maximum parsed, estimated, or actual DOM nodes, with the measurement kind
  recorded alongside the value;
- current cache entries and declared cache limits; and
- cancellation and cleanup outcomes.

XLSX additionally keeps the existing 3.2.0 dense-streaming gates: `3,000 ms`
parse, `192 MiB` heap growth, and a `256 KiB` worksheet-data buffer. Production
artifacts remain limited to a `500 KiB` `dist/main.js` and an `8 MiB` release
zip.

Generated files are written to `benchmark-results/`:

- `performance.latest.json`: full machine-readable samples and aggregates;
- `performance-trend.json`: compact CI trend point with commit/run metadata;
- `performance-summary.md`: human-readable CI job summary.

GitHub Actions uploads all three files for 90 days on every CI run, including
failed budget runs. A failure identifies the format, aggregate, stage/metric,
observed value, and threshold.

Node.js measurements do not model Electron layout, paint, or long tasks. Use
the companion [Obsidian Desktop baseline](./OBSIDIAN_DESKTOP_BASELINE.md) to
calibrate these gates against actual application behavior.
