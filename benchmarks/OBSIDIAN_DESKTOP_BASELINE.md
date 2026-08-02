# Obsidian Desktop performance baseline

Use this worksheet to calibrate Node.js performance gates against a real
Obsidian Desktop session. Keep the vault local and offline. Record metrics
only; never paste document text, speaker notes, cell values, internal XML,
absolute vault paths, screenshots of private content, or heap snapshots into
issues or CI artifacts.

## Environment

| Field | Value |
| --- | --- |
| Date | |
| Office Reader commit/version | |
| Obsidian version | |
| Electron/Chromium version | |
| OS and architecture | |
| CPU class / logical cores | |
| Installed memory | |
| Theme and UI scale | |

## Procedure

1. Restart Obsidian with only Office Reader and the test-vault plugins needed
   for the run. Open Developer Tools, enable the Performance panel's memory
   track, and clear the console.
2. For each format, run at least five cold opens. Capture plugin activation,
   package/parse work, first readable content, search, navigation, and a
   representative continuous scroll. Report median and p95.
3. Record long tasks (`>= 50 ms`), peak used JavaScript heap, and the maximum
   actual node count under the reader root. Do not export a heap snapshot.
4. Close the file, wait two animation frames, and record Blob URLs, format
   cache entries, queued/background work, observers, timers, listeners, and
   reader DOM nodes that remain.
5. Compare actual DOM and interaction timings with the matching Node.js trend
   point. If a budget changes, retain the five-run evidence and explain the
   runtime/fixture difference in the pull request.

## Results

| Format | Activation median / p95 | First readable median / p95 | Parse median / p95 | Search median / p95 | Navigation median / p95 | Scroll frame p95 | Long tasks / max | Peak heap | Max actual DOM | Cache / limit | Resources after close | Cleanup passed |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |
| DOCX | | | | | | | | | | | | |
| PPTX | | | | | | | | | | | | |
| XLSX | | | | | | | | | | | | |

## Safe diagnostic check

For each format, run **Copy Office performance diagnostics** and verify that
the JSON uses `kind: "performance"`, `measurement:
"obsidian-desktop-runtime"`, and explicitly reports that it contains no
document content, speaker notes, cell values, internal XML, or vault path.
File names may be anonymized before sharing.
