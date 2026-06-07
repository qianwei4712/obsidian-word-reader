# Obsidian community plugin

## Required standards

- Follow the latest Obsidian Developer Policies and Plugin Guidelines.
- Use the official `eslint-plugin-obsidianmd` recommended configuration.
- Treat every ESLint warning as a release blocker.
- Use supported Obsidian APIs compatible with `manifest.json.minAppVersion`.
- Preserve local-only, read-only Word document behavior unless a task explicitly changes the product scope.
- Do not add telemetry, remote code execution, dynamic script injection, or undisclosed network requests.

## Required checks

Before completing any source change, run:

```bash
npm run lint
npm run typecheck
npm run build
```

Prefer the combined command:

```bash
npm run check
```

Before publishing a release, additionally run:

```bash
npm run release
node scripts/release-check.mjs --tag X.Y.Z
```

Do not claim completion while any required command reports warnings or errors.

## Obsidian compatibility

- Use `activeDocument`, `activeWindow`, or a node's `doc`/`win` for popout-window compatibility.
- Use Obsidian's `.instanceOf(...)` for DOM node and UI event type checks.
- Do not include the plugin ID in command IDs.
- Avoid deprecated APIs and update `minAppVersion` when adopting newer APIs.
- Use `requestUrl` for actual network requests. Keep this plugin offline by default.
- Guard desktop-only Node.js and Electron access with `Platform.isDesktopApp`.
- Register disposable listeners and timers through Obsidian lifecycle helpers where ownership permits.

## References

- https://github.com/obsidianmd/obsidian-sample-plugin
- https://github.com/obsidianmd/eslint-plugin
- https://docs.obsidian.md/Developer+policies
- https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines
