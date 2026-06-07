# Obsidian Community Plugin Submission Checklist

Use this checklist before submitting the initial release to the Obsidian
Community directory and before publishing later releases.

## Automated checks

- [ ] Run `npm ci` from a clean checkout.
- [ ] Run `npm run check`.
- [ ] Run `npm run release`.
- [ ] Run `node scripts/release-check.mjs --tag X.Y.Z`.
- [ ] Confirm all commands finish without warnings or errors.
- [ ] Confirm `release/obsidian-word-reader-X.Y.Z.zip` contains only
  `main.js`, `manifest.json`, and `styles.css`.

## Repository metadata

- [ ] `README.md` explains the plugin purpose, installation, usage, support
  boundaries, security behavior, and any external access.
- [ ] `LICENSE` exists and dependency licenses are compatible.
- [ ] `manifest.json` is committed on the default branch and contains the
  intended release version.
- [ ] `versions.json` maps the release version to the correct minimum Obsidian
  version.
- [ ] The manifest ID is unique, lowercase, contains only letters and hyphens,
  does not contain `obsidian`, and does not end in `plugin`.
- [ ] The manifest description is no more than 250 characters, ends with a
  period, and describes the user-visible action.
- [ ] `isDesktopOnly` remains `true` while the plugin uses Node.js or Electron
  APIs.
- [ ] The repository is public and its source code matches the release assets.

## Policy and review

- [ ] Review the current Obsidian Developer policies.
- [ ] Review the current submission requirements and plugin guidelines.
- [ ] Confirm there is no telemetry, self-update mechanism, obfuscated code,
  dynamic script injection, or undisclosed network use.
- [ ] Confirm the plugin does not modify or write back to source `.docx`
  files.
- [ ] Confirm UI text uses sentence case and settings do not repeat the plugin
  name as a heading.
- [ ] Confirm desktop-only Electron and Node.js access is guarded.
- [ ] Run the manual scenarios in `STABILITY.md` in a clean test vault.

## GitHub release

- [ ] Create a plain semantic-version tag such as `1.5.0`; do not add a `v`
  prefix.
- [ ] Confirm the tag exactly matches `package.json`, `package-lock.json`,
  `manifest.json`, `versions.json`, and the latest `CHANGELOG.md` section.
- [ ] Confirm the GitHub release has binary attachments named `main.js`,
  `manifest.json`, and `styles.css`.
- [ ] Install the release assets into
  `.obsidian/plugins/obsidian-word-reader/` and verify that Obsidian loads the
  plugin without console errors.

## Initial directory submission

- [ ] Sign in at <https://community.obsidian.md> and link the GitHub account
  that owns the repository.
- [ ] Submit
  <https://github.com/qianwei4712/obsidian-word-reader> as a new plugin.
- [ ] Confirm the directory reads the expected manifest from the default
  branch.
- [ ] Address automated review feedback in a new patch release rather than
  replacing an existing release.

## Official references

- [Submit your plugin](https://docs.obsidian.md/Plugins/Releasing/Submit+your+plugin)
- [Submission requirements](https://docs.obsidian.md/Plugins/Releasing/Submission+requirements+for+plugins)
- [Plugin guidelines](https://docs.obsidian.md/Plugins/Releasing/Plugin+guidelines)
- [Developer policies](https://docs.obsidian.md/Developer+policies)
- [Manifest reference](https://docs.obsidian.md/Reference/Manifest)
