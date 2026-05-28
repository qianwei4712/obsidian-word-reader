# Changelog

## 0.5.0 - 2026-05-28

### Added

- Added previous and next controls for rendered text search results.
- Added `Enter` and `Shift` + `Enter` keyboard navigation in the search input.
- Added current search result highlighting.

### Changed

- Search result counts now show the current result and total result count.
- Search navigation now scrolls the active result into view.

## 0.4.0 - 2026-05-28

### Added

- Added an Obsidian settings tab.
- Added persisted settings for default zoom, default fit width, image preview, and large file warning size.
- Added an external opening note explaining that `.docx` files use the operating system default application.

### Changed

- Word previews now use the configured default zoom and fit-width behavior when opened.
- Large file warnings now use the configured size threshold.
- Image click preview can now be disabled from settings.

## 0.3.0 - 2026-05-28

### Added

- Added continuous document zoom with the toolbar percentage input.
- Added `Ctrl` + mouse wheel zoom in the Word preview.
- Added click-to-preview image modal for rendered Word images.

### Changed

- Replaced fixed zoom levels with percentage-based zoom.
- Improved zoom behavior to keep the cursor position stable while zooming.

### Docs

- Updated README files for zoom, image preview, and `.doc` support limits.

## 0.2.0

- Added read-only `.docx` view registration.
- Added rendering through `docx-preview`.
- Added toolbar actions for reload, zoom, fit width, search, copy text, external open, and Markdown note creation.
- Added safe same-name Markdown summary note generation.

## 0.1.0

- Initial project scaffold.
