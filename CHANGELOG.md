# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.4] - 2026-06-22

### Added
- README latest-release badge linking to the newest GitHub release.

### Changed
- Documentation only — no functional changes to the macros since 1.1.1.

## [1.1.3] - 2026-06-22

### Added
- README **Examples** section with exported `histplot` and `boxplot` chart
  screenshots (`docs/`).

### Changed
- Documentation only — no functional changes to the macros since 1.1.1.

## [1.1.2] - 2026-06-22

### Changed
- Documentation only — no functional changes to the macros since 1.1.1.
- Rewrote the README introduction and added a changelog (linked from the README).

## [1.1.1] - 2026-06-22

### Added
- `histplot` and `boxplot` now appear at the bottom of the **Data** menu, just
  below the built-in **Statistics** submenu (labelled `histplot` and `boxplot`).

### Fixed
- Data-menu entries now actually render. In the withdrawn 1.1.0 they were merged
  into the **Data ▸ Statistics** submenu, which is built lazily from a separate
  popup definition, so a static `Addons.xcu` could not inject into it and the
  items silently failed to appear.

### Notes
- 1.1.0 was published and then withdrawn because its menu entries did not show
  up. Use 1.1.1 instead.

## [1.0.0] - 2026-06-22

### Added
- Initial release, verified working in LibreOffice Calc.
- `histplot` — histogram (vertical column chart) from a selected column, with a
  user-specified bin width.
- `boxplot` — box-and-whisker plot with the five-number summary, median line,
  whiskers, and Tukey (1.5×IQR) outlier detection and markers.
- Installable `.oxt` package; build from source with `build.ps1`.

[1.1.4]: https://github.com/davidjayjackson/stats_plots/releases/tag/v1.1.4
[1.1.3]: https://github.com/davidjayjackson/stats_plots/releases/tag/v1.1.3
[1.1.2]: https://github.com/davidjayjackson/stats_plots/releases/tag/v1.1.2
[1.1.1]: https://github.com/davidjayjackson/stats_plots/releases/tag/v1.1.1
[1.0.0]: https://github.com/davidjayjackson/stats_plots/releases/tag/v1.0.0
