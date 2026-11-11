# Changelog

All notable changes to this project are documented here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses
[Semantic Versioning](https://semver.org/): `v1.x.y` releases never break the
`@v1` contract (inputs, outputs, comment format markers).

## [Unreleased]

## [1.0.0] - 2026-11-11

First tagged semantic-version release. The floating `v1` tag points at the
latest `v1.x.y` release.

### Added
- GitHub Action that posts an audit diff of changed `.xlsx`/`.xlsm` workbooks
  as a single sticky PR comment, plus the full report in the job summary and a
  JSON report.
- Risk flags: formula overwritten with a hardcoded value, new `#REF!`/error
  results, hardcoded numbers inside formulas, inconsistent formulas,
  order-of-magnitude input changes, external links, named-range and sheet
  changes.
- Impact analysis: recalculated values of unchanged formulas and the
  downstream dependency graph, with final outputs highlighted.
- Row/column insertion and sheet-rename detection, so shifted formulas are not
  reported as changes.
- `fail-on` input to fail the check at a chosen risk severity.
- Google Sheets snapshot sync (`sheetaudit gsheet-sync`), experimental.
- `sheetaudit` CLI: `diff`, `git`, `comment`, `textconv`, `gsheet-sync`.
- MIT license.

### Changed
- Action runtime dependencies use the Node 24 majors
  (`actions/setup-python@v7`; examples use `actions/checkout@v7`).
- Dependencies are pinned below their next major version
  (`openpyxl>=3.1,<4`).
- Requires Python 3.10+.

[Unreleased]: https://github.com/RyMB-007/sheet-audit-trail/compare/v1.0.0...HEAD
[1.0.0]: https://github.com/RyMB-007/sheet-audit-trail/releases/tag/v1.0.0
