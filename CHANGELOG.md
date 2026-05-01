# Changelog

All notable changes to Identifyer for Zotero are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.2.1] — 2026-04-30

A focused parser and matcher overhaul driven by real-world testing on a
Chicago-style academic manuscript. False negatives went from 23 to 0
(out of 165 citations); 11 more citations are now correctly matched
without manual intervention.

### Added
- **Last-resort fallback in matcher.** When all normal matching passes fail,
  the matcher now searches the document's own bibliography first
  (author surname fuzzy, ignoring year), then Zotero ignoring year. The UI
  surfaces both as gold-orange hints under the failed citation, so the user
  can quickly tell whether a reference is genuinely missing or just has a
  typo'd year.
- **`run.sh` / `run.bat` launchers.** First-run creates a local `.venv` and
  installs dependencies automatically. Resolves the PEP 668 block on
  Homebrew Python and similar externally-managed interpreters.
- **`CHANGELOG.md`.** This file.
- **14 regression tests** covering each new bug class (73 passing total).

### Fixed
- **Curly apostrophes in surnames.** `O'Gorman` and `Pellegrino's` (where
  Word auto-corrects `'` → `’`) now parse correctly.
- **Prose words no longer matched as authors.** `UNIT_RE` is case-sensitive
  on author tokens, so words like *the*, *for*, *common humanity* cannot be
  picked up as surnames inside long discursive parentheticals. `et al.` /
  `m.fl.` continue to work via inline case-insensitive flag.
- **Prefix tokens stripped before matching.** `see`, `cf.`, `e.g.`, `also`,
  *i.e.*, *compare*, etc. at parenthetical start or after `;` are now
  masked before `UNIT_RE` runs, so they cannot be picked up as authors.
  Previously produced false matches like `(f. Stone et al., 2018)` and
  `(also & Lane, 2014)`.
- **Initial-without-period stripped.** `M Nilsson` (initial M for Magnus
  without period) now reduces to `Nilsson` for matcher lookup, so it
  matches a Zotero entry stored as `Nilsson, Magnus`.
- **Trailing possessive `'s` stripped** before fuzzy author matching.
- **Inline multi-author citations.** `Koponen and Niva (2020)` — `and`
  between two capitalised name words is now a co-author connector, not a
  sentence-initial stopword. Sentence-initial *And* is still excluded.
  Also fixed: inline-captured authors are now split into individual
  surnames, so the matcher compares each against Zotero separately
  (previously stored as a single string, sinking fuzzy ratio below the
  threshold).
- **Chicago narrative citations with the author outside the parenthetical.**
  Patterns like `Lane (2014, 20; cf. Bildtgård 2010)` and
  `"… quote …" (year, page)` now resolve the leading year to the correct
  author. The scanner skips quoted spans and falls back to a
  sentence-level search for possessive (`Lane's`) or verb-led
  (`Lane argues`, `Smith describes`) author mentions.

### Notes for users upgrading from binaries
The binary distribution for v0.2.1 is built automatically by the release
workflow. If you run from source, pull `main` and re-run `./run.sh`
(macOS/Linux) or `run.bat` (Windows) — the launcher will install any
new dependencies automatically.

### Known limitations
Simple Chicago "quote followed by `(year, page)`" citations *without*
sub-citations are still missed in some cases — the leading-year handler
intentionally stays out of non-compound parens to avoid duplicating
INLINE matches. These need manual review for now.

## [0.2.0] — 2026-04-23

### Added
- Release workflow that builds standalone PyInstaller binaries for
  Linux x64, Windows x64, macOS x64 (Rosetta), and macOS arm64 on every
  `v*` tag.

### Fixed
- Traffic logging now upserts all 14 days returned by the GitHub API
  and back-fills missed days, instead of only writing today's row.

## [0.1.0] — 2026-03-31

Initial public beta release.

[0.2.1]: https://github.com/JonasBaath/identifyer-for-zotero/compare/v0.2.0...v0.2.1
[0.2.0]: https://github.com/JonasBaath/identifyer-for-zotero/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/JonasBaath/identifyer-for-zotero/releases/tag/v0.1.0
