# Changelog

## Version 2.3.2 — Separator Auto-Alignment Fix

### Fixed
- Footnote separator line now correctly aligns **right** for Arabic (RTL) documents
  and **left** for English (LTR) documents automatically.
- Root cause: `WritingMode` was not set on page styles before `FootnoteLineAdjust`,
  causing LibreOffice to silently reset the separator direction.
- `adjust_footnote_separator()` is now called in **all three** operations:
  Convert, Format & Align, and Unify Font — previously it was missing from
  Convert and Unify Font.

### Changed
- `adjust_footnote_separator` now sets `WritingMode` on each page style **first**,
  then explicitly sets `FootnoteLineAdjust` to guarantee the correct direction.

---

## Version 2.3.1


### Added
- Improved bilingual user messaging across the add-on.
- Stronger language detection for Arabic documents.
- Better handling for empty markers and punctuation cleanup.
- More consistent undo behavior before showing dialogs.

### Improved
- Footnote formatting logic is more defensive when encountering invalid or partially formatted content.
- UI text is clearer and more consistent across Arabic and English.
- Documentation and release notes are prepared for a cleaner publishing process.

### Notes
- This update retains the current working functionality while improving operational stability and release readiness.
