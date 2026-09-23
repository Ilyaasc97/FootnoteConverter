# Testing and Validation Guide

## Scope

This guide is intended to verify the add-on before release while keeping the current stable behavior intact.

## Required manual tests

### A. Arabic document test
1. Open LibreOffice Writer.
2. Create a new Arabic document.
3. Write a paragraph with marker text like: <الحاشية>.
4. Run the conversion command.
5. Confirm that the marker becomes a footnote and the reference is inserted as (1).
6. Confirm that the footnote text stays under the page and is aligned correctly.

### B. English document test
1. Create an English document.
2. Use markers like <source>.
3. Run the conversion.
4. Confirm that the output remains correct and the punctuation remains attached.

### C. Edge cases
- Empty marker: <>
- Marker next to punctuation: <note>.
- Existing footnote text already inserted
- Mixed Arabic and English content
- Large document with many markers

### D. Undo test
- After conversion, press Ctrl + Z.
- Confirm the document returns to the original state.

## Compatibility matrix

| Environment | Status | Notes |
| --- | --- | --- |
| LibreOffice 7.x | Recommended | Primary compatibility target |
| LibreOffice 24.x | Recommended | Current compatibility target |
| Windows 10/11 | Required | Main OS target |
| Linux | Optional | Should be checked |
| macOS | Optional | Should be checked |

## Validation rule

The extension is ready for publication only when the following are true:

- UI labels are bilingual and readable.
- Conversion works on real documents.
- Empty markers do not create phantom footnotes.
- Undo works reliably.
- Compatibility is confirmed on at least the supported Windows versions.

## Important note

This repository includes logic self-tests, but true final acceptance must be validated in a machine with LibreOffice Writer installed.
