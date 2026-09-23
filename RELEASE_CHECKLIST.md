# Release Readiness Checklist

This checklist is intended to preserve the current working state of the add-on while preparing it for a stable public release.

## 1) LibreOffice UI testing

Required before publishing:

- [ ] Install the extension on LibreOffice Writer on Windows.
- [ ] Open a document with Arabic text and verify the menu entries appear correctly.
- [ ] Open a document with English text and verify the menu entries appear correctly.
- [ ] Test the commands: Wrap, Convert, Reverse, Unify Font, Numbering per page, Continuous numbering.
- [ ] Confirm that the add-on works without crashing on a large document.
- [ ] Confirm that Ctrl + Z restores the previous document state after each operation.

## 2) Bilingual user messaging

- [ ] All menu titles are available in Arabic and English.
- [ ] All message boxes display both Arabic and English text where available.
- [ ] No untranslated strings remain in the add-on UI.
- [ ] The README, release notes, and extension metadata are consistent in language.

## 3) LibreOffice compatibility testing

Check at least:

- [ ] LibreOffice 7.x
- [ ] LibreOffice 24.x
- [ ] Windows 10/11
- [ ] Linux (optional but recommended)
- [ ] macOS (optional but recommended)

Use the same test document for each version and compare the output.

## 4) Edge-case handling

- [ ] Empty markers such as <> are removed cleanly.
- [ ] Text with quotes and hadith markers is preserved.
- [ ] Text with punctuation remains correctly glued to the footnote reference.
- [ ] Existing footnotes are detected and formatted without duplication.
- [ ] Reverse conversion returns to a valid < > marker format.
- [ ] A second conversion does not create duplicate footnotes.

## 5) Release documentation

- [ ] README is polished and accurate.
- [ ] Installation steps are simple and tested.
- [ ] Change log is updated.
- [ ] Compatibility matrix is included.
- [ ] Support and issue reporting links are present.

## 6) Final release gate

- [ ] Self-test script passes locally on a machine with Python available.
- [ ] The generated .oxt file is built and installed successfully.
- [ ] User-facing strings are reviewed one final time.
- [ ] The extension is tested in a real Writer document, not only in source code.

## Notes

The project is already in a strong state; the goal here is to protect that stability while tightening deployment readiness and language coverage.
