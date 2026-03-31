# Identifyer for Zotero

A local desktop app that finds plain-text citations in Word (.docx) and LibreOffice (.odt) documents, matches them against your local Zotero library, and converts them to proper Zotero field codes — ready for formatting with any citation style.

The typical use case is a document that has been written with plain-text citations — either because it was drafted without a reference manager, converted from another format, or received from a collaborator who uses a different system. Identifyer for Zotero bridges the gap: it reads the document as-is, identifies every citation, looks each one up in your Zotero library using fuzzy matching, and produces a new document where all recognised citations have been replaced with live Zotero field codes. Once opened in Word or LibreOffice with the Zotero plugin, the document behaves exactly as if the citations had been inserted through Zotero from the start — you can switch citation styles, update references, and let Zotero regenerate the bibliography automatically. The matching engine handles the messiness of real-world academic text: author names may be abbreviated, hyphenated, or inconsistently formatted; citations may include page numbers, section locators, or prefixes such as "cf." or "see also"; and multi-author citations with "et al." abbreviations are resolved against the full author list in Zotero. Where a citation is genuinely ambiguous — for example, an author with multiple publications in the same year — the app presents the candidates side by side so you can pick the right one. Citations that cannot be matched are flagged for review, and the app also identifies references that appear in the document's reference list but are never cited in the text.

## How it works

1. You provide a document and your Zotero database
2. The app extracts all in-text citations (author–year and numbered styles)
3. Each citation is fuzzy-matched against your Zotero library
4. You review matches, accept suggestions, and resolve ambiguous cases
5. Save a new document with Zotero field codes in place of plain-text citations
6. Open the new document in Word/LibreOffice with Zotero to refresh the bibliography

## Requirements

- Python 3.9 or later
- [Zotero](https://www.zotero.org/) installed locally (the app reads your local SQLite database)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

```bash
python3 main.py
```

The app opens automatically in your default browser at `http://127.0.0.1:7474`.

## Supported formats

| Input | Output |
|-------|--------|
| `.docx` (Word) | `.docx` with Zotero field codes |
| `.odt` (LibreOffice) | `.odt` with Zotero field codes |

## Features

- Fuzzy author and title matching — handles typos, abbreviations, and name variants
- Detects citations in body text, footnotes, endnotes, and table cells
- Accepts tracked changes before parsing — works on documents with revision history
- Strips existing Zotero/EndNote field codes and re-matches from scratch
- Adjustable matching sensitivity (author threshold, year tolerance)
- Reports citations present in the text but missing from the reference list
- Works on macOS, Windows, and Linux

## Notes

- The original document is never modified — output is always saved as a new file
- Zotero must be closed (or the database unlocked) when running the app
- The app runs entirely locally — no data leaves your machine
