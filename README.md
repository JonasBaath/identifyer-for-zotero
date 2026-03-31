# Identifyer for Zotero

A local desktop app that finds plain-text citations in Word (.docx) and LibreOffice (.odt) documents, matches them against your local Zotero library, and converts them to proper Zotero field codes — ready for formatting with any citation style.

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
