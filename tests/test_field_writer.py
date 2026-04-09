"""Tests for field_writer.py."""
import os
import tempfile
import pytest
from pathlib import Path
from docx import Document as DocxDocument


FIXTURES_DIR = Path(__file__).parent / "fixtures"
TEST_DB = FIXTURES_DIR / "test_library.sqlite"


def _make_simple_docx(text: str) -> str:
    """Create a minimal .docx with one paragraph of text. Returns temp path."""
    doc = DocxDocument()
    doc.add_paragraph(text)
    tmp = tempfile.mktemp(suffix=".docx")
    doc.save(tmp)
    return tmp


def _load_library():
    from zotero_client import ZoteroClient
    return ZoteroClient(str(TEST_DB)).load_library()


def _match(library, text):
    from citation_parser import CitationParser
    from matcher import CitationMatcher
    tmp = _make_simple_docx(text)
    try:
        cits = CitationParser(tmp).parse()
        results = CitationMatcher(library).match_all(cits)
        return tmp, results
    except Exception:
        os.unlink(tmp)
        raise


class TestOutputFile:

    def test_original_not_overwritten(self):
        """Output must be saved to a new file; original unchanged."""
        from field_writer import write_zotero_document
        library = _load_library()
        src, results = _match(library, "See (Smith, 2020) for details.")
        mtime_before = os.path.getmtime(src)
        out = src.replace(".docx", "_zotero.docx")
        try:
            write_zotero_document(src, out, results)
            assert os.path.exists(out), "Output file was not created"
            assert os.path.getmtime(src) == mtime_before, "Original file was modified"
        finally:
            os.unlink(src)
            if os.path.exists(out):
                os.unlink(out)

    def test_output_contains_field_code(self):
        """Output .docx must contain a Zotero ADDIN field code for matched citation."""
        from field_writer import write_zotero_document
        library = _load_library()
        src, results = _match(library, "See (Smith, 2020) for details.")
        out = src.replace(".docx", "_zotero.docx")
        try:
            write_zotero_document(src, out, results)
            W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            doc = DocxDocument(out)
            found = any(
                elem.text and "ZOTERO_ITEM" in elem.text
                for elem in doc.element.body.iter(f"{{{W}}}instrText")
            )
            assert found, "No Zotero field code found in output document"
        finally:
            os.unlink(src)
            if os.path.exists(out):
                os.unlink(out)

    def test_unmatched_citation_stays_plain(self):
        """Unmatched citations must remain as plain text in the output."""
        from field_writer import write_zotero_document
        library = _load_library()
        src, results = _match(library, "See (Zzzyxq, 9999) for details.")
        out = src.replace(".docx", "_zotero.docx")
        try:
            write_zotero_document(src, out, results)
            doc = DocxDocument(out)
            full_text = " ".join(p.text for p in doc.paragraphs)
            assert "Zzzyxq" in full_text, "Unmatched citation text was removed"
        finally:
            os.unlink(src)
            if os.path.exists(out):
                os.unlink(out)
