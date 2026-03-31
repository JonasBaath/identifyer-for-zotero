"""Basic smoke tests — no external dependencies beyond requirements.txt."""
from citation_parser import CitationParser


def test_parse_paragraph():
    cp = CitationParser.__new__(CitationParser)
    cp.docx_path = "dummy.docx"
    cp._doc = None
    cp._all_paras = []
    cp._ref_section_start = None
    cp._ref_text = ""

    results = cp._parse_paragraph(
        "See (Smith, 2020) and (Jones & Lee, 2019:45).", 0, in_footnote=False
    )
    assert len(results) == 2, f"Expected 2 citations, got {len(results)}"
    assert results[0].authors == ["Smith"], f"Unexpected authors: {results[0].authors}"
    assert results[0].year == "2020"
    assert results[1].year == "2019"
    print(f"Parsed {len(results)} citations OK")


def test_imports():
    import citation_parser  # noqa: F401
    import zotero_client    # noqa: F401
    import matcher          # noqa: F401
    import field_writer     # noqa: F401
    print("All imports OK")


if __name__ == "__main__":
    test_imports()
    test_parse_paragraph()
    print("All tests passed.")
