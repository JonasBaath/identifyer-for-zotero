"""Shared pytest fixtures."""
import os
import pytest
from pathlib import Path

FIXTURES_DIR = Path(__file__).parent / "fixtures"
TEST_DB = FIXTURES_DIR / "test_library.sqlite"


@pytest.fixture(scope="session")
def library():
    """Load test Zotero library once per session."""
    from zotero_client import ZoteroClient
    return ZoteroClient(str(TEST_DB)).load_library()


@pytest.fixture(scope="session")
def matcher(library):
    """CitationMatcher loaded with test library."""
    from matcher import CitationMatcher
    return CitationMatcher(library)


@pytest.fixture
def parser():
    """Bare CitationParser instance with no document loaded."""
    from citation_parser import CitationParser
    cp = CitationParser.__new__(CitationParser)
    cp.docx_path = "dummy.docx"
    cp._doc = None
    cp._all_paras = []
    cp._ref_section_start = None
    cp._ref_text = ""
    return cp
