"""Tests for ZoteroClient collection helpers."""
import sqlite3
import tempfile
from pathlib import Path

import pytest

from zotero_client import ZoteroClient


def _make_db_with_collections(path: str) -> None:
    """Build a minimal Zotero-like sqlite fixture with a collection tree.

    Shape:
        Root (id=1)
          ├── Child A (id=2) — 2 items
          └── Child B (id=3)
                └── Grandchild (id=4) — 1 item
        Orphan (id=5) — 1 item, no parent
        Deleted (id=6) — should be filtered out

    Plus one item (id=200) that is in no collection, and one item (id=201)
    soft-deleted.
    """
    conn = sqlite3.connect(path)
    conn.executescript(
        """
        CREATE TABLE collections (
            collectionID INTEGER PRIMARY KEY,
            collectionName TEXT NOT NULL,
            parentCollectionID INTEGER,
            libraryID INTEGER NOT NULL
        );
        CREATE TABLE deletedCollections (collectionID INTEGER PRIMARY KEY);
        CREATE TABLE items (
            itemID INTEGER PRIMARY KEY,
            libraryID INTEGER NOT NULL
        );
        CREATE TABLE collectionItems (
            collectionID INTEGER NOT NULL,
            itemID INTEGER NOT NULL,
            PRIMARY KEY (collectionID, itemID)
        );
        CREATE TABLE deletedItems (itemID INTEGER PRIMARY KEY);

        INSERT INTO collections VALUES
            (1, 'Root',       NULL, 1),
            (2, 'Child A',    1,    1),
            (3, 'Child B',    1,    1),
            (4, 'Grandchild', 3,    1),
            (5, 'Orphan',     NULL, 1),
            (6, 'Deleted',    NULL, 1);
        INSERT INTO deletedCollections VALUES (6);

        INSERT INTO items VALUES
            (100, 1), (101, 1), (102, 1), (103, 1), (200, 1), (201, 1);
        INSERT INTO deletedItems VALUES (201);

        INSERT INTO collectionItems VALUES
            (2, 100), (2, 101),   -- Child A: two items
            (4, 102),             -- Grandchild: one item
            (5, 103),             -- Orphan: one item
            (6, 201);             -- Deleted collection + deleted item
        """
    )
    conn.commit()
    conn.close()


@pytest.fixture
def collections_db(tmp_path):
    p = tmp_path / "collections.sqlite"
    _make_db_with_collections(str(p))
    return p


class TestListCollections:
    def test_returns_all_non_deleted(self, collections_db):
        cols = ZoteroClient(str(collections_db)).list_collections()
        names = {c["name"] for c in cols}
        assert names == {"Root", "Child A", "Child B", "Grandchild", "Orphan"}
        assert "Deleted" not in names

    def test_path_is_hierarchical(self, collections_db):
        cols = ZoteroClient(str(collections_db)).list_collections()
        by_name = {c["name"]: c for c in cols}
        assert by_name["Root"]["path"] == "Root"
        assert by_name["Child A"]["path"] == "Root / Child A"
        assert by_name["Grandchild"]["path"] == "Root / Child B / Grandchild"
        assert by_name["Orphan"]["path"] == "Orphan"

    def test_item_count_direct_only(self, collections_db):
        """item_count is direct members only (sub-collections not included)."""
        cols = ZoteroClient(str(collections_db)).list_collections()
        by_name = {c["name"]: c for c in cols}
        assert by_name["Root"]["item_count"] == 0
        assert by_name["Child A"]["item_count"] == 2
        assert by_name["Child B"]["item_count"] == 0
        assert by_name["Grandchild"]["item_count"] == 1
        assert by_name["Orphan"]["item_count"] == 1

    def test_item_count_excludes_deleted_items(self, collections_db):
        """Deleted item (201) is in 'Deleted' collection — but even if the
        collection existed, the item is filtered out. Add a dangling
        deleted-item membership to a live collection and confirm it's ignored."""
        conn = sqlite3.connect(str(collections_db))
        conn.execute("INSERT INTO collectionItems VALUES (2, 201)")
        conn.commit()
        conn.close()
        cols = ZoteroClient(str(collections_db)).list_collections()
        by_name = {c["name"]: c for c in cols}
        # Child A still reports 2, not 3.
        assert by_name["Child A"]["item_count"] == 2

    def test_parent_id_preserved(self, collections_db):
        cols = ZoteroClient(str(collections_db)).list_collections()
        by_name = {c["name"]: c for c in cols}
        assert by_name["Root"]["parent_id"] is None
        assert by_name["Child A"]["parent_id"] == 1
        assert by_name["Grandchild"]["parent_id"] == 3

    def test_missing_db_raises(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            ZoteroClient(str(tmp_path / "nope.sqlite")).list_collections()


class TestCollectDescendantIds:
    def test_root_returns_whole_tree(self, collections_db):
        conn = sqlite3.connect(str(collections_db))
        conn.row_factory = sqlite3.Row
        try:
            ids = ZoteroClient._collect_descendant_ids(conn, 1)
        finally:
            conn.close()
        assert set(ids) == {1, 2, 3, 4}

    def test_leaf_returns_self_only(self, collections_db):
        conn = sqlite3.connect(str(collections_db))
        conn.row_factory = sqlite3.Row
        try:
            ids = ZoteroClient._collect_descendant_ids(conn, 4)
        finally:
            conn.close()
        assert ids == [4]

    def test_mid_level_returns_subtree(self, collections_db):
        conn = sqlite3.connect(str(collections_db))
        conn.row_factory = sqlite3.Row
        try:
            ids = ZoteroClient._collect_descendant_ids(conn, 3)
        finally:
            conn.close()
        assert set(ids) == {3, 4}

    def test_unknown_id_returns_self_only(self, collections_db):
        conn = sqlite3.connect(str(collections_db))
        conn.row_factory = sqlite3.Row
        try:
            ids = ZoteroClient._collect_descendant_ids(conn, 999)
        finally:
            conn.close()
        assert ids == [999]
