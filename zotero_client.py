"""
zotero_client.py
Reads the local Zotero SQLite database and exposes library items
with full CSL JSON data needed to build Zotero field codes.
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

try:
    import requests
    _REQUESTS_OK = True
except ImportError:
    _REQUESTS_OK = False


# ---------------------------------------------------------------------------
# Zotero type → CSL type mapping
# ---------------------------------------------------------------------------

ZOTERO_TO_CSL_TYPE: Dict[str, str] = {
    "artwork": "graphic",
    "audioRecording": "song",
    "bill": "bill",
    "blogPost": "post-weblog",
    "book": "book",
    "bookSection": "chapter",
    "case": "legal_case",
    "computerProgram": "software",
    "conferencePaper": "paper-conference",
    "dictionaryEntry": "entry-dictionary",
    "document": "document",
    "email": "personal_communication",
    "encyclopediaArticle": "entry-encyclopedia",
    "film": "motion_picture",
    "forumPost": "post",
    "hearing": "hearing",
    "instantMessage": "personal_communication",
    "interview": "interview",
    "journalArticle": "article-journal",
    "letter": "personal_communication",
    "magazineArticle": "article-magazine",
    "manuscript": "manuscript",
    "map": "map",
    "newspaperArticle": "article-newspaper",
    "patent": "patent",
    "podcast": "broadcast",
    "preprint": "article",
    "presentation": "speech",
    "radioBroadcast": "broadcast",
    "report": "report",
    "statute": "legislation",
    "thesis": "thesis",
    "tvBroadcast": "broadcast",
    "videoRecording": "motion_picture",
    "webpage": "webpage",
}

# Zotero field name → CSL field name mapping (incomplete but covers common fields)
ZOTERO_TO_CSL_FIELD: Dict[str, str] = {
    "title": "title",
    "abstractNote": "abstract",
    "url": "URL",
    "accessDate": "accessed",
    "archive": "archive",
    "archiveLocation": "archive_location",
    "callNumber": "call-number",
    "date": "issued",
    "edition": "edition",
    "extra": "note",
    "ISBN": "ISBN",
    "ISSN": "ISSN",
    "issue": "issue",
    "journalAbbreviation": "journalAbbreviation",
    "language": "language",
    "DOI": "DOI",
    "numPages": "number-of-pages",
    "pages": "page",
    "place": "publisher-place",
    "publicationTitle": "container-title",
    "publisher": "publisher",
    "series": "collection-title",
    "seriesTitle": "collection-title",
    "seriesNumber": "collection-number",
    "shortTitle": "title-short",
    "university": "publisher",
    "volume": "volume",
    "number": "number",
    "patentNumber": "number",
    "reportNumber": "number",
    "reportType": "genre",
    "thesisType": "genre",
    "artworkSize": "dimensions",
    "bookTitle": "container-title",
    "conferenceName": "event",
    "websiteTitle": "container-title",
    "blogTitle": "container-title",
    "forumTitle": "container-title",
    "programTitle": "container-title",
    "episodeNumber": "number",
    "network": "publisher",
    "studio": "publisher",
    "label": "publisher",
    "distributor": "publisher",
    "videoRecordingFormat": "medium",
    "audioRecordingFormat": "medium",
    "medium": "medium",
    "runningTime": "dimensions",
    "scale": "scale",
    "mapType": "genre",
    "court": "authority",
    "docketNumber": "number",
    "firstPage": "page",
    "history": "references",
    "session": "chapter-number",
    "code": "container-title",
    "codeVolume": "volume",
    "codePages": "page",
    "legislativeBody": "authority",
    "statute": "title",
    "billNumber": "number",
    "system": "title-short",
    "company": "publisher",
    "programmingLanguage": "genre",
    "versionNumber": "version",
}


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class ZoteroItem:
    key: str
    item_type_zotero: str           # e.g. "journalArticle"
    item_type_csl: str              # e.g. "article-journal"
    title: str
    year: str                       # "2023" or "2023-04-15"
    authors: List[str]              # last names only, ordered
    library_id: int
    user_id: str                    # Zotero user ID string
    item_id: int = 0               # numeric SQLite itemID (for Zotero field codes)
    editors: List[str] = field(default_factory=list)  # editor last names
    csl_data: dict = field(default_factory=dict)  # full CSL JSON object

    @property
    def uri(self) -> str:
        if self.user_id:
            return f"http://zotero.org/users/{self.user_id}/items/{self.key}"
        return f"http://zotero.org/groups/0/items/{self.key}"


# ---------------------------------------------------------------------------
# Client
# ---------------------------------------------------------------------------

class ZoteroClient:

    @staticmethod
    def _find_default_db() -> Path:
        """Return the first existing Zotero SQLite path for this platform."""
        import sys
        home = Path.home()
        candidates: List[Path] = []
        if sys.platform == "win32":
            appdata = os.environ.get("APPDATA", "")
            if appdata:
                zp = Path(appdata) / "Zotero" / "Zotero" / "Profiles"
                if zp.is_dir():
                    for prof in sorted(zp.iterdir()):
                        candidates.append(prof / "zotero.sqlite")
        candidates.append(home / "Zotero" / "zotero.sqlite")        # macOS & common Linux
        if sys.platform == "linux":
            # Zotero 6 / standalone on Linux
            dot_zotero = home / ".zotero" / "zotero"
            if dot_zotero.is_dir():
                for prof in sorted(dot_zotero.iterdir()):
                    candidates.append(prof / "zotero.sqlite")
            # Snap / Flatpak variants
            candidates.append(home / "snap" / "zotero-snap" / "common" / "Zotero" / "zotero.sqlite")
        for p in candidates:
            if p.exists():
                return p
        # Fallback: return the most common path so the error message is helpful
        return home / "Zotero" / "zotero.sqlite"

    def __init__(self, db_path: Optional[str] = None):
        self.db_path = Path(db_path) if db_path else self._find_default_db()
        self._items: Optional[List[ZoteroItem]] = None
        self._user_id: Optional[str] = None

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------

    def ping_zotero(self) -> bool:
        """Return True if Zotero desktop is running (via connector endpoint)."""
        if not _REQUESTS_OK:
            return False
        try:
            r = requests.get(
                "http://localhost:23119/connector/ping",
                timeout=1.5,
                headers={"X-Zotero-Version": "5.0"},
            )
            return r.status_code == 200
        except Exception:
            return False

    def load_library(
        self,
        progress_cb=None,
        collection_id: Optional[int] = None,
    ) -> List[ZoteroItem]:
        """Load library items from the Zotero SQLite database.

        If *collection_id* is given, only items belonging to that collection
        or any of its descendant sub-collections are returned.
        """
        if not self.db_path.exists():
            raise FileNotFoundError(
                f"Zotero database not found at {self.db_path}.\n"
                "Please ensure Zotero is installed and the database path is correct."
            )

        # Copy to temp file to avoid locking the live database
        with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            shutil.copy2(str(self.db_path), tmp_path)
            items = self._read_database(
                tmp_path, progress_cb=progress_cb, collection_id=collection_id
            )
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

        self._items = items
        return items

    def list_collections(self) -> List[Dict]:
        """Return all non-trashed collections as a flat list of dicts with
        keys: ``id`` (int), ``name``, ``parent_id`` (int or None),
        ``library_id`` (int), ``path`` (``"Parent / Child"`` for display),
        ``item_count`` (number of direct items, not including sub-collections).
        """
        if not self.db_path.exists():
            raise FileNotFoundError(
                f"Zotero database not found at {self.db_path}."
            )
        with tempfile.NamedTemporaryFile(suffix=".sqlite", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            shutil.copy2(str(self.db_path), tmp_path)
            conn = sqlite3.connect(tmp_path)
            conn.row_factory = sqlite3.Row
            try:
                rows = conn.execute(
                    """
                    SELECT c.collectionID as id, c.collectionName as name,
                           c.parentCollectionID as parent_id,
                           c.libraryID as library_id
                    FROM collections c
                    WHERE c.collectionID NOT IN (
                        SELECT collectionID FROM deletedCollections
                    )
                    ORDER BY c.libraryID, c.collectionName
                    """
                ).fetchall()
                counts = {
                    r["collectionID"]: r["n"]
                    for r in conn.execute(
                        """
                        SELECT ci.collectionID, COUNT(*) as n
                        FROM collectionItems ci
                        JOIN items i ON ci.itemID = i.itemID
                        WHERE i.itemID NOT IN (SELECT itemID FROM deletedItems)
                        GROUP BY ci.collectionID
                        """
                    )
                }
            finally:
                conn.close()
        finally:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass

        by_id = {r["id"]: dict(r) for r in rows}
        # Build a "Parent / Child / Grandchild" path for each collection.
        def _path(cid: int) -> str:
            names: List[str] = []
            seen = set()
            while cid in by_id and cid not in seen:
                seen.add(cid)
                rec = by_id[cid]
                names.append(rec["name"])
                cid = rec["parent_id"]
            return " / ".join(reversed(names))

        out: List[Dict] = []
        for r in rows:
            out.append({
                "id": r["id"],
                "name": r["name"],
                "parent_id": r["parent_id"],
                "library_id": r["library_id"],
                "path": _path(r["id"]),
                "item_count": counts.get(r["id"], 0),
            })
        out.sort(key=lambda c: (c["library_id"], c["path"].lower()))
        return out

    @property
    def items(self) -> List[ZoteroItem]:
        if self._items is None:
            raise RuntimeError("Library not loaded. Call load_library() first.")
        return self._items

    # ------------------------------------------------------------------
    # Database reading
    # ------------------------------------------------------------------

    def _read_database(
        self,
        db_path: str,
        progress_cb=None,
        collection_id: Optional[int] = None,
    ) -> List[ZoteroItem]:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        try:
            user_id = self._get_user_id(conn)
            self._user_id = user_id
            items = self._fetch_items(
                conn, user_id,
                progress_cb=progress_cb,
                collection_id=collection_id,
            )
        finally:
            conn.close()
        return items

    @staticmethod
    def _collect_descendant_ids(
        conn: sqlite3.Connection, root_id: int
    ) -> List[int]:
        """Return *root_id* plus every descendant collectionID."""
        rows = conn.execute(
            "SELECT collectionID, parentCollectionID FROM collections"
        ).fetchall()
        children: Dict[int, List[int]] = {}
        for r in rows:
            pid = r["parentCollectionID"]
            if pid is not None:
                children.setdefault(pid, []).append(r["collectionID"])
        result = [root_id]
        stack = [root_id]
        while stack:
            cur = stack.pop()
            for child in children.get(cur, []):
                result.append(child)
                stack.append(child)
        return result

    @staticmethod
    def _get_user_id(conn: sqlite3.Connection) -> str:
        """Extract Zotero user ID from settings table."""
        try:
            row = conn.execute(
                "SELECT value FROM settings WHERE setting='account' AND key='userID'"
            ).fetchone()
            if row:
                return str(row["value"])
        except sqlite3.OperationalError:
            pass
        # Fallback: try to get from library
        try:
            row = conn.execute(
                "SELECT libraryID FROM libraries WHERE type='user' LIMIT 1"
            ).fetchone()
            if row:
                return str(row["libraryID"])
        except sqlite3.OperationalError:
            pass
        return "0"

    def _fetch_items(
        self,
        conn: sqlite3.Connection,
        user_id: str,
        progress_cb=None,
        collection_id: Optional[int] = None,
    ) -> List[ZoteroItem]:
        # Get item type IDs to exclude (attachments=14, notes=1)
        excluded_type_ids = self._get_excluded_type_ids(conn)

        # If a collection was requested, restrict to items in that collection
        # or any of its descendants.
        collection_clause = ""
        params: tuple = ()
        if collection_id is not None:
            ids = self._collect_descendant_ids(conn, collection_id)
            placeholders = ",".join("?" for _ in ids)
            collection_clause = (
                f" AND i.itemID IN (SELECT itemID FROM collectionItems "
                f"WHERE collectionID IN ({placeholders}))"
            )
            params = tuple(ids)

        # Get all non-deleted, non-attachment items
        rows = conn.execute(
            """
            SELECT i.itemID, i.key, i.libraryID,
                   it.typeName as itemTypeName
            FROM items i
            JOIN itemTypes it ON i.itemTypeID = it.itemTypeID
            WHERE i.itemTypeID NOT IN ({excluded})
              AND i.itemID NOT IN (
                  SELECT itemID FROM deletedItems
              )
              {collection_clause}
            ORDER BY i.dateAdded DESC
            """.format(
                excluded=",".join(str(x) for x in excluded_type_ids),
                collection_clause=collection_clause,
            ),
            params,
        ).fetchall()

        if not rows:
            return []

        # Build fieldID → fieldName lookup
        field_map = {
            r["fieldID"]: r["fieldName"]
            for r in conn.execute("SELECT fieldID, fieldName FROM fields")
        }

        # Build creatorTypeID → typeName lookup
        creator_type_map = {
            r["creatorTypeID"]: r["creatorType"]
            for r in conn.execute(
                "SELECT creatorTypeID, creatorType FROM creatorTypes"
            )
        }

        items: List[ZoteroItem] = []
        total = len(rows)

        for i, row in enumerate(rows):
            if progress_cb and i % 50 == 0:
                progress_cb(i, total)

            item_id = row["itemID"]
            item_key = row["key"]
            library_id = row["libraryID"]
            type_name = row["itemTypeName"]
            csl_type = ZOTERO_TO_CSL_TYPE.get(type_name, "document")

            # Fetch field data
            field_rows = conn.execute(
                """
                SELECT id.fieldID, idv.value
                FROM itemData id
                JOIN itemDataValues idv ON id.valueID = idv.valueID
                WHERE id.itemID = ?
                """,
                (item_id,),
            ).fetchall()

            fields: Dict[str, str] = {}
            for fr in field_rows:
                fname = field_map.get(fr["fieldID"], "")
                if fname:
                    fields[fname] = fr["value"]

            title = fields.get("title", "")
            date_raw = fields.get("date", "")
            year = self._extract_year(date_raw)

            # Fetch creators
            creator_rows = conn.execute(
                """
                SELECT c.lastName, c.firstName, ic.creatorTypeID, ic.orderIndex
                FROM itemCreators ic
                JOIN creators c ON ic.creatorID = c.creatorID
                WHERE ic.itemID = ?
                ORDER BY ic.orderIndex
                """,
                (item_id,),
            ).fetchall()

            # Build authors list (last names) and CSL author array
            author_last_names: List[str] = []
            editor_last_names: List[str] = []
            csl_authors: List[dict] = []
            csl_editors: List[dict] = []

            for cr in creator_rows:
                ctype = creator_type_map.get(cr["creatorTypeID"], "author")
                person = {}
                if cr["lastName"]:
                    person["family"] = cr["lastName"]
                if cr["firstName"]:
                    person["given"] = cr["firstName"]
                if not person:
                    continue

                if ctype == "author":
                    author_last_names.append(cr["lastName"] or cr["firstName"])
                    csl_authors.append(person)
                elif ctype == "editor":
                    editor_last_names.append(cr["lastName"] or cr["firstName"])
                    csl_editors.append(person)

            # Build CSL data dict
            # id MUST be a numeric integer — Zotero drops field codes with string ids on refresh
            csl_data: dict = {
                "id": item_id,
                "type": csl_type,
            }

            if title:
                csl_data["title"] = title
            if csl_authors:
                csl_data["author"] = csl_authors
            if csl_editors:
                csl_data["editor"] = csl_editors
            if year:
                csl_data["issued"] = {"date-parts": [[int(year)]]}
            if date_raw and year not in date_raw:
                csl_data["issued"] = {"raw": date_raw}
            elif year:
                try:
                    csl_data["issued"] = {"date-parts": [[int(year[:4])]]}
                except ValueError:
                    csl_data["issued"] = {"raw": date_raw}

            # Map other Zotero fields to CSL
            for zfield, value in fields.items():
                if zfield in ("title", "date"):
                    continue
                csl_field = ZOTERO_TO_CSL_FIELD.get(zfield)
                if csl_field and csl_field not in csl_data:
                    csl_data[csl_field] = value

            zitem = ZoteroItem(
                key=item_key,
                item_type_zotero=type_name,
                item_type_csl=csl_type,
                title=title,
                year=year,
                authors=author_last_names,
                library_id=library_id,
                user_id=user_id,
                item_id=item_id,
                editors=editor_last_names,
                csl_data=csl_data,
            )
            items.append(zitem)

        if progress_cb:
            progress_cb(total, total)

        return items

    @staticmethod
    def _get_excluded_type_ids(conn: sqlite3.Connection) -> List[int]:
        """Return itemTypeIDs for attachment (14) and note (1)."""
        excluded_names = {"attachment", "note"}
        try:
            rows = conn.execute(
                "SELECT itemTypeID, typeName FROM itemTypes"
            ).fetchall()
            return [r["itemTypeID"] for r in rows if r["typeName"] in excluded_names]
        except sqlite3.OperationalError:
            return [1, 14]

    @staticmethod
    def _extract_year(date_str: str) -> str:
        if not date_str:
            return ""
        m = re.search(r"\b(\d{4})\b", date_str)
        return m.group(1) if m else ""
