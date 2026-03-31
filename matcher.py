"""
matcher.py
Fuzzy-matches extracted citations against Zotero library items.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple

from rapidfuzz import fuzz

from citation_parser import Citation
from zotero_client import ZoteroItem

# Nobiliary particles — stripped before fuzzy author comparison so that
# "Le Velly" and "Velly", or "van Giesen" and "Giesen", compare equally.
_PARTICLES = frozenset(
    "van von de du da di del della der den het dos das do le la ten ter te".split()
)

def _strip_particles(name: str) -> str:
    """Remove leading nobiliary particles from a name.
    'Le Velly' → 'Velly', 'van der Berg' → 'Berg', 'Smith' → 'Smith'."""
    words = name.split()
    while len(words) > 1 and words[0].lower() in _PARTICLES:
        words.pop(0)
    return " ".join(words)


# ---------------------------------------------------------------------------
# Result model
# ---------------------------------------------------------------------------

@dataclass
class MatchResult:
    citation: Citation
    zotero_item: Optional[ZoteroItem]
    confidence: float       # 0.0 – 1.0
    matched: bool
    is_suggestion: bool = False   # True = year within ±1, needs user confirmation
    suggestion_year: str = ""     # the actual year found in Zotero for this suggestion
    is_ambiguous: bool = False    # True = multiple library items match equally well
    candidates: List = field(default_factory=list)  # [(ZoteroItem, float), ...] sorted desc

    def summary(self) -> str:
        if self.is_suggestion and self.zotero_item:
            z = self.zotero_item
            return (
                f"? {self.citation.display()} "
                f"→ [{z.key}] {'; '.join(z.authors[:2])} ({z.year}) — {z.title[:50]}… "
                f"(year differs: doc={self.citation.year}, lib={z.year})"
            )
        if not self.matched or self.zotero_item is None:
            return f"✗ {self.citation.display()} — no match found"
        z = self.zotero_item
        authors_str = "; ".join(z.authors[:3])
        if len(z.authors) > 3:
            authors_str += " et al."
        return (
            f"✓ {self.citation.display()} "
            f"→ [{z.key}] {authors_str} ({z.year}) — {z.title[:60]}… "
            f"({self.confidence:.0%})"
        )


# ---------------------------------------------------------------------------
# Matcher
# ---------------------------------------------------------------------------

AUTHOR_THRESHOLD = 82   # minimum fuzz.ratio score for first-author match
TITLE_THRESHOLD = 68    # minimum token_sort_ratio for title match
CANDIDATE_THRESHOLD = 60  # lower bound for showing possible-match candidates


class CitationMatcher:
    def __init__(
        self,
        library: List[ZoteroItem],
        author_threshold: int = AUTHOR_THRESHOLD,
        candidate_threshold: int = CANDIDATE_THRESHOLD,
        year_tolerance: int = 1,
    ):
        self.library = library
        self._author_threshold    = author_threshold
        self._candidate_threshold = candidate_threshold
        self._year_tolerance      = year_tolerance
        # Build year-indexed lookup for fast pre-filtering
        self._by_year: dict[str, List[ZoteroItem]] = {}
        for item in library:
            y = item.year[:4] if item.year else ""
            self._by_year.setdefault(y, []).append(item)

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def match_all(
        self, citations: List[Citation], progress_cb=None
    ) -> List[MatchResult]:
        results: List[MatchResult] = []
        total = len(citations)
        for i, cit in enumerate(citations):
            if progress_cb:
                progress_cb(i, total)
            result = self.match_one(cit)
            results.append(result)
        if progress_cb:
            progress_cb(total, total)
        return results

    def match_one(self, citation: Citation) -> MatchResult:
        if citation.style == "author-year":
            all_matches = self._match_author_year_all(citation)

            if len(all_matches) >= 2:
                # Multiple items above threshold — try to disambiguate
                # using the bibliography/reference-list entry text.
                all_matches = self._reorder_by_ref_text(all_matches, citation)
                best_item, best_conf = all_matches[0]
                return MatchResult(citation=citation, zotero_item=best_item,
                                   confidence=best_conf, matched=False,
                                   is_ambiguous=True, candidates=all_matches)

            if len(all_matches) == 1:
                item, conf = all_matches[0]
                return MatchResult(citation=citation, zotero_item=item,
                                   confidence=conf, matched=True)

            # No exact match — try year ±tolerance as a suggestion
            sug_item, sug_conf, sug_year = self._try_year_tolerance(citation, self._year_tolerance)
            if sug_item is not None:
                return MatchResult(citation=citation, zotero_item=sug_item,
                                   confidence=sug_conf, matched=False,
                                   is_suggestion=True, suggestion_year=sug_year)

            # Still no match — look for lower-confidence candidates so the
            # user can pick manually rather than seeing a bare "unmatched".
            low_candidates = self._find_low_candidates(citation)
            if low_candidates:
                low_candidates = self._reorder_by_ref_text(low_candidates, citation)
                best_item, best_conf = low_candidates[0]
                return MatchResult(citation=citation, zotero_item=best_item,
                                   confidence=best_conf, matched=False,
                                   is_ambiguous=True, candidates=low_candidates)

            return MatchResult(citation=citation, zotero_item=None,
                               confidence=0.0, matched=False)
        else:
            item, conf = self._match_numbered(citation)
            if item is not None:
                return MatchResult(citation=citation, zotero_item=item,
                                   confidence=conf, matched=True)
            return MatchResult(citation=citation, zotero_item=None,
                               confidence=conf, matched=False)

    # ------------------------------------------------------------------
    # Reference-text disambiguation
    # ------------------------------------------------------------------

    @staticmethod
    def _reorder_by_ref_text(
        candidates: List[Tuple[ZoteroItem, float]],
        citation: Citation,
    ) -> List[Tuple[ZoteroItem, float]]:
        """If the citation has a bibliography entry (ref_text), compare each
        candidate's title against it and promote the best title match to the
        top of the list.  The remaining candidates keep their original order."""
        ref_text = getattr(citation, "ref_text", "") or ""
        if not ref_text or len(candidates) < 2:
            return candidates

        ref_norm = re.sub(r"[^\w\s]", " ", ref_text.lower())

        best_idx = 0
        best_title_score = 0.0

        for i, (item, _conf) in enumerate(candidates):
            if not item.title:
                continue
            title_norm = re.sub(r"[^\w\s]", " ", item.title.lower())
            score = fuzz.token_sort_ratio(title_norm, ref_norm) / 100.0
            if score > best_title_score:
                best_title_score = score
                best_idx = i

        # Only promote if the best title match is meaningfully better
        # than the runner-up (avoids random reshuffling when all are similar).
        if best_idx != 0 and best_title_score > 0.3:
            runner_up = 0.0
            for i, (item, _) in enumerate(candidates):
                if i == best_idx or not item.title:
                    continue
                title_norm = re.sub(r"[^\w\s]", " ", item.title.lower())
                s = fuzz.token_sort_ratio(title_norm, ref_norm) / 100.0
                if s > runner_up:
                    runner_up = s
            # Require a clear margin (>10 pp) before reordering
            if best_title_score - runner_up > 0.10:
                promoted = candidates[best_idx]
                reordered = [promoted] + candidates[:best_idx] + candidates[best_idx + 1:]
                return reordered

        return candidates

    # ------------------------------------------------------------------
    # Author-year matching
    # ------------------------------------------------------------------

    @staticmethod
    def _is_numeric_year(year: str) -> bool:
        """True if year is a standard 4-digit year (e.g. '2023'), False for
        non-numeric tokens like 'forthcoming', 'in press', 'n.d.'."""
        return bool(re.match(r"\d{4}", year))

    def _match_author_year_all(
        self, citation: Citation, year_override: str = ""
    ) -> List[Tuple[ZoteroItem, float]]:
        """Return ALL library items above threshold, sorted by score descending."""
        raw_year = year_override or citation.year
        year = raw_year[:4] if self._is_numeric_year(raw_year) else ""
        if not citation.authors:
            return []

        first_author = citation.authors[0].lower()
        first_author_stripped = _strip_particles(citation.authors[0]).lower()

        # If the first author contains an all-caps word (acronym like IPCC,
        # WHO, FAO), also try matching with just the acronym.  This handles
        # citations like "(Intergovernmental Panel on Climate Change (IPCC) 2022)".
        author_variants = [first_author]
        if first_author_stripped != first_author:
            author_variants.append(first_author_stripped)
        acronyms = re.findall(r'\b[A-Z]{2,}\b', citation.authors[0])
        for acr in acronyms:
            acr_lower = acr.lower()
            if acr_lower not in author_variants:
                author_variants.append(acr_lower)

        # Pre-filter by year
        candidates = self._by_year.get(year, [])
        if not candidates:
            candidates = self.library

        threshold = self._author_threshold / 100.0
        results: List[Tuple[ZoteroItem, float]] = []
        seen_keys: set = set()

        for item in candidates:
            if item.key in seen_keys:
                continue
            # Use authors, falling back to editors (e.g. "IPCC (ed.) (2022)")
            item_names = item.authors if item.authors else item.editors
            if not item_names:
                continue
            # "et al." in the citation means the source has multiple additional
            # authors beyond the first.  Only consider library items with 3+
            # authors (first author + more than one additional).
            if citation.has_et_al and len(item_names) < 3:
                continue
            item_first = item_names[0].lower()
            item_first_stripped = _strip_particles(item_names[0]).lower()
            if not item_first:
                continue

            # Try each author variant (full name, then acronyms) against
            # both the full and particle-stripped item name.
            best_score = 0.0
            item_variants = [item_first]
            if item_first_stripped != item_first:
                item_variants.append(item_first_stripped)

            for variant in author_variants:
                for iv in item_variants:
                    if iv[0] != variant[0]:
                        continue
                    score = fuzz.ratio(variant, iv) / 100.0

                    if len(citation.authors) > 1:
                        if len(item_names) > 1:
                            cit_second = _strip_particles(citation.authors[1]).lower()
                            item_second = _strip_particles(item_names[1]).lower()
                            second_score = fuzz.ratio(cit_second, item_second) / 100.0
                            score = 0.7 * score + 0.3 * second_score
                        else:
                            # Citation has multiple authors but item has only one —
                            # penalise so that a matching multi-author item ranks higher.
                            score *= 0.85

                    if item.year[:4] != year and year:
                        score *= 0.5

                    best_score = max(best_score, score)

            if best_score >= threshold:
                results.append((item, best_score))
                seen_keys.add(item.key)

        results.sort(key=lambda x: x[1], reverse=True)
        return results

    def _match_author_year(
        self, citation: Citation, year_override: str = ""
    ) -> Tuple[Optional[ZoteroItem], float]:
        """Return the single best match above threshold (used by year-tolerance fallback)."""
        matches = self._match_author_year_all(citation, year_override)
        if matches:
            return matches[0]
        return None, 0.0

    def _find_low_candidates(
        self, citation: Citation
    ) -> List[Tuple[ZoteroItem, float]]:
        """
        For citations that produced no match above AUTHOR_THRESHOLD, return up
        to 5 items scoring between CANDIDATE_THRESHOLD and AUTHOR_THRESHOLD.
        Searches the same year first, then ±1 year with a small score penalty,
        so results stay year-relevant.
        """
        year = citation.year[:4] if self._is_numeric_year(citation.year) else ""
        if not citation.authors:
            return []

        first_author = citation.authors[0].lower()
        first_author_stripped = _strip_particles(citation.authors[0]).lower()
        low  = self._candidate_threshold / 100.0
        high = self._author_threshold   / 100.0

        results: List[Tuple[ZoteroItem, float]] = []
        seen: set = set()

        years_to_search = [year]
        if year:
            try:
                base = int(year)
                years_to_search += [str(base - 1), str(base + 1)]
            except ValueError:
                pass

        for search_year in (years_to_search if year else [""]):
            penalty = 1.0 if search_year == year else 0.9
            pool = self._by_year.get(search_year, []) if search_year else self.library
            for item in pool:
                item_names = item.authors if item.authors else item.editors
                if item.key in seen or not item_names:
                    continue
                if citation.has_et_al and len(item_names) < 3:
                    continue
                item_first = item_names[0].lower()
                item_first_stripped = _strip_particles(item_names[0]).lower()
                if not item_first:
                    continue
                # Compare with and without particles
                best_score = 0.0
                for cv in (first_author, first_author_stripped):
                    for iv in (item_first, item_first_stripped):
                        if iv[0] != cv[0]:
                            continue
                        s = fuzz.ratio(cv, iv) / 100.0 * penalty
                        if len(citation.authors) > 1 and len(item_names) > 1:
                            cit_second = _strip_particles(citation.authors[1]).lower()
                            item_second = _strip_particles(item_names[1]).lower()
                            second_score = fuzz.ratio(cit_second, item_second) / 100.0
                            s = 0.7 * s + 0.3 * second_score * penalty
                        best_score = max(best_score, s)
                if low <= best_score < high:
                    results.append((item, best_score))
                    seen.add(item.key)

        results.sort(key=lambda x: x[1], reverse=True)
        return results[:5]

    def _try_year_tolerance(
        self, citation: Citation, tolerance: int = 1
    ) -> Tuple[Optional[ZoteroItem], float, str]:
        """Try matching with year ±tolerance. Returns (item, confidence, found_year)."""
        if not citation.year:
            return None, 0.0, ""
        try:
            base_year = int(citation.year[:4])
        except ValueError:
            return None, 0.0, ""

        best_item: Optional[ZoteroItem] = None
        best_score = 0.0
        best_year = ""

        for delta in range(1, tolerance + 1):
            for sign in (-1, 1):
                alt_year = str(base_year + sign * delta)
                item, score = self._match_author_year(citation, year_override=alt_year)
                if item is not None and score > best_score:
                    best_score = score
                    best_item = item
                    best_year = alt_year

        if best_item is not None:
            # Apply a confidence penalty to reflect the year uncertainty
            return best_item, best_score * 0.80, best_year
        return None, 0.0, ""

    # ------------------------------------------------------------------
    # Numbered matching
    # ------------------------------------------------------------------

    def _match_numbered(
        self, citation: Citation
    ) -> Tuple[Optional[ZoteroItem], float]:
        # If we have author + year from the reference list, try author-year first
        if citation.authors and citation.year:
            item, conf = self._match_author_year(citation)
            if item is not None:
                return item, conf

        # Fallback: title fuzzy match against all items
        ref_text = citation.raw_text
        if not ref_text or ref_text.startswith("["):
            return None, 0.0

        query = re.sub(r"[^\w\s]", " ", ref_text.lower())

        best_item: Optional[ZoteroItem] = None
        best_score = 0.0

        for item in self.library:
            if not item.title:
                continue
            title_norm = re.sub(r"[^\w\s]", " ", item.title.lower())
            score = fuzz.token_sort_ratio(query, title_norm) / 100.0

            if citation.year and item.year[:4] == citation.year[:4]:
                score = min(1.0, score * 1.15)

            if score > best_score:
                best_score = score
                best_item = item

        threshold = TITLE_THRESHOLD / 100.0
        if best_score >= threshold:
            return best_item, best_score
        return None, best_score
