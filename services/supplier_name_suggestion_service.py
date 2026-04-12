"""
Servizio read-only per suggerimento nomi fornitore.

Sorgenti v1:
- richiesta_fornitori (RFQ) [prioritaria]
- potential_suppliers (Derisking)
"""

from __future__ import annotations

from dataclasses import dataclass, field

from database_manager import DatabaseManager
from utils.supplier_name_normalization import (
    build_supplier_match_keys,
    normalize_supplier_name_for_match,
)


@dataclass
class _NameStats:
    rfq_count: int = 0
    potential_count: int = 0


@dataclass
class _AggregateEntry:
    strict_key: str
    base_key: str
    has_suffix: bool
    aliases: dict[str, _NameStats] = field(default_factory=dict)

    @property
    def rfq_count(self) -> int:
        return sum(v.rfq_count for v in self.aliases.values())

    @property
    def potential_count(self) -> int:
        return sum(v.potential_count for v in self.aliases.values())

    def best_display_name(self) -> str:
        if not self.aliases:
            return ""
        ranked = sorted(
            self.aliases.items(),
            key=lambda item: (
                -item[1].rfq_count,
                -item[1].potential_count,
                len(item[0]),
                item[0].lower(),
            ),
        )
        return ranked[0][0]


class SupplierNameSuggestionIndex:
    """Indice in-memory per suggerimenti e warning soft."""

    def __init__(self, entries: list[_AggregateEntry]):
        self._entries = entries
        self._by_strict = {e.strict_key: e for e in entries}
        self._by_base: dict[str, list[_AggregateEntry]] = {}
        for entry in entries:
            self._by_base.setdefault(entry.base_key, []).append(entry)

    def suggest(self, query: str, limit: int = 8) -> list[str]:
        query_key = normalize_supplier_name_for_match(query)
        if not query_key:
            return []

        matches: list[tuple[tuple, str]] = []
        for entry in self._entries:
            if query_key not in entry.strict_key:
                continue
            display = entry.best_display_name()
            if not display:
                continue
            rank = (
                0 if entry.strict_key.startswith(query_key) else 1,
                -entry.rfq_count,
                -entry.potential_count,
                display.lower(),
            )
            matches.append((rank, display))

        matches.sort(key=lambda item: item[0])

        out: list[str] = []
        seen = set()
        for _, display in matches:
            key = display.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(display)
            if len(out) >= limit:
                break
        return out

    def get_soft_duplicate_candidates(self, input_name: str, limit: int = 3) -> list[str]:
        """
        Ritorna possibili duplicati semantici per warning non bloccante.

        Strategia conservativa:
        - match su strict_key
        - fallback su base_key solo se entrano in gioco suffissi legali
        """
        strict_key, base_key, has_suffix = build_supplier_match_keys(input_name)
        if not strict_key:
            return []

        candidates: dict[str, tuple[int, int]] = {}

        strict_entry = self._by_strict.get(strict_key)
        if strict_entry:
            for alias, stats in strict_entry.aliases.items():
                if alias.strip().lower() == input_name.strip().lower():
                    continue
                candidates[alias] = (stats.rfq_count, stats.potential_count)

        if not candidates and base_key:
            for entry in self._by_base.get(base_key, []):
                # Base-key warning solo se almeno uno dei due usa suffisso legale
                if not (has_suffix or entry.has_suffix):
                    continue
                for alias, stats in entry.aliases.items():
                    if alias.strip().lower() == input_name.strip().lower():
                        continue
                    candidates[alias] = (
                        max(candidates.get(alias, (0, 0))[0], stats.rfq_count),
                        max(candidates.get(alias, (0, 0))[1], stats.potential_count),
                    )

        ranked = sorted(
            candidates.items(),
            key=lambda item: (-item[1][0], -item[1][1], len(item[0]), item[0].lower()),
        )
        return [name for name, _ in ranked[:limit]]


class SupplierNameSuggestionService:
    """Factory indice suggerimenti da DB."""

    @staticmethod
    def build_index(db_path: str) -> SupplierNameSuggestionIndex:
        entries: dict[str, _AggregateEntry] = {}

        with DatabaseManager(db_path, read_only=True) as db:
            cursor = db.cursor

            cursor.execute(
                """
                SELECT TRIM(nome_fornitore) AS supplier_name, COUNT(*) AS freq
                FROM richiesta_fornitori
                WHERE nome_fornitore IS NOT NULL AND TRIM(nome_fornitore) != ''
                GROUP BY TRIM(nome_fornitore)
                """
            )
            rfq_rows = cursor.fetchall()

            cursor.execute(
                """
                SELECT TRIM(supplier_name) AS supplier_name, COUNT(*) AS freq
                FROM potential_suppliers
                WHERE supplier_name IS NOT NULL AND TRIM(supplier_name) != ''
                GROUP BY TRIM(supplier_name)
                """
            )
            potential_rows = cursor.fetchall()

        for name, freq in rfq_rows:
            SupplierNameSuggestionService._accumulate(entries, name, int(freq or 0), 0)
        for name, freq in potential_rows:
            SupplierNameSuggestionService._accumulate(entries, name, 0, int(freq or 0))

        all_entries = list(entries.values())
        all_entries.sort(
            key=lambda e: (
                -e.rfq_count,
                -e.potential_count,
                e.best_display_name().lower(),
            )
        )
        return SupplierNameSuggestionIndex(all_entries)

    @staticmethod
    def _accumulate(
        entries: dict[str, _AggregateEntry],
        alias_name: str,
        rfq_count: int,
        potential_count: int,
    ) -> None:
        clean_alias = (alias_name or "").strip()
        strict_key, base_key, has_suffix = build_supplier_match_keys(clean_alias)
        if not strict_key:
            return

        entry = entries.get(strict_key)
        if entry is None:
            entry = _AggregateEntry(
                strict_key=strict_key,
                base_key=base_key,
                has_suffix=has_suffix,
            )
            entries[strict_key] = entry

        stats = entry.aliases.get(clean_alias)
        if stats is None:
            stats = _NameStats()
            entry.aliases[clean_alias] = stats
        stats.rfq_count += max(0, rfq_count)
        stats.potential_count += max(0, potential_count)
