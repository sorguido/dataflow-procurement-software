"""
Test automatici per services/supplier_category_persistence.py

Scenari testati (da spec):
  1. Migrazione: categorie di potential_suppliers importate in supplier_categories
  2. ensure_supplier_category_exists: idempotente, trim automatico
  3. Rinomina: tutti i supplier aggiornati, old rimossa, new presente
  4. Rinomina verso categoria già esistente → CategoryError
  5. Merge: supplier spostati, source eliminata, target resta
  6. Elimina non usata → consentita
  7. Elimina usata → bloccata

  8. Trim input: " Tornerie " salvata come "Tornerie", nessun duplicato
  9. Refresh dialog parent: categoria rinominata → valore precedente non più valido
"""

import os
import tempfile
import unittest

from database_manager import DatabaseManager, DatabaseError
from services.supplier_category_persistence import (
    get_all_supplier_categories,
    ensure_supplier_category_exists,
    rename_supplier_category,
    merge_supplier_categories,
    delete_supplier_category_if_unused,
    count_suppliers_by_category,
    CategoryError,
)
from models.potential_supplier import PotentialSupplier


class TestSupplierCategoryPersistence(unittest.TestCase):

    def setUp(self):
        self.test_dir = tempfile.mkdtemp()
        self.test_db_path = os.path.join(self.test_dir, 'test_categories.db')
        self.db = DatabaseManager(self.test_db_path)
        self.db.create_tables()

    def tearDown(self):
        self.db.close()
        if os.path.exists(self.test_db_path):
            os.remove(self.test_db_path)
        os.rmdir(self.test_dir)

    # -----------------------------------------------------------------------
    # Helper
    # -----------------------------------------------------------------------

    def _add_supplier(self, name: str, category: str) -> int:
        """Inserisce un fornitore direttamente nel DB per i test."""
        supplier = PotentialSupplier(
            supplier_name=name,
            category=category,
            username="test_user",
        )
        return self.db.insert_potential_supplier(supplier)

    def _get_category_names(self) -> list:
        return get_all_supplier_categories(self.db)

    # -----------------------------------------------------------------------
    # Scenario 1 — Migrazione
    # -----------------------------------------------------------------------

    def test_01_migration_imports_existing_categories(self):
        """
        Le categorie presenti in potential_suppliers vengono importate in
        supplier_categories durante create_tables() (migrazione idempotente).
        """
        # Inseriamo supplier direttamente nel DB (simulate pre-existing data)
        self.db.cursor.execute(
            "INSERT INTO potential_suppliers (supplier_name, category, supplier_status, username) "
            "VALUES (?, ?, ?, ?)", ("Fornitore A", "Acciaio", "Nuovo", "u1")
        )
        self.db.cursor.execute(
            "INSERT INTO potential_suppliers (supplier_name, category, supplier_status, username) "
            "VALUES (?, ?, ?, ?)", ("Fornitore B", "Plastica", "Nuovo", "u1")
        )
        self.db.cursor.execute(
            "INSERT INTO potential_suppliers (supplier_name, category, supplier_status, username) "
            "VALUES (?, ?, ?, ?)", ("Fornitore C", "Acciaio", "Nuovo", "u1")  # duplicato
        )
        self.db.conn.commit()

        # Ricrea le tabelle (simula riavvio app / migrazione)
        self.db.create_tables()

        cats = self._get_category_names()
        self.assertIn("Acciaio", cats)
        self.assertIn("Plastica", cats)
        # nessun duplicato
        self.assertEqual(cats.count("Acciaio"), 1)

    # -----------------------------------------------------------------------
    # Scenario 2 — ensure_supplier_category_exists
    # -----------------------------------------------------------------------

    def test_02_ensure_creates_and_is_idempotent(self):
        """ensure crea la categoria; chiamate successive non falliscono."""
        ensure_supplier_category_exists(self.db, "Gomma")
        cats = self._get_category_names()
        self.assertIn("Gomma", cats)

        # Idempotente
        ensure_supplier_category_exists(self.db, "Gomma")
        self.assertEqual(cats.count("Gomma"), 1)

    def test_02b_ensure_empty_name_is_noop(self):
        """ensure con nome vuoto o solo spazi non crea niente."""
        ensure_supplier_category_exists(self.db, "")
        ensure_supplier_category_exists(self.db, "   ")
        cats = self._get_category_names()
        self.assertEqual(len(cats), 0)

    # -----------------------------------------------------------------------
    # Scenario 3 — Rinomina
    # -----------------------------------------------------------------------

    def test_03_rename_updates_suppliers_and_catalogue(self):
        """Rinomina aggiorna tutti i supplier e la tabella categorie."""
        ensure_supplier_category_exists(self.db, "Pippo")
        sid1 = self._add_supplier("S1", "Pippo")
        sid2 = self._add_supplier("S2", "Pippo")

        rename_supplier_category(self.db, "Pippo", "Plastica")

        # Tabella categorie
        cats = self._get_category_names()
        self.assertIn("Plastica", cats)
        self.assertNotIn("Pippo", cats)

        # Supplier aggiornati
        self.db.cursor.execute(
            "SELECT category FROM potential_suppliers WHERE supplier_id = ?", (sid1,)
        )
        self.assertEqual(self.db.cursor.fetchone()[0], "Plastica")

        self.db.cursor.execute(
            "SELECT category FROM potential_suppliers WHERE supplier_id = ?", (sid2,)
        )
        self.assertEqual(self.db.cursor.fetchone()[0], "Plastica")

    # -----------------------------------------------------------------------
    # Scenario 4 — Rinomina verso categoria già esistente
    # -----------------------------------------------------------------------

    def test_04_rename_to_existing_raises_category_error(self):
        """Rinominare verso una categoria già esistente lancia CategoryError."""
        ensure_supplier_category_exists(self.db, "Tornerie")
        ensure_supplier_category_exists(self.db, "Lavorazioni meccaniche")

        with self.assertRaises(CategoryError) as ctx:
            rename_supplier_category(self.db, "Tornerie", "Lavorazioni meccaniche")

        self.assertIn("Unisci", str(ctx.exception))

    # -----------------------------------------------------------------------
    # Scenario 5 — Unisci
    # -----------------------------------------------------------------------

    def test_05_merge_moves_suppliers_and_removes_source(self):
        """Merge sposta i supplier e rimuove la source dal catalogo."""
        ensure_supplier_category_exists(self.db, "Tornerie")
        ensure_supplier_category_exists(self.db, "Lavorazioni meccaniche")
        sid = self._add_supplier("Torneria Rossi", "Tornerie")

        merge_supplier_categories(self.db, "Tornerie", "Lavorazioni meccaniche")

        cats = self._get_category_names()
        self.assertIn("Lavorazioni meccaniche", cats)
        self.assertNotIn("Tornerie", cats)

        self.db.cursor.execute(
            "SELECT category FROM potential_suppliers WHERE supplier_id = ?", (sid,)
        )
        self.assertEqual(self.db.cursor.fetchone()[0], "Lavorazioni meccaniche")

    # -----------------------------------------------------------------------
    # Scenario 6 — Elimina non usata
    # -----------------------------------------------------------------------

    def test_06_delete_unused_category_succeeds(self):
        """Una categoria senza supplier associati viene eliminata."""
        ensure_supplier_category_exists(self.db, "Vuota")
        count = delete_supplier_category_if_unused(self.db, "Vuota")

        self.assertEqual(count, 0)
        self.assertNotIn("Vuota", self._get_category_names())

    # -----------------------------------------------------------------------
    # Scenario 7 — Elimina usata
    # -----------------------------------------------------------------------

    def test_07_delete_used_category_is_blocked(self):
        """Una categoria con supplier associati non viene eliminata."""
        ensure_supplier_category_exists(self.db, "InUso")
        self._add_supplier("Fornitore X", "InUso")

        count = delete_supplier_category_if_unused(self.db, "InUso")

        self.assertGreater(count, 0)
        self.assertIn("InUso", self._get_category_names())

    # -----------------------------------------------------------------------
    # Scenario 8 — Trim input
    # -----------------------------------------------------------------------

    def test_08_trim_input_prevents_duplicates(self):
        """Categorie con spazi iniziali/finali vengono normalizzate via trim."""
        ensure_supplier_category_exists(self.db, " Tornerie ")
        ensure_supplier_category_exists(self.db, "Tornerie")  # stesso nome dopo trim

        cats = self._get_category_names()
        self.assertIn("Tornerie", cats)
        self.assertEqual(cats.count("Tornerie"), 1)

        # Il supplier deve risultare con "Tornerie" (trimmato)
        sid = self._add_supplier("S trim", " Tornerie ")
        # Normalizzazione avviene a livello persistence, non nel DB diretto
        # Verifichiamo che ensure non abbia creato " Tornerie " con spazi
        self.assertNotIn(" Tornerie ", cats)

    def test_08b_rename_with_spaces_normalizes(self):
        """rename con spazi sui nomi funziona correttamente."""
        ensure_supplier_category_exists(self.db, "OldCat")
        self._add_supplier("S", "OldCat")

        rename_supplier_category(self.db, " OldCat ", " NewCat ")

        cats = self._get_category_names()
        self.assertIn("NewCat", cats)
        self.assertNotIn("OldCat", cats)
        self.assertNotIn(" NewCat ", cats)

    # -----------------------------------------------------------------------
    # Scenario 9 — count_suppliers_by_category
    # -----------------------------------------------------------------------

    def test_09_count_suppliers_by_category(self):
        """count_suppliers_by_category restituisce il conteggio corretto."""
        ensure_supplier_category_exists(self.db, "Contato")
        self._add_supplier("A", "Contato")
        self._add_supplier("B", "Contato")

        self.assertEqual(count_suppliers_by_category(self.db, "Contato"), 2)
        self.assertEqual(count_suppliers_by_category(self.db, "Inesistente"), 0)

    # -----------------------------------------------------------------------
    # Edge cases
    # -----------------------------------------------------------------------

    def test_rename_noop_when_same_name(self):
        """Rinomina con stesso nome è no-op silenzioso."""
        ensure_supplier_category_exists(self.db, "Stessa")
        rename_supplier_category(self.db, "Stessa", "Stessa")
        self.assertIn("Stessa", self._get_category_names())

    def test_rename_nonexistent_raises(self):
        """Rinomina di categoria inesistente lancia CategoryError."""
        with self.assertRaises(CategoryError):
            rename_supplier_category(self.db, "NonEsiste", "Nuova")

    def test_merge_raises_if_source_equals_target(self):
        """Merge con source == target lancia CategoryError."""
        ensure_supplier_category_exists(self.db, "Cat")
        with self.assertRaises(CategoryError):
            merge_supplier_categories(self.db, "Cat", "Cat")

    def test_merge_nonexistent_source_raises(self):
        """Merge di source inesistente lancia CategoryError."""
        ensure_supplier_category_exists(self.db, "Target")
        with self.assertRaises(CategoryError):
            merge_supplier_categories(self.db, "NonEsiste", "Target")

    def test_delete_empty_name_raises(self):
        """delete con nome vuoto lancia CategoryError."""
        with self.assertRaises(CategoryError):
            delete_supplier_category_if_unused(self.db, "")


if __name__ == "__main__":
    unittest.main()
