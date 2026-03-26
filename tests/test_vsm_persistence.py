"""
Test suite per VSM Persistence Layer

Verifica il pattern DELETE-REGENERATE-SAVE e l'idempotenza delle operazioni.
"""

import unittest
import os
import tempfile
from datetime import datetime, timedelta

from database_manager import DatabaseManager
from models.vsm_event import VSMEvent
from models.vsm_impact import VSMImpact
from services.vsm_persistence import (
    save_event_with_impacts,
    update_event_with_impacts,
    delete_event_and_impacts,
    get_event_with_impacts,
    VSMError
)


class TestVSMPersistence(unittest.TestCase):
    """Test per il layer di persistenza VSM"""
    
    def setUp(self):
        """Crea database in-memory per ogni test"""
        # Usa database temporaneo su disco (SQLite :memory: non funziona bene con multithread)
        self.db_file = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        self.db_file.close()
        self.db_manager = DatabaseManager(self.db_file.name)
        
        # Crea tabelle utenti (richiesta da FK)
        self.db_manager.cursor.execute('''
            CREATE TABLE IF NOT EXISTS utenti (
                username TEXT PRIMARY KEY,
                password TEXT
            )
        ''')
        self.db_manager.cursor.execute("INSERT INTO utenti (username, password) VALUES ('test_user', 'pwd')")
        self.db_manager.conn.commit()
        
        # Crea tabelle VSM chiamando il metodo del DatabaseManager
        self.db_manager.create_tables()
    
    def tearDown(self):
        """Chiude e elimina database temporaneo"""
        if hasattr(self, 'db_manager'):
            self.db_manager.close()
        if hasattr(self, 'db_file'):
            try:
                os.unlink(self.db_file.name)
            except:
                pass
    
    def _create_test_event_repetitive(self):
        """Helper: crea evento ripetitivo di test"""
        return VSMEvent(
            id=None,
            username='test_user',
            event_date=datetime(2024, 6, 15),  # 15 giugno 2024
            buyer='Mario Rossi',
            event_type='Saving',
            action='Negoziazione',
            description='Test saving ripetitivo',
            reference='RFQ-001',
            importo_bdg=10000.0,
            importo_negoziato=8000.0,
            quantita_annua=1.0,
            percent_realizzo=80.0,
            driver='Prezzo',
            spending_annuo=96000.0,
            opex_ripetitivo=True,  # Ripetitivo: 24 impatti con pro-rata
            note='Test note'
        )
    
    def _create_test_event_one_shot(self):
        """Helper: crea evento one-shot di test"""
        return VSMEvent(
            id=None,
            username='test_user',
            event_date=datetime(2024, 6, 15),
            buyer='Luigi Verdi',
            event_type='Cost Avoidance',
            action='Negoziazione',
            description='Test cost avoidance one-shot',
            reference='PO-002',
            importo_richiesto_iniziale=15000.0,
            importo_negoziato=10000.0,
            quantita_annua=1.0,
            percent_realizzo=100.0,
            driver='Prezzo',
            spending_annuo=120000.0,
            opex_ripetitivo=False,  # One-shot: 1 solo impatto
            note='Test note one-shot'
        )
    
    def test_save_event_with_impacts(self):
        """Test salvataggio evento + generazione impatti"""
        event = self._create_test_event_repetitive()
        
        # Salva evento
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Verifica evento salvato
        self.assertIsNotNone(event_id)
        self.assertGreater(event_id, 0)
        
        # Verifica evento recuperabile
        saved_event = self.db_manager.get_vsm_event_by_id(event_id)
        self.assertIsNotNone(saved_event)
        self.assertEqual(saved_event.username, 'test_user')
        self.assertEqual(saved_event.event_type, 'Saving')
        
        # Verifica impatti generati
        impacts = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts), 24)  # Evento ripetitivo = 24 mesi
        
        # Verifica primo impatto ha pro-rata (mese parziale)
        first_impact = impacts[0]
        self.assertEqual(first_impact.year, 2024)
        self.assertEqual(first_impact.month, 6)
        # Pro-rata: 15 giorni su 30 = 0.5 coefficiente
        # Verifica che il primo mese ha valore minore degli altri (pro-rata applicato)
        second_impact = impacts[1]
        self.assertLess(first_impact.valore_teorico, second_impact.valore_teorico)
        
        # Verifica conservazione valore totale
        total_teorico = sum(i.valore_teorico for i in impacts)
        total_effettivo = sum(i.valore_effettivo for i in impacts)
        self.assertAlmostEqual(total_teorico, 2000.0, places=1)  # 10000 - 8000
        self.assertAlmostEqual(total_effettivo, 1600.0, places=1)  # 2000 * 0.8
    
    def test_update_event_with_impacts(self):
        """Test aggiornamento evento con rigenerazione impatti (DELETE-REGENERATE-SAVE)"""
        event = self._create_test_event_repetitive()
        
        # Salva evento iniziale
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Verifica impatti iniziali
        impacts_before = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_before), 24)
        
        # Modifica evento (cambia percentuale realizzo)
        event.id = event_id
        event.percent_realizzo = 50.0  # Da 80% a 50%
        
        # Aggiorna evento
        update_event_with_impacts(self.db_manager, event)
        
        # Verifica impatti rigenerati
        impacts_after = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_after), 24)
        
        # Verifica nuovo valore effettivo riflette il cambio
        total_effettivo = sum(i.valore_effettivo for i in impacts_after)
        self.assertAlmostEqual(total_effettivo, 1000.0, places=1)  # 2000 * 0.5
        
        # Verifica nessun duplicato (DELETE funziona)
        # Conta impatti per periodo: ogni (year, month) deve apparire esattamente 1 volta
        period_counts = {}
        for impact in impacts_after:
            key = (impact.year, impact.month)
            period_counts[key] = period_counts.get(key, 0) + 1
        
        for count in period_counts.values():
            self.assertEqual(count, 1, "Trovati impatti duplicati dopo update!")
    
    def test_delete_event_and_impacts(self):
        """Test eliminazione evento e impatti correlati"""
        event = self._create_test_event_repetitive()
        
        # Salva evento
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Verifica esistenza
        impacts_before = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_before), 24)
        
        # Elimina evento + impatti
        delete_event_and_impacts(self.db_manager, event_id)
        
        # Verifica evento eliminato
        deleted_event = self.db_manager.get_vsm_event_by_id(event_id)
        self.assertIsNone(deleted_event)
        
        # Verifica impatti eliminati
        impacts_after = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_after), 0)
    
    def test_update_twice_no_duplication(self):
        """Test CRITICO: verifica che aggiornamenti multipli non creino duplicati"""
        event = self._create_test_event_repetitive()
        
        # Salva evento
        event_id = save_event_with_impacts(self.db_manager, event)
        event.id = event_id
        
        # Update 1
        event.percent_realizzo = 90.0
        update_event_with_impacts(self.db_manager, event)
        impacts_1 = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_1), 24)
        
        # Update 2
        event.percent_realizzo = 70.0
        update_event_with_impacts(self.db_manager, event)
        impacts_2 = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_2), 24)
        
        # Update 3
        event.importo_negoziato = 7500.0
        update_event_with_impacts(self.db_manager, event)
        impacts_3 = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts_3), 24)
        
        # Verifica SQL diretta: nessun duplicato per periodo
        self.db_manager.cursor.execute("""
            SELECT anno, mese, COUNT(*) as cnt
            FROM vsm_impacts
            WHERE event_id = ?
            GROUP BY anno, mese
            HAVING cnt > 1
        """, (event_id,))
        duplicates = self.db_manager.cursor.fetchall()
        self.assertEqual(len(duplicates), 0, f"Trovati duplicati: {duplicates}")
    
    def test_one_shot_event_persistence(self):
        """Test evento one-shot crea esattamente 1 impatto"""
        event = self._create_test_event_one_shot()
        
        # Salva evento one-shot
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Verifica 1 solo impatto
        impacts = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts), 1)
        
        # Verifica impatto nel mese evento
        impact = impacts[0]
        self.assertEqual(impact.year, 2024)
        self.assertEqual(impact.month, 6)
        
        # Verifica valore pieno (no pro-rata per one-shot)
        self.assertAlmostEqual(impact.valore_teorico, 5000.0, places=1)  # 15000 - 10000
        self.assertAlmostEqual(impact.valore_effettivo, 5000.0, places=1)  # 100% realizzo
    
    def test_repetitive_event_persistence(self):
        """Test evento ripetitivo crea 24 impatti con pro-rata primo mese"""
        event = self._create_test_event_repetitive()
        
        # Salva evento ripetitivo
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Verifica 24 impatti
        impacts = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(impacts), 24)
        
        # Verifica periodo: da giugno 2024 a maggio 2026
        self.assertEqual(impacts[0].year, 2024)
        self.assertEqual(impacts[0].month, 6)
        self.assertEqual(impacts[-1].year, 2026)
        self.assertEqual(impacts[-1].month, 5)
        
        # Verifica pro-rata primo mese
        first_impact = impacts[0]
        # 15 giugno = 15 giorni rimanenti su 30 = 0.5 coefficiente
        # Ma distribuzione normalizzata considera tutti i mesi
        # Valore primo mese < valore mesi successivi
        second_impact = impacts[1]
        self.assertLess(first_impact.valore_teorico, second_impact.valore_teorico)
    
    def test_save_event_without_id(self):
        """Test validazione: save richiede event_id None o 0"""
        event = self._create_test_event_repetitive()
        event.id = 999  # ID preesistente non valido per save
        
        with self.assertRaises(VSMError) as ctx:
            save_event_with_impacts(self.db_manager, event)
        
        self.assertIn("evento nuovo", str(ctx.exception))
    
    def test_update_event_requires_id(self):
        """Test validazione: update richiede event_id valido"""
        event = self._create_test_event_repetitive()
        event.id = None  # Nessun ID
        
        with self.assertRaises(VSMError) as ctx:
            update_event_with_impacts(self.db_manager, event)
        
        self.assertIn("evento esistente", str(ctx.exception))
    
    def test_delete_requires_valid_id(self):
        """Test validazione: delete richiede event_id valido"""
        with self.assertRaises(VSMError):
            delete_event_and_impacts(self.db_manager, None)
        
        with self.assertRaises(VSMError):
            delete_event_and_impacts(self.db_manager, 0)
        
        with self.assertRaises(VSMError):
            delete_event_and_impacts(self.db_manager, -1)
    
    def test_get_event_with_impacts(self):
        """Test recupero evento completo con impatti"""
        event = self._create_test_event_repetitive()
        
        # Salva evento
        event_id = save_event_with_impacts(self.db_manager, event)
        
        # Recupera evento + impatti
        retrieved_event, impacts = get_event_with_impacts(self.db_manager, event_id)
        
        # Verifica evento
        self.assertIsNotNone(retrieved_event)
        self.assertEqual(retrieved_event.username, 'test_user')
        
        # Verifica impatti
        self.assertEqual(len(impacts), 24)
    
    def test_get_impacts_by_period(self):
        """Test recupero impatti per periodo specifico"""
        event1 = self._create_test_event_repetitive()
        event2 = self._create_test_event_one_shot()
        
        # Salva due eventi (stesso utente, stesso mese)
        event_id1 = save_event_with_impacts(self.db_manager, event1)
        event_id2 = save_event_with_impacts(self.db_manager, event2)
        
        # Recupera impatti per giugno 2024
        impacts_june = self.db_manager.get_vsm_impacts_by_period(2024, 6, 'test_user')
        
        # Verifica 2 impatti (1 da evento ripetitivo + 1 da one-shot)
        self.assertEqual(len(impacts_june), 2)
        
        # Verifica event_id distinti
        event_ids = {i.event_id for i in impacts_june}
        self.assertEqual(len(event_ids), 2)
        self.assertIn(event_id1, event_ids)
        self.assertIn(event_id2, event_ids)
    
    def test_event_id_not_null_constraint(self):
        """Test: event_id deve essere NOT NULL nel database"""
        # Tenta inserimento diretto di impact senza event_id
        from models.vsm_impact import VSMImpact
        
        impact = VSMImpact(
            event_id=None,
            username='test_user',
            year=2024,
            month=6,
            value_type='Teorico',
            valore_teorico=100.0,
            valore_effettivo=80.0
        )
        
        # Tentativo inserimento tramite batch
        with self.assertRaises(Exception):  # SQLite solleverà IntegrityError
            self.db_manager.insert_vsm_impacts_batch([impact])
    
    # ========================================================================
    # TEST ATOMICITÀ TRANSAZIONI
    # ========================================================================
    
    def test_save_rollback_on_impact_insert_failure(self):
        """
        Test ATOMICITÀ save: se inserimento impacts fallisce,
        anche l'evento deve essere annullato (rollback completo).
        """
        # Crea evento valido
        event = self._create_test_event_repetitive()
        
        # Memorizza conteggio eventi prima del test
        self.db_manager.cursor.execute("SELECT COUNT(*) FROM vsm_events")
        events_before = self.db_manager.cursor.fetchone()[0]
        
        # Forza errore manomettendo l'evento per causare failure in generate_impacts
        # Usa un event_date None per forzare errore durante generazione impatti
        event.event_date = None  # Questo causerà errore in generate_impacts_for_event
        
        # Tentativo di salvataggio deve fallire
        with self.assertRaises(Exception):
            save_event_with_impacts(self.db_manager, event)
        
        # VERIFICA ROLLBACK: nessun evento salvato nel database
        self.db_manager.cursor.execute("SELECT COUNT(*) FROM vsm_events")
        events_after = self.db_manager.cursor.fetchone()[0]
        
        self.assertEqual(events_before, events_after, 
                        "Evento NON dovrebbe esistere dopo rollback")
        
        # VERIFICA: nessun impact orfano nel database
        self.db_manager.cursor.execute("SELECT COUNT(*) FROM vsm_impacts")
        impacts_count = self.db_manager.cursor.fetchone()[0]
        self.assertEqual(impacts_count, 0, 
                        "Nessun impact dovrebbe esistere dopo rollback")
    
    def test_update_rollback_preserves_original_state(self):
        """
        Test ATOMICITÀ update: se aggiornamento fallisce,
        evento e impacts devono rimanere nello stato originale.
        """
        # Setup: crea e salva evento originale
        original_event = self._create_test_event_repetitive()
        event_id = save_event_with_impacts(self.db_manager, original_event)
        
        # Memorizza stato originale
        original_saved = self.db_manager.get_vsm_event_by_id(event_id)
        original_impacts = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        original_impacts_count = len(original_impacts)
        original_percent = original_saved.percent_realizzo
        
        # Prepara update con dati che causeranno errore
        updated_event = self._create_test_event_repetitive()
        updated_event.id = event_id
        updated_event.percent_realizzo = 50.0  # Modifica valida
        updated_event.event_date = None  # Questo causerà errore durante regenerate_impacts
        
        # Tentativo di update deve fallire
        with self.assertRaises(Exception):
            update_event_with_impacts(self.db_manager, updated_event)
        
        # VERIFICA ROLLBACK: evento rimane con valori originali
        current_event = self.db_manager.get_vsm_event_by_id(event_id)
        self.assertEqual(current_event.percent_realizzo, original_percent,
                        "Evento dovrebbe mantenere percent_realizzo originale dopo rollback")
        
        # VERIFICA: impacts rimangono nella versione originale
        current_impacts = self.db_manager.get_vsm_impacts_by_event_id(event_id)
        self.assertEqual(len(current_impacts), original_impacts_count,
                        "Conteggio impacts dovrebbe rimanere invariato dopo rollback")
        
        # VERIFICA: valori degli impacts non modificati
        for orig_imp, curr_imp in zip(original_impacts, current_impacts):
            self.assertAlmostEqual(orig_imp.valore_effettivo, curr_imp.valore_effettivo,
                                  places=2, 
                                  msg="Valori impacts NON dovrebbero cambiare dopo rollback")
    
    def test_save_atomicity_no_orphan_events(self):
        """
        Test ATOMICITÀ: verifica che non rimangano eventi orfani
        senza impacts in caso di fallimento parziale.
        """
        # Crea evento con dati che causeranno errore in fase di impacts
        event = self._create_test_event_repetitive()
        event.spending_annuo = 0  # Causerà divisione per zero o errore in calcolo impatti
        event.importo_negoziato = None  # Valore non valido
        
        # Tentativo salvataggio
        with self.assertRaises(Exception):
            save_event_with_impacts(self.db_manager, event)
        
        # VERIFICA: query per eventi senza impacts correlati
        self.db_manager.cursor.execute("""
            SELECT e.event_id
            FROM vsm_events e
            LEFT JOIN vsm_impacts i ON e.event_id = i.event_id
            WHERE i.impact_id IS NULL
        """)
        orphan_events = self.db_manager.cursor.fetchall()
        
        self.assertEqual(len(orphan_events), 0,
                        f"Trovati {len(orphan_events)} eventi orfani senza impacts")


if __name__ == '__main__':
    unittest.main(verbosity=2)
