"""
Unit tests for VSM Engine module.

Test cases covered:
1. Saving ripetitivo con durata fino a 24 mesi
2. Cost Avoidance non ripetitivo distribuito fino a dicembre anno evento
3. Primo mese pro-rata
4. Derisking → lista vuota
5. Propagazione corretta di username e event_id
6. Ordinamento cronologico corretto
7. Errori su dati mancanti o tipo evento non supportato
8. Conservazione matematica del totale distribuito
"""

import unittest
import sys
from pathlib import Path
from datetime import datetime

# Aggiungi il path del progetto
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.vsm_event import VSMEvent
from models.vsm_impact import VSMImpact
from services.vsm_engine import (
    generate_impacts_for_event,
    generate_impacts_for_events,
    VSMError,
    _calculate_first_month_coefficient,
    _calculate_distribution_months,
    _distribute_value
)


class TestVSMEngineHelpers(unittest.TestCase):
    """Test per le funzioni helper private."""
    
    def test_calculate_first_month_coefficient_middle_month(self):
        """Test coefficiente pro-rata per evento a metà mese (giorno 16)."""
        event_date = datetime(2026, 3, 16)
        coeff = _calculate_first_month_coefficient(event_date)
        
        # Giorni residui: 30 - 16 + 1 = 15
        # Coefficiente: 15 / 30 = 0.5
        self.assertAlmostEqual(coeff, 0.5, places=4)
    
    def test_calculate_first_month_coefficient_start_month(self):
        """Test coefficiente pro-rata per evento a inizio mese (giorno 1)."""
        event_date = datetime(2026, 3, 1)
        coeff = _calculate_first_month_coefficient(event_date)
        
        # Giorni residui: 30 - 1 + 1 = 30
        # Coefficiente: 30 / 30 = 1.0
        self.assertAlmostEqual(coeff, 1.0, places=4)
    
    def test_calculate_first_month_coefficient_end_month(self):
        """Test coefficiente pro-rata per evento a fine mese (giorno 30)."""
        event_date = datetime(2026, 3, 30)
        coeff = _calculate_first_month_coefficient(event_date)
        
        # Giorni residui: 30 - 30 + 1 = 1
        # Coefficiente: 1 / 30 = 0.0333...
        self.assertAlmostEqual(coeff, 1/30, places=4)
    
    def test_calculate_distribution_months_non_repetitive(self):
        """Test calcolo mesi per evento non ripetitivo (one-shot)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            opex_ripetitivo=False
        )
        
        months = _calculate_distribution_months(event)
        
        # One-shot: un solo mese (quello dell'evento)
        self.assertEqual(len(months), 1)
        self.assertEqual(months[0], (2026, 3))
    
    def test_calculate_distribution_months_repetitive_24_months(self):
        """Test calcolo mesi per evento ripetitivo (24 mesi)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            opex_ripetitivo=True
        )
        
        months = _calculate_distribution_months(event)
        
        # 24 mesi a partire da marzo 2026
        self.assertEqual(len(months), 24)
        self.assertEqual(months[0], (2026, 3))
        self.assertEqual(months[-1], (2028, 2))  # Marzo 2026 + 24 mesi = Febbraio 2028
    
    def test_distribute_value_conservation(self):
        """Test conservazione matematica del valore totale nella distribuzione."""
        total_value = 10000.0
        months = [(2026, 3), (2026, 4), (2026, 5)]
        first_month_coeff = 0.5
        
        monthly_values = _distribute_value(total_value, months, first_month_coeff)
        
        # Verifica conservazione totale (con tolleranza per arrotondamenti)
        self.assertAlmostEqual(sum(monthly_values), total_value, places=2)
    
    def test_distribute_value_zero(self):
        """Test distribuzione con valore zero."""
        total_value = 0.0
        months = [(2026, 3), (2026, 4)]
        first_month_coeff = 0.5
        
        monthly_values = _distribute_value(total_value, months, first_month_coeff)
        
        self.assertEqual(monthly_values, [0.0, 0.0])


class TestGenerateImpactsForEvent(unittest.TestCase):
    """Test per la funzione principale generate_impacts_for_event."""
    
    def test_saving_repetitive_full_24_months(self):
        """Test 1: Saving ripetitivo con durata fino a 24 mesi."""
        event = VSMEvent(
            id=1,
            event_date=datetime(2026, 1, 1),
            username="buyer1",
            event_type="Saving",
            opex_ripetitivo=True,
            importo_bdg=12000.0,
            importo_negoziato=10000.0,
            percent_realizzo=100.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Verifica numero impatti (24 mesi)
        self.assertEqual(len(impacts), 24)
        
        # Verifica propagazione username e event_id
        for impact in impacts:
            self.assertEqual(impact.username, "buyer1")
            self.assertEqual(impact.event_id, 1)
            self.assertEqual(impact.value_type, "Saving")
        
        # Verifica ordinamento cronologico
        for i in range(len(impacts) - 1):
            current = (impacts[i].year, impacts[i].month)
            next_period = (impacts[i+1].year, impacts[i+1].month)
            self.assertLess(current, next_period)
        
        # Verifica conservazione valore totale
        total_teorico = sum(imp.valore_teorico for imp in impacts)
        total_effettivo = sum(imp.valore_effettivo for imp in impacts)
        
        self.assertAlmostEqual(total_teorico, 2000.0, places=2)  # 12000 - 10000
        self.assertAlmostEqual(total_effettivo, 2000.0, places=2)  # 100% realizzo
    
    def test_cost_avoidance_non_repetitive_year_end(self):
        """Test 2: Cost Avoidance non ripetitivo (one-shot, impatto singolo)."""
        event = VSMEvent(
            id=2,
            event_date=datetime(2026, 3, 1),
            username="buyer2",
            event_type="Cost Avoidance",
            opex_ripetitivo=False,
            driver="Prezzo",
            importo_richiesto_iniziale=15000.0,
            importo_negoziato=12000.0,
            percent_realizzo=80.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # One-shot: un solo impatto nel mese dell'evento
        self.assertEqual(len(impacts), 1)
        
        # Verifica che sia marzo 2026
        self.assertEqual(impacts[0].year, 2026)
        self.assertEqual(impacts[0].month, 3)
        
        # Verifica conservazione valore
        total_teorico = sum(imp.valore_teorico for imp in impacts)
        total_effettivo = sum(imp.valore_effettivo for imp in impacts)
        
        expected_teorico = 15000.0 - 12000.0  # 3000
        expected_effettivo = expected_teorico * 0.8  # 2400
        
        self.assertAlmostEqual(total_teorico, expected_teorico, places=2)
        self.assertAlmostEqual(total_effettivo, expected_effettivo, places=2)
    
    def test_first_month_prorata(self):
        """Test 3: Primo mese pro-rata corretto (solo per eventi ripetitivi)."""
        event = VSMEvent(
            id=3,
            event_date=datetime(2026, 3, 16),  # Metà mese
            username="buyer3",
            event_type="Saving",
            opex_ripetitivo=True,  # Ripetitivo per testare il pro-rata
            importo_bdg=10000.0,
            importo_negoziato=9000.0,
            percent_realizzo=100.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Ripetitivo: 24 mesi
        self.assertEqual(len(impacts), 24)
        
        # Primo impatto (marzo) dovrebbe essere circa metà del valore medio
        # Coefficienti: 0.5 (marzo) + 23 * 1.0 = 23.5 totale
        # Valore unitario = 1000 / 23.5 = 42.55
        # Marzo = 42.55 * 0.5 = 21.28
        
        # Il primo mese deve essere significativamente più basso degli altri
        self.assertLess(impacts[0].valore_teorico, impacts[1].valore_teorico)
        
        # Verifica conservazione totale
        total_teorico = sum(imp.valore_teorico for imp in impacts)
        self.assertAlmostEqual(total_teorico, 1000.0, places=2)
    
    def test_one_shot_no_prorata(self):
        """Test 3b: Eventi one-shot NON hanno pro-rata (valore intero nel mese evento)."""
        event = VSMEvent(
            id=30,
            event_date=datetime(2026, 3, 16),  # Metà mese
            username="buyer_oneshot",
            event_type="Saving",
            opex_ripetitivo=False,  # One-shot
            importo_bdg=10000.0,
            importo_negoziato=9000.0,
            percent_realizzo=100.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # One-shot: esattamente 1 impatto
        self.assertEqual(len(impacts), 1)
        
        # Deve essere nel mese dell'evento (marzo)
        self.assertEqual(impacts[0].year, 2026)
        self.assertEqual(impacts[0].month, 3)
        
        # Valore intero (NO pro-rata): deve essere l'intero valore dell'evento
        expected_value = 1000.0  # 10000 - 9000
        self.assertAlmostEqual(impacts[0].valore_teorico, expected_value, places=2)
        self.assertAlmostEqual(impacts[0].valore_effettivo, expected_value, places=2)
    
    def test_derisking_empty_impacts(self):
        """Test 4: Derisking restituisce lista vuota."""
        event = VSMEvent(
            id=4,
            event_date=datetime(2026, 3, 15),
            username="buyer4",
            event_type="Derisking",
            opex_ripetitivo=True
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Derisking non genera impatti economici
        self.assertEqual(impacts, [])
    
    def test_username_event_id_propagation(self):
        """Test 5: Propagazione corretta di username e event_id."""
        event = VSMEvent(
            id=999,
            event_date=datetime(2026, 6, 1),
            username="test_buyer_xyz",
            event_type="Saving",
            opex_ripetitivo=False,
            importo_bdg=5000.0,
            importo_negoziato=4500.0,
            percent_realizzo=90.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Tutti gli impatti devono avere username e event_id corretti
        for impact in impacts:
            self.assertEqual(impact.username, "test_buyer_xyz")
            self.assertEqual(impact.event_id, 999)
    
    def test_chronological_ordering(self):
        """Test 6: Ordinamento cronologico corretto."""
        event = VSMEvent(
            id=6,
            event_date=datetime(2025, 11, 1),  # Novembre, vicino a fine anno
            username="buyer6",
            event_type="Saving",
            opex_ripetitivo=True,
            importo_bdg=12000.0,
            importo_negoziato=11000.0,
            percent_realizzo=100.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Verifica ordinamento stretto
        for i in range(len(impacts) - 1):
            current_year = impacts[i].year
            current_month = impacts[i].month
            next_year = impacts[i+1].year
            next_month = impacts[i+1].month
            
            # Anno deve essere ≤ anno successivo
            self.assertLessEqual(current_year, next_year)
            
            # Se stesso anno, mese deve essere < mese successivo
            if current_year == next_year:
                self.assertLess(current_month, next_month)
    
    def test_missing_event_date_error(self):
        """Test 7a: Errore su event_date mancante."""
        event = VSMEvent(
            id=7,
            event_date=None,  # Mancante
            username="buyer7",
            event_type="Saving"
        )
        
        with self.assertRaises(VSMError) as context:
            generate_impacts_for_event(event)
        
        self.assertIn("event_date", str(context.exception))
    
    def test_missing_username_error(self):
        """Test 7b: Errore su username mancante."""
        event = VSMEvent(
            id=8,
            event_date=datetime(2026, 3, 15),
            username="",  # Mancante
            event_type="Saving"
        )
        
        with self.assertRaises(VSMError) as context:
            generate_impacts_for_event(event)
        
        self.assertIn("username", str(context.exception))
    
    def test_invalid_event_type_error(self):
        """Test 7c: Errore su tipo evento non supportato."""
        event = VSMEvent(
            id=9,
            event_date=datetime(2026, 3, 15),
            username="buyer9",
            event_type="InvalidType"  # Non valido
        )
        
        with self.assertRaises(VSMError) as context:
            generate_impacts_for_event(event)
        
        self.assertIn("non valido", str(context.exception).lower())
    
    def test_value_conservation_theoretical_and_effective(self):
        """Test 8: Conservazione matematica valori teorici ed effettivi."""
        event = VSMEvent(
            id=10,
            event_date=datetime(2026, 5, 10),
            username="buyer10",
            event_type="Cost Avoidance",
            opex_ripetitivo=True,
            driver="Prezzo",
            importo_richiesto_iniziale=20000.0,
            importo_negoziato=18000.0,
            percent_realizzo=75.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Calcola totali dagli impatti
        total_teorico_impacts = sum(imp.valore_teorico for imp in impacts)
        total_effettivo_impacts = sum(imp.valore_effettivo for imp in impacts)
        
        # Calcola totali dall'evento
        expected_teorico = event.calculate_theoretical_value()
        expected_effettivo = event.calculate_effective_value()
        
        # Verifica conservazione (tolleranza 0.01 per arrotondamenti)
        self.assertAlmostEqual(total_teorico_impacts, expected_teorico, places=2)
        self.assertAlmostEqual(total_effettivo_impacts, expected_effettivo, places=2)
        
        # Verifica anche relazione teorico/effettivo
        self.assertAlmostEqual(
            total_effettivo_impacts,
            total_teorico_impacts * 0.75,
            places=2
        )


class TestGenerateImpactsForEventsBatch(unittest.TestCase):
    """Test per la funzione batch generate_impacts_for_events."""
    
    def test_batch_processing_all_success(self):
        """Test batch processing con tutti eventi validi."""
        events = [
            VSMEvent(
                id=1,
                event_date=datetime(2026, 1, 1),
                username="buyer1",
                event_type="Saving",
                opex_ripetitivo=False,
                importo_bdg=1000.0,
                importo_negoziato=900.0
            ),
            VSMEvent(
                id=2,
                event_date=datetime(2026, 2, 1),
                username="buyer2",
                event_type="Cost Avoidance",
                opex_ripetitivo=False,
                importo_richiesto_iniziale=2000.0,
                importo_negoziato=1800.0
            ),
        ]
        
        impacts_map = generate_impacts_for_events(events)
        
        # Entrambi gli eventi dovrebbero essere presenti
        self.assertIn(1, impacts_map)
        self.assertIn(2, impacts_map)
        self.assertEqual(len(impacts_map), 2)
    
    def test_batch_processing_with_failures(self):
        """Test batch processing con alcuni eventi non validi (continua comunque)."""
        events = [
            VSMEvent(
                id=1,
                event_date=datetime(2026, 1, 1),
                username="buyer1",
                event_type="Saving",
                opex_ripetitivo=False,
                importo_bdg=1000.0,
                importo_negoziato=900.0
            ),
            VSMEvent(
                id=2,
                event_date=None,  # Invalido
                username="buyer2",
                event_type="Saving"
            ),
            VSMEvent(
                id=3,
                event_date=datetime(2026, 3, 1),
                username="buyer3",
                event_type="Cost Avoidance",
                opex_ripetitivo=False,
                importo_richiesto_iniziale=3000.0,
                importo_negoziato=2700.0
            ),
        ]
        
        impacts_map = generate_impacts_for_events(events)
        
        # Eventi 1 e 3 dovrebbero essere presenti, evento 2 no
        self.assertIn(1, impacts_map)
        self.assertNotIn(2, impacts_map)
        self.assertIn(3, impacts_map)
        self.assertEqual(len(impacts_map), 2)
    
    def test_batch_processing_with_derisking(self):
        """Test batch processing con evento Derisking (lista vuota)."""
        events = [
            VSMEvent(
                id=1,
                event_date=datetime(2026, 1, 1),
                username="buyer1",
                event_type="Derisking",
                opex_ripetitivo=False
            ),
        ]
        
        impacts_map = generate_impacts_for_events(events)
        
        # Evento presente ma con lista vuota
        self.assertIn(1, impacts_map)
        self.assertEqual(impacts_map[1], [])


class TestEdgeCases(unittest.TestCase):
    """Test per casi limite."""
    
    def test_event_id_none_accepted(self):
        """Test che event_id=None sia accettato e mantenuto (evento non ancora persistito)."""
        event = VSMEvent(
            id=None,  # Non ancora persistito
            event_date=datetime(2026, 3, 15),
            username="buyer_new",
            event_type="Saving",
            opex_ripetitivo=False,
            importo_bdg=1000.0,
            importo_negoziato=900.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Deve generare impatti normalmente
        self.assertGreater(len(impacts), 0)
        
        # event_id negli impatti deve rimanere None (non convertito)
        for impact in impacts:
            self.assertIsNone(impact.event_id)
    
    def test_event_last_month_of_year(self):
        """Test evento a dicembre (non ripetitivo genera solo 1 impatto)."""
        event = VSMEvent(
            id=100,
            event_date=datetime(2026, 12, 1),
            username="buyer_dec",
            event_type="Saving",
            opex_ripetitivo=False,
            importo_bdg=1000.0,
            importo_negoziato=900.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Solo dicembre
        self.assertEqual(len(impacts), 1)
        self.assertEqual(impacts[0].month, 12)
        self.assertEqual(impacts[0].year, 2026)
    
    def test_percent_realizzo_zero(self):
        """Test con percent_realizzo = 0 (valore effettivo = 0)."""
        event = VSMEvent(
            id=101,
            event_date=datetime(2026, 6, 1),
            username="buyer_zero",
            event_type="Saving",
            opex_ripetitivo=False,
            driver="Prezzo",
            importo_bdg=1000.0,
            importo_negoziato=900.0,
            percent_realizzo=0.0
        )
        
        impacts = generate_impacts_for_event(event)
        
        # Valore teorico presente, effettivo zero
        total_teorico = sum(imp.valore_teorico for imp in impacts)
        total_effettivo = sum(imp.valore_effettivo for imp in impacts)
        
        self.assertAlmostEqual(total_teorico, 100.0, places=2)
        self.assertAlmostEqual(total_effettivo, 0.0, places=2)


if __name__ == '__main__':
    unittest.main(verbosity=2)
