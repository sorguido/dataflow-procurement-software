"""
Unit tests for VSM Event Model calculations.

Test cases for calculate_theoretical_value() and calculate_effective_value()
with different annual quantities and drivers.
"""

import unittest
import sys
from pathlib import Path
from datetime import datetime

sys.path.insert(0, str(Path(__file__).parent.parent))

from models.vsm_event import VSMEvent


class TestVSMEventCalculations(unittest.TestCase):
    """Test per i metodi di calcolo del modello VSMEvent."""
    
    def test_saving_price_qty_1(self):
        """Saving + Prezzo + quantità 1 (caso base)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=1.0,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_saving_price_qty_large(self):
        """Saving + Prezzo + quantità > 1 (produzione volumi)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=1.5,  # prezzo unitario
            importo_negoziato=1.3,  # prezzo unitario
            quantita_annua=20000.0,  # pezzi/anno
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertAlmostEqual(theoretical, 4000.0, places=2)  # 20000 * 0.2
        self.assertAlmostEqual(effective, 4000.0, places=2)
    
    def test_saving_price_qty_decimal(self):
        """Saving + Prezzo + quantità decimale."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=100.0,
            importo_negoziato=80.0,
            quantita_annua=150.5,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        expected = 150.5 * 20.0  # 3010.0
        self.assertAlmostEqual(theoretical, expected, places=2)
        self.assertAlmostEqual(effective, expected, places=2)
    
    def test_saving_price_qty_none_defaults_to_1(self):
        """Saving + Prezzo + quantità None deve usare default 1.0."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=None,  # NULL nel DB
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        # Deve comportarsi come qty=1
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_saving_price_qty_zero_defaults_to_1(self):
        """Saving + Prezzo + quantità 0 deve usare default 1.0."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Prezzo",
            importo_bdg=20000.0,
            importo_negoziato=18000.0,
            quantita_annua=0.0,
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        # Deve comportarsi come qty=1
        self.assertAlmostEqual(theoretical, 2000.0, places=2)
        self.assertAlmostEqual(effective, 2000.0, places=2)
    
    def test_cost_avoidance_price_qty_large(self):
        """Cost Avoidance + Prezzo + quantità > 1."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Cost Avoidance",
            driver="Prezzo",
            importo_richiesto_iniziale=2.0,
            importo_negoziato=1.8,
            quantita_annua=15000.0,
            percent_realizzo=80.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        expected_theoretical = 15000.0 * 0.2  # 3000.0
        expected_effective = 3000.0 * 0.8  # 2400.0
        
        self.assertAlmostEqual(theoretical, expected_theoretical, places=2)
        self.assertAlmostEqual(effective, expected_effective, places=2)
    
    def test_pagamenti_driver_qty_ignored(self):
        """Driver Pagamenti: quantità NON deve influenzare calcolo."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Saving",
            driver="Pagamenti",
            spending_annuo=100000.0,
            giorni_pagamento_attuali=30,
            giorni_pagamento_negoziati=60,
            quantita_annua=999.0,  # Deve essere ignorato
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        
        # Formula: spending * (delta / 30) * coeff
        # 100000 * (30 / 30) * 0.005 = 500
        self.assertAlmostEqual(theoretical, 500.0, places=2)
    
    def test_derisking_qty_ignored(self):
        """Derisking: quantità NON deve influenzare calcolo (sempre 0)."""
        event = VSMEvent(
            event_date=datetime(2026, 3, 15),
            username="test_user",
            event_type="Derisking",
            driver="Prezzo",
            quantita_annua=999.0,  # Deve essere ignorato
            percent_realizzo=100.0
        )
        
        theoretical = event.calculate_theoretical_value()
        effective = event.calculate_effective_value()
        
        self.assertEqual(theoretical, 0.0)
        self.assertEqual(effective, 0.0)


if __name__ == '__main__':
    unittest.main()
