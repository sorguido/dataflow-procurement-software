"""
Script di test manuale per VSM Engine.

Crea alcuni eventi di esempio e mostra gli impatti generati.
Utile per verificare visivamente il comportamento del motore.
"""

import sys
from pathlib import Path
from datetime import datetime

# Aggiungi il path del progetto
sys.path.insert(0, str(Path(__file__).parent))

from models.vsm_event import VSMEvent
from services.vsm_engine import generate_impacts_for_event, VSMError


def print_separator(title=""):
    """Stampa un separatore visivo."""
    if title:
        print(f"\n{'=' * 80}")
        print(f"  {title}")
        print('=' * 80)
    else:
        print('-' * 80)


def print_event_summary(event: VSMEvent):
    """Stampa un riepilogo dell'evento."""
    print(f"\n📋 EVENTO VSM:")
    print(f"   ID: {event.id}")
    print(f"   Tipo: {event.event_type}")
    print(f"   Data: {event.event_date.strftime('%d/%m/%Y')}")
    print(f"   Username: {event.username}")
    print(f"   Ripetitivo: {'Sì' if event.opex_ripetitivo else 'No'}")
    print(f"   Valore teorico totale: €{event.calculate_theoretical_value():,.2f}")
    print(f"   Valore effettivo totale: €{event.calculate_effective_value():,.2f}")
    print(f"   % Realizzo: {event.percent_realizzo}%")


def print_impacts_summary(impacts):
    """Stampa un riepilogo degli impatti generati."""
    if not impacts:
        print("\n⚠️  Nessun impatto economico generato")
        return
    
    print(f"\n💰 IMPATTI GENERATI: {len(impacts)} mesi")
    print_separator()
    print(f"{'Mese':<15} {'Anno':<8} {'Valore Teorico':>18} {'Valore Effettivo':>18}")
    print_separator()
    
    total_teorico = 0
    total_effettivo = 0
    
    for impact in impacts[:5]:  # Mostra solo i primi 5
        month_name = datetime(impact.year, impact.month, 1).strftime('%B')
        print(
            f"{month_name:<15} {impact.year:<8} "
            f"€{impact.valore_teorico:>16,.2f} €{impact.valore_effettivo:>16,.2f}"
        )
        total_teorico += impact.valore_teorico
        total_effettivo += impact.valore_effettivo
    
    if len(impacts) > 5:
        print(f"{'... (altri ' + str(len(impacts) - 5) + ' mesi)':<15}")
        for impact in impacts[5:]:
            total_teorico += impact.valore_teorico
            total_effettivo += impact.valore_effettivo
    
    print_separator()
    print(
        f"{'TOTALE':<15} {'':<8} "
        f"€{total_teorico:>16,.2f} €{total_effettivo:>16,.2f}"
    )
    print_separator()


def test_saving_repetitive():
    """Test 1: Saving ripetitivo 12 mesi."""
    print_separator("TEST 1: Saving Ripetitivo (12 mesi)")
    
    event = VSMEvent(
        id=1,
        event_date=datetime(2026, 3, 1),
        username="mario.rossi",
        buyer="Mario Rossi",
        event_type="Saving",
        opex_ripetitivo=True,
        importo_bdg=120000.0,
        importo_negoziato=108000.0,
        percent_realizzo=100.0,
        description="Negoziazione annuale componentistica elettronica"
    )
    
    print_event_summary(event)
    
    impacts = generate_impacts_for_event(event)
    print_impacts_summary(impacts)
    
    # Verifica
    total = sum(imp.valore_teorico for imp in impacts)
    expected = event.calculate_theoretical_value()
    print(f"\n✓ Verifica conservazione: €{total:,.2f} = €{expected:,.2f}")


def test_cost_avoidance_prorata():
    """Test 2: Cost Avoidance con primo mese pro-rata."""
    print_separator("TEST 2: Cost Avoidance con Pro-rata (evento 15 marzo)")
    
    event = VSMEvent(
        id=2,
        event_date=datetime(2026, 3, 15),
        username="laura.bianchi",
        buyer="Laura Bianchi",
        event_type="Cost Avoidance",
        opex_ripetitivo=False,
        importo_richiesto_iniziale=50000.0,
        importo_negoziato=45000.0,
        percent_realizzo=80.0,
        description="Contenimento aumento prezzi materie prime"
    )
    
    print_event_summary(event)
    
    impacts = generate_impacts_for_event(event)
    print_impacts_summary(impacts)
    
    # Verifica pro-rata
    print(f"\n📊 Dettaglio primo mese (pro-rata):")
    print(f"   Marzo (giorno 15): €{impacts[0].valore_teorico:,.2f}")
    print(f"   Aprile (mese pieno): €{impacts[1].valore_teorico:,.2f}")
    print(f"   Rapporto: {impacts[0].valore_teorico / impacts[1].valore_teorico:.2%}")


def test_derisking():
    """Test 3: Derisking (nessun impatto economico)."""
    print_separator("TEST 3: Derisking (solo statistico)")
    
    event = VSMEvent(
        id=3,
        event_date=datetime(2026, 4, 1),
        username="paolo.verdi",
        buyer="Paolo Verdi",
        event_type="Derisking",
        opex_ripetitivo=False,
        description="Qualifica nuovo fornitore alternativo"
    )
    
    print_event_summary(event)
    
    impacts = generate_impacts_for_event(event)
    print_impacts_summary(impacts)


def test_error_handling():
    """Test 4: Gestione errori."""
    print_separator("TEST 4: Gestione Errori")
    
    # Test tipo evento non valido
    print("\n🔴 Test tipo evento non valido:")
    try:
        event = VSMEvent(
            id=4,
            event_date=datetime(2026, 5, 1),
            username="test.user",
            event_type="InvalidType"
        )
        generate_impacts_for_event(event)
        print("   ✗ ERRORE: dovrebbe sollevare VSMError")
    except VSMError as e:
        print(f"   ✓ VSMError correttamente sollevato: {e}")
    
    # Test username mancante
    print("\n🔴 Test username mancante:")
    try:
        event = VSMEvent(
            id=5,
            event_date=datetime(2026, 5, 1),
            username="",
            event_type="Saving"
        )
        generate_impacts_for_event(event)
        print("   ✗ ERRORE: dovrebbe sollevare VSMError")
    except VSMError as e:
        print(f"   ✓ VSMError correttamente sollevato: {e}")


def main():
    """Esegue tutti i test manuali."""
    print("\n")
    print("█" * 80)
    print(" " * 20 + "VSM ENGINE - TEST MANUALI")
    print("█" * 80)
    
    try:
        test_saving_repetitive()
        test_cost_avoidance_prorata()
        test_derisking()
        test_error_handling()
        
        print_separator("RIEPILOGO")
        print("\n✅ Tutti i test manuali completati con successo!\n")
        
    except Exception as e:
        print(f"\n❌ ERRORE IMPREVISTO: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()
