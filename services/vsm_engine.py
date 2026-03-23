"""
VSM Engine Module

Motore di calcolo per la generazione automatica degli impatti economici mensili (VSMImpact)
a partire da eventi VSM (VSMEvent).

Il modulo implementa la logica di business per:
- Distribuzione del valore economico nel tempo
- Calcolo pro-rata per il primo mese
- Gestione del riverbero per eventi OPEX ripetitivi (max 24 mesi)
- Propagazione dei dati di multiutenza
"""

import logging
from datetime import datetime
from typing import List, Tuple, Dict, Optional

# Import dei modelli VSM
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent))

from models.vsm_event import VSMEvent
from models.vsm_impact import VSMImpact


# Inizializzazione logger
logger = logging.getLogger('DataFlow.VSMEngine')


# Eccezione custom per errori VSM Engine
class VSMError(Exception):
    """
    Eccezione sollevata per errori nella logica di business del VSM Engine.
    
    Utilizzata per segnalare:
    - Dati mancanti o incompleti
    - Tipi di evento non supportati
    - Errori di validazione
    """
    pass


# Tipi di evento supportati (valori esatti richiesti)
VALID_EVENT_TYPES = {"Saving", "Cost Avoidance", "Derisking"}


def _validate_event(event: VSMEvent) -> None:
    """
    Valida che l'evento contenga i dati minimi necessari per il calcolo.
    
    Args:
        event: Evento VSM da validare
        
    Raises:
        VSMError: Se l'evento non contiene dati validi
    """
    if not event.event_date:
        raise VSMError("Campo 'event_date' mancante o non valido")
    
    if not event.username:
        raise VSMError("Campo 'username' mancante (richiesto per multiutenza)")
    
    if event.event_type not in VALID_EVENT_TYPES:
        raise VSMError(
            f"Tipo evento non valido: '{event.event_type}'. "
            f"Valori ammessi: {', '.join(sorted(VALID_EVENT_TYPES))}"
        )


def _calculate_first_month_coefficient(event_date: datetime) -> float:
    """
    Calcola il coefficiente pro-rata per il primo mese.
    
    Convenzione commerciale a 30 giorni:
    - coefficiente = giorni_residui_incluso_giorno_evento / 30
    
    Esempio:
    - Evento al giorno 16 del mese
    - Giorni residui: 30 - 16 + 1 = 15
    - Coefficiente: 15 / 30 = 0.5
    
    Args:
        event_date: Data dell'evento
        
    Returns:
        float: Coefficiente pro-rata (0.0 - 1.0)
    """
    day = event_date.day
    
    # Convenzione commerciale: 30 giorni per mese
    # Giorni residui = 30 - (giorno - 1) = 30 - giorno + 1
    giorni_residui = 30 - day + 1
    
    # Coefficiente pro-rata
    coefficient = giorni_residui / 30.0
    
    logger.debug(
        f"Calcolo pro-rata primo mese: giorno {day}, "
        f"giorni residui {giorni_residui}, coefficiente {coefficient:.4f}"
    )
    
    return coefficient


def _calculate_distribution_months(event: VSMEvent) -> List[Tuple[int, int]]:
    """
    Calcola i mesi su cui distribuire l'impatto economico.
    
    Logica:
    - Se evento ripetitivo (opex_ripetitivo=True): massimo 24 mesi dal mese dell'evento
    - Se evento non ripetitivo (one-shot): un solo mese (il mese dell'evento)
    
    Args:
        event: Evento VSM
        
    Returns:
        List[Tuple[int, int]]: Lista di tuple (year, month) in ordine cronologico
    """
    if not event.event_date:
        return []
    
    start_year = event.event_date.year
    start_month = event.event_date.month
    
    months = []
    
    if event.opex_ripetitivo:
        # Evento ripetitivo: massimo 24 mesi
        current_year = start_year
        current_month = start_month
        
        for _ in range(24):
            months.append((current_year, current_month))
            
            # Passa al mese successivo
            current_month += 1
            if current_month > 12:
                current_month = 1
                current_year += 1
    else:
        # Evento non ripetitivo (one-shot): un solo impatto nel mese dell'evento
        # Tipicamente eventi Capex che impattano una sola volta
        months.append((start_year, start_month))
    
    logger.debug(
        f"Distribuzione mesi: evento {'ripetitivo' if event.opex_ripetitivo else 'non ripetitivo'}, "
        f"{len(months)} mesi calcolati"
    )
    
    return months


def _distribute_value(
    total_value: float,
    months: List[Tuple[int, int]],
    first_month_coefficient: float
) -> List[float]:
    """
    Distribuisce il valore totale sui mesi usando coefficienti normalizzati.
    
    Logica matematica corretta:
    1. Calcolare coefficienti: primo mese = first_month_coefficient, altri = 1.0
    2. Calcolare somma coefficienti: total_coeff = sum(coefficients)
    3. Calcolare valore unitario: unit_value = total_value / total_coeff
    4. Assegnare quote: quota_mese = unit_value * coefficiente_mese
    5. Aggiustare l'ultimo mese per garantire somma esatta = total_value
    
    Args:
        total_value: Valore totale da distribuire
        months: Lista di mesi su cui distribuire
        first_month_coefficient: Coefficiente pro-rata per il primo mese
        
    Returns:
        List[float]: Lista di valori mensili (stessa lunghezza di months)
    """
    if not months:
        return []
    
    if total_value == 0:
        return [0.0] * len(months)
    
    # 1. Costruisci lista coefficienti
    coefficients = [first_month_coefficient] + [1.0] * (len(months) - 1)
    
    # 2. Calcola somma coefficienti
    total_coeff = sum(coefficients)
    
    # 3. Calcola valore unitario
    unit_value = total_value / total_coeff
    
    # 4. Assegna quote mensili
    monthly_values = [unit_value * coeff for coeff in coefficients]
    
    # 5. Aggiusta l'ultimo mese per compensare arrotondamenti
    calculated_sum = sum(monthly_values)
    if len(monthly_values) > 0:
        adjustment = total_value - calculated_sum
        monthly_values[-1] += adjustment
    
    logger.debug(
        f"Distribuzione valore: totale {total_value:.2f}, "
        f"coeff totale {total_coeff:.4f}, "
        f"valore unitario {unit_value:.2f}, "
        f"aggiustamento ultimo mese {adjustment:.6f}"
    )
    
    return monthly_values


def generate_impacts_for_event(event: VSMEvent) -> List[VSMImpact]:
    """
    Genera la lista di impatti economici mensili per un evento VSM.
    
    Questa è la funzione principale del modulo. Gestisce:
    - Validazione dati evento
    - Calcolo distribuzione temporale
    - Generazione impatti con valori teorici ed effettivi
    - Propagazione dati multiutenza
    - Ordinamento cronologico
    
    Comportamento per tipo evento:
    - "Saving": genera impatti economici distribuiti
    - "Cost Avoidance": genera impatti economici distribuiti
    - "Derisking": restituisce lista vuota (solo statistico)
    
    Comportamento per opex_ripetitivo:
    - True: distribuzione su più mesi (max 24), primo mese con pro-rata
    - False: un solo impatto nel mese evento (one-shot, tipicamente Capex)
    
    Args:
        event: Evento VSM da processare
        
    Returns:
        List[VSMImpact]: Lista di impatti mensili ordinati cronologicamente
        
    Raises:
        VSMError: Se l'evento non è valido o contiene dati insufficienti
        
    Examples:
        >>> # Evento ripetitivo (OPEX) - distribuzione multi-mese
        >>> event_opex = VSMEvent(
        ...     event_date=datetime(2026, 3, 16),
        ...     username="buyer1",
        ...     event_type="Saving",
        ...     opex_ripetitivo=True,
        ...     importo_bdg=10000,
        ...     importo_negoziato=9000,
        ...     percent_realizzo=100.0
        ... )
        >>> impacts = generate_impacts_for_event(event_opex)
        >>> len(impacts)  # 24 mesi (ripetitivo)
        24
        
        >>> # Evento one-shot (CAPEX) - impatto singolo
        >>> event_capex = VSMEvent(
        ...     event_date=datetime(2026, 3, 16),
        ...     username="buyer1",
        ...     event_type="Saving",
        ...     opex_ripetitivo=False,
        ...     importo_bdg=10000,
        ...     importo_negoziato=9000,
        ...     percent_realizzo=100.0
        ... )
        >>> impacts = generate_impacts_for_event(event_capex)
        >>> len(impacts)  # 1 solo impatto (one-shot)
        1
    """
    # Validazione evento
    _validate_event(event)
    
    logger.debug(
        f"Generazione impatti per evento: tipo={event.event_type}, "
        f"id={event.id}, user={event.username}, date={event.event_date}"
    )
    
    # Gestione evento Derisking (nessun impatto economico)
    if event.event_type == "Derisking":
        logger.debug("Evento Derisking: nessun impatto economico generato")
        return []
    
    # Calcola valori totali dall'evento
    valore_teorico_totale = event.calculate_theoretical_value()
    valore_effettivo_totale = event.calculate_effective_value()
    
    # Calcola mesi di distribuzione
    months = _calculate_distribution_months(event)
    
    if not months:
        logger.warning("Nessun mese di distribuzione calcolato")
        return []
    
    # Calcola coefficiente pro-rata primo mese (solo per eventi ripetitivi)
    # Per eventi one-shot, il valore intero va assegnato all'unico mese
    if event.opex_ripetitivo:
        first_month_coeff = _calculate_first_month_coefficient(event.event_date)
    else:
        # One-shot: coefficiente 1.0 (valore intero)
        first_month_coeff = 1.0
    
    # Distribuisci valori teorici ed effettivi
    valori_teorici_mensili = _distribute_value(
        valore_teorico_totale,
        months,
        first_month_coeff
    )
    
    valori_effettivi_mensili = _distribute_value(
        valore_effettivo_totale,
        months,
        first_month_coeff
    )
    
    # Genera lista impatti
    impacts = []
    
    for i, (year, month) in enumerate(months):
        impact = VSMImpact(
            id=None,  # Sarà assegnato dalla persistenza
            event_id=event.id,  # Mantenere None se evento non persistito
            username=event.username,
            year=year,
            month=month,
            value_type=event.event_type,
            valore_teorico=valori_teorici_mensili[i],
            valore_effettivo=valori_effettivi_mensili[i]
        )
        impacts.append(impact)
    
    # Gli impatti sono già in ordine cronologico per costruzione
    logger.info(
        f"Generati {len(impacts)} impatti per evento {event.id or 'non persistito'} "
        f"(tipo={event.event_type}, valore_teorico_totale={valore_teorico_totale:.2f})"
    )
    
    return impacts


def generate_impacts_for_events(
    events: List[VSMEvent]
) -> Dict[Optional[int], List[VSMImpact]]:
    """
    Genera impatti per una lista di eventi (batch processing).
    
    Gestione robusta:
    - Se un evento fallisce, logga l'errore e continua con gli altri
    - Eventi falliti non sono inclusi nel risultato
    
    Args:
        events: Lista di eventi VSM da processare
        
    Returns:
        Dict[Optional[int], List[VSMImpact]]: Mappa event_id -> lista impatti
        
    Examples:
        >>> events = [event1, event2, event3]
        >>> impacts_map = generate_impacts_for_events(events)
        >>> impacts_map[event1.id]
        [VSMImpact(...), VSMImpact(...), ...]
    """
    logger.info(f"Elaborazione batch di {len(events)} eventi")
    
    impacts_map = {}
    successes = 0
    failures = 0
    
    for event in events:
        try:
            impacts = generate_impacts_for_event(event)
            event_key = event.id if event.id is not None else None
            impacts_map[event_key] = impacts
            successes += 1
            
        except VSMError as e:
            failures += 1
            logger.error(
                f"Errore generazione impatti per evento {event.id or 'non persistito'}: {e}",
                exc_info=True
            )
            # Continua con gli altri eventi
            continue
        
        except Exception as e:
            failures += 1
            logger.error(
                f"Errore imprevisto per evento {event.id or 'non persistito'}: {e}",
                exc_info=True
            )
            # Continua con gli altri eventi
            continue
    
    logger.info(
        f"Elaborazione batch completata: {successes} successi, {failures} fallimenti"
    )
    
    return impacts_map
