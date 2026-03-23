"""
VSM Event Model

Rappresenta un evento VSM (Saving, Cost Avoidance, Derisking).
Gli eventi sono le azioni negoziali che generano valore economico.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


@dataclass
class VSMEvent:
    """
    Classe che rappresenta un evento VSM.
    
    Un evento VSM traccia un'azione negoziale che genera valore economico
    per l'ufficio acquisti, indipendentemente dal workflow RFQ.
    
    Attributi:
        id: Identificativo univoco dell'evento
        event_date: Data dell'evento
        username: Username del creatore (obbligatorio per multiutenza)
        buyer: Nome dell'acquirente responsabile
        event_type: Tipo di evento (Saving / Cost Avoidance / Derisking)
        action: Tipo di azione (Negoziazione / Derisking)
        description: Descrizione dell'evento
        reference: Riferimento esterno (es. RFQ, PO, fornitore)
        importo_bdg: Importo a budget
        importo_negoziato: Importo negoziato finale
        importo_richiesto_iniziale: Importo richiesto inizialmente (solo per Cost Avoidance)
        quantita_annua: Quantità annua prevista
        percent_realizzo: Percentuale di realizzo (valore effettivo vs teorico)
        driver: Tipo di driver del saving (Prezzo / Pagamenti)
        giorni_pagamento_attuali: Giorni di pagamento attuali
        giorni_pagamento_negoziati: Giorni di pagamento negoziati
        spending_annuo: Spending annuo previsto
        opex_ripetitivo: Flag per OPEX ripetitivo (riverbero fino a 24 mesi)
        note: Note aggiuntive
        created_at: Timestamp di creazione
        updated_at: Timestamp ultimo aggiornamento
    """
    
    # Campi obbligatori
    id: Optional[int] = None
    event_date: Optional[datetime] = None
    username: str = ""  # Obbligatorio per multiutenza
    buyer: str = ""
    event_type: str = ""  # Saving / Cost Avoidance / Derisking
    action: str = ""  # Negoziazione / Derisking
    
    # Descrizione e riferimenti
    description: str = ""
    reference: str = ""
    
    # Importi
    importo_bdg: float = 0.0
    importo_negoziato: float = 0.0
    importo_richiesto_iniziale: Optional[float] = None  # Solo per Cost Avoidance
    
    # Quantità e realizzo
    quantita_annua: float = 0.0
    percent_realizzo: float = 100.0  # Default 100%
    
    # Driver e pagamenti
    driver: str = ""  # Prezzo / Pagamenti
    giorni_pagamento_attuali: Optional[int] = None
    giorni_pagamento_negoziati: Optional[int] = None
    spending_annuo: float = 0.0
    
    # Flags
    opex_ripetitivo: bool = False
    
    # Note e metadata
    note: str = ""
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)
    
    def __post_init__(self):
        """
        Validazione post-inizializzazione.
        Converte stringhe datetime se necessario.
        """
        # Converte event_date se è una stringa
        if isinstance(self.event_date, str):
            try:
                self.event_date = datetime.fromisoformat(self.event_date)
            except (ValueError, AttributeError):
                pass
        
        # Converte created_at se è una stringa
        if isinstance(self.created_at, str):
            try:
                self.created_at = datetime.fromisoformat(self.created_at)
            except (ValueError, AttributeError):
                self.created_at = datetime.now()
        
        # Converte updated_at se è una stringa
        if isinstance(self.updated_at, str):
            try:
                self.updated_at = datetime.fromisoformat(self.updated_at)
            except (ValueError, AttributeError):
                self.updated_at = datetime.now()
    
    def to_dict(self) -> dict:
        """
        Converte l'evento in un dizionario per la persistenza.
        
        Returns:
            dict: Rappresentazione dell'evento come dizionario
        """
        return {
            'id': self.id,
            'event_date': self.event_date.isoformat() if self.event_date else None,
            'username': self.username,
            'buyer': self.buyer,
            'event_type': self.event_type,
            'action': self.action,
            'description': self.description,
            'reference': self.reference,
            'importo_bdg': self.importo_bdg,
            'importo_negoziato': self.importo_negoziato,
            'importo_richiesto_iniziale': self.importo_richiesto_iniziale,
            'quantita_annua': self.quantita_annua,
            'percent_realizzo': self.percent_realizzo,
            'driver': self.driver,
            'giorni_pagamento_attuali': self.giorni_pagamento_attuali,
            'giorni_pagamento_negoziati': self.giorni_pagamento_negoziati,
            'spending_annuo': self.spending_annuo,
            'opex_ripetitivo': self.opex_ripetitivo,
            'note': self.note,
            'created_at': self.created_at.isoformat() if self.created_at else None,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'VSMEvent':
        """
        Crea un'istanza di VSMEvent da un dizionario.
        
        Args:
            data: Dizionario contenente i dati dell'evento
            
        Returns:
            VSMEvent: Nuova istanza dell'evento
        """
        return cls(**data)
    
    def calculate_theoretical_value(self) -> float:
        """
        Calcola il valore teorico dell'evento.
        
        Returns:
            float: Valore teorico calcolato
        """
        if self.event_type == "Cost Avoidance" and self.importo_richiesto_iniziale:
            # Cost Avoidance: differenza tra richiesto iniziale e negoziato
            return self.importo_richiesto_iniziale - self.importo_negoziato
        elif self.event_type == "Saving":
            # Saving: differenza tra budget e negoziato
            return self.importo_bdg - self.importo_negoziato
        else:
            # Derisking: nessun valore economico diretto
            return 0.0
    
    def calculate_effective_value(self) -> float:
        """
        Calcola il valore effettivo applicando la percentuale di realizzo.
        
        Returns:
            float: Valore effettivo calcolato
        """
        theoretical = self.calculate_theoretical_value()
        return theoretical * (self.percent_realizzo / 100.0)
