"""
VSM Impact Model

Rappresenta l'impatto economico mensile generato da un evento VSM.
Gli impatti sono generati automaticamente dal motore di calcolo e rappresentano
la quota mensile del valore generato dall'evento.
"""

from dataclasses import dataclass
from typing import Optional


@dataclass
class VSMImpact:
    """
    Classe che rappresenta un impatto economico mensile di un evento VSM.
    
    Gli impatti sono generati automaticamente per distribuire il valore
    dell'evento nel tempo (riverbero fino a 24 mesi per OPEX ripetitivo).
    
    Attributi:
        id: Identificativo univoco dell'impatto
        event_id: ID dell'evento VSM di riferimento
        username: Username del creatore dell'evento (propagato per multiutenza)
        year: Anno di riferimento dell'impatto
        month: Mese di riferimento dell'impatto (1-12)
        value_type: Tipo di valore (Saving / Cost Avoidance)
        valore_teorico: Valore teorico mensile
        valore_effettivo: Valore effettivo mensile (con % realizzo applicata)
    """
    
    # Campi obbligatori
    id: Optional[int] = None
    event_id: int = 0  # Riferimento all'evento VSM
    username: str = ""  # Propagato dall'evento per multiutenza
    
    # Periodo di riferimento
    year: int = 0
    month: int = 0  # 1-12
    
    # Tipo e valori
    value_type: str = ""  # Saving / Cost Avoidance
    valore_teorico: float = 0.0
    valore_effettivo: float = 0.0
    
    def __post_init__(self):
        """
        Validazione post-inizializzazione.
        Verifica che il mese sia compreso tra 1 e 12.
        """
        if self.month < 1 or self.month > 12:
            raise ValueError(f"Mese deve essere compreso tra 1 e 12, ricevuto: {self.month}")
    
    def to_dict(self) -> dict:
        """
        Converte l'impatto in un dizionario per la persistenza.
        
        Returns:
            dict: Rappresentazione dell'impatto come dizionario
        """
        return {
            'id': self.id,
            'event_id': self.event_id,
            'username': self.username,
            'year': self.year,
            'month': self.month,
            'value_type': self.value_type,
            'valore_teorico': self.valore_teorico,
            'valore_effettivo': self.valore_effettivo,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'VSMImpact':
        """
        Crea un'istanza di VSMImpact da un dizionario.
        
        Args:
            data: Dizionario contenente i dati dell'impatto
            
        Returns:
            VSMImpact: Nuova istanza dell'impatto
        """
        return cls(**data)
    
    @property
    def period_key(self) -> str:
        """
        Restituisce una chiave univoca per il periodo (YYYY-MM).
        
        Returns:
            str: Chiave nel formato "YYYY-MM"
        """
        return f"{self.year}-{self.month:02d}"
    
    def get_realizzo_percentage(self) -> float:
        """
        Calcola la percentuale di realizzo effettiva (valore_effettivo / valore_teorico).
        
        Returns:
            float: Percentuale di realizzo (0-100)
        """
        if self.valore_teorico == 0:
            return 0.0
        return (self.valore_effettivo / self.valore_teorico) * 100.0
