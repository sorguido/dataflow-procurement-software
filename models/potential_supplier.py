"""
Potential Supplier Model

Rappresenta un fornitore potenziale nell'anagrafica del tab Derisking.
Entità separata da VSMEvent: nessuna dipendenza dal modulo VSM.
"""

from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional


# Valori ammessi per supplier_status (usati in UI e query KPI)
SUPPLIER_STATUS_ACTIVE   = "Attivo"
SUPPLIER_STATUS_PROSPECT = "Prospect"
SUPPLIER_STATUS_INACTIVE = "Non attivo"

SUPPLIER_STATUS_CHOICES = [
    SUPPLIER_STATUS_ACTIVE,
    SUPPLIER_STATUS_PROSPECT,
    SUPPLIER_STATUS_INACTIVE,
]


@dataclass
class PotentialSupplier:
    """
    Anagrafica fornitore potenziale.

    Traccia i fornitori potenziali valutati / introdotti come parte
    delle attività di derisking della supply chain.

    Attributi:
        id:                Identificativo univoco (None per record non ancora persistito)
        supplier_name:     Ragione sociale / nome fornitore (obbligatorio)
        macrocategory:     Macrocategoria merceologica (es. "Acciaio", "Plastica")
        merchandise_class: Classe merceologica di dettaglio
        supplier_status:   Stato del fornitore (Attivo / Prospect / Non attivo)
        contact_name:      Nome referente commerciale
        email:             Email di contatto
        phone:             Telefono di contatto
        website:           Sito web aziendale
        notes:             Note libere
        username:          Username del buyer che ha inserito il record
        created_at:        Timestamp di creazione record
        updated_at:        Timestamp ultimo aggiornamento
    """

    # Identificativo
    id: Optional[int] = None

    # Dati anagrafici
    supplier_name: str = ""
    macrocategory: str = ""
    merchandise_class: str = ""
    supplier_status: str = SUPPLIER_STATUS_PROSPECT

    # Contatti
    contact_name: str = ""
    email: str = ""
    phone: str = ""
    website: str = ""

    # Note e metadata
    notes: str = ""
    username: str = ""
    created_at: datetime = field(default_factory=datetime.now)
    updated_at: datetime = field(default_factory=datetime.now)

    def __post_init__(self):
        """Normalizza i tipi datetime se ricevuti come stringa dal DB."""
        if isinstance(self.created_at, str):
            try:
                self.created_at = datetime.fromisoformat(self.created_at)
            except (ValueError, AttributeError):
                self.created_at = datetime.now()

        if isinstance(self.updated_at, str):
            try:
                self.updated_at = datetime.fromisoformat(self.updated_at)
            except (ValueError, AttributeError):
                self.updated_at = datetime.now()

    def to_dict(self) -> dict:
        """Converte il record in dizionario per la persistenza."""
        return {
            'id':                self.id,
            'supplier_name':     self.supplier_name,
            'macrocategory':     self.macrocategory,
            'merchandise_class': self.merchandise_class,
            'supplier_status':   self.supplier_status,
            'contact_name':      self.contact_name,
            'email':             self.email,
            'phone':             self.phone,
            'website':           self.website,
            'notes':             self.notes,
            'username':          self.username,
            'created_at':        self.created_at.isoformat() if self.created_at else None,
            'updated_at':        self.updated_at.isoformat() if self.updated_at else None,
        }

    @classmethod
    def from_row(cls, row) -> 'PotentialSupplier':
        """
        Crea un'istanza da una riga del database (sqlite3.Row o tuple).

        Mappa le colonne della tabella potential_suppliers sui campi del dataclass.
        Robusto a row dict-like (Row) e a tuple posizionale.
        """
        if hasattr(row, 'keys'):
            # sqlite3.Row con row_factory abilitato
            data = dict(row)
        else:
            # Fallback tuple posizionale (ordine colonne come in CREATE TABLE)
            data = {
                'supplier_id':       row[0],
                'supplier_name':     row[1],
                'macrocategory':     row[2],
                'merchandise_class': row[3],
                'supplier_status':   row[4],
                'contact_name':      row[5],
                'email':             row[6],
                'phone':             row[7],
                'website':           row[8],
                'notes':             row[9],
                'username':          row[10],
                'created_at':        row[11],
                'updated_at':        row[12],
            }

        return cls(
            id=data.get('supplier_id'),
            supplier_name=data.get('supplier_name') or '',
            macrocategory=data.get('macrocategory') or '',
            merchandise_class=data.get('merchandise_class') or '',
            supplier_status=data.get('supplier_status') or SUPPLIER_STATUS_PROSPECT,
            contact_name=data.get('contact_name') or '',
            email=data.get('email') or '',
            phone=data.get('phone') or '',
            website=data.get('website') or '',
            notes=data.get('notes') or '',
            username=data.get('username') or '',
            created_at=data.get('created_at') or datetime.now(),
            updated_at=data.get('updated_at') or datetime.now(),
        )
