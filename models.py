"""Data classes per il generatore preventivi."""

from dataclasses import dataclass
from enum import Enum
from typing import Optional


class MatchMethod(Enum):
    """Metodo di matching usato."""
    MANUAL = "manual"
    FUZZY = "fuzzy"
    AI = "ai"
    NO_MATCH = "no_match"


@dataclass
class Lavoro:
    """Un nuovo lavoro da elaborare."""
    codice: str
    data_richiesta: str
    data_esecuzione: str
    comune: str
    via: str
    descrizione: str
    quantita: int = 1


@dataclass
class Prestazione:
    """Una prestazione dal catalogo."""
    codice_sap: str
    codice_elanco: str
    definizione: str
    prezzo_unitario: float
    prezzo_scontato: float
    um: str


@dataclass
class MatchResult:
    """Risultato del matching di una descrizione a una prestazione."""
    codice_elanco: Optional[str]
    metodo: MatchMethod
    confidence: float
    quantita: int = 1