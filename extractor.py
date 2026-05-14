"""Estrazione quantità dalla descrizione lavoro."""

import re


# Pattern regex per estrarre quantità
QUANTITY_PATTERNS = [
    r'(\d+)\s*unita',       # "5 unità", "3 UNITA"
    r'(\d+)\s*pz',          # "5 pz", "3 PZ"
    r'n\.\s*(\d+)',         # "n. 5", "N. 5"
    r'quantita\s*:?\s*(\d+)', # "quantità 5", "quantità: 5" (senza accento)
    r'quantità\s*:?\s*(\d+)', # "quantità 5", "quantità: 5" (con accento)
    r'(\d+)\s*pezzi',       # "5 pezzi"
    r'pezzi\s*:\s*(\d+)',   # "pezzi: 3"
    r'(\d+)\s*pali',        # "5 pali"
    r'(\d+)\s*apparecchi',  # "5 apparecchi"
    r'(\d+)\s*sostegni',    # "5 sostegni"
    r'(\d+)\s*bracci',      # "5 bracci"
    r'(\d+)\s*lanterne',    # "5 lanterne"
    r'(\d+)\s*interruttori', # "5 interruttori"
]


def extract_quantity(descrizione: str) -> int:
    """
    Estrae la quantità dalla descrizione del lavoro.
    
    Args:
        descrizione: Testo libero della descrizione
        
    Returns:
        Quantità come intero (default 1 se non trovata)
    """
    descrizione_lower = descrizione.lower()
    
    for pattern in QUANTITY_PATTERNS:
        match = re.search(pattern, descrizione_lower)
        if match:
            return int(match.group(1))
    
    # Default: 1
    return 1