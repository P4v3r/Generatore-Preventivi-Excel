"""Pipeline matching: Manual → Fuzzy → AI."""

from typing import Optional, List, Dict
from rapidfuzz import process
from models import MatchMethod, MatchResult, Prestazione


class Matcher:
    """Gestisce il matching descrizione → codice prestazione."""
    
    def __init__(
        self,
        mapping: Dict[str, str],
        catalogo: List[Prestazione],
        fuzzy_threshold: int = 85,
        use_ai: bool = False,
    ):
        self.mapping = mapping
        self.catalogo = catalogo
        self.fuzzy_threshold = fuzzy_threshold
        self.use_ai = use_ai
    
    def match(self, descrizione: str) -> MatchResult:
        """
        Trova il codice elanco per la descrizione.
        
        Pipeline:
        1. Dizionario manuale (priorità alta)
        2. Fuzzy matching sul catalogo
        3. AI fallback (se abilitato)
        """
        # Step 1: Manual mapping
        result = self._try_manual(descrizione)
        if result:
            return result
        
        # Step 2: Fuzzy matching
        result = self._try_fuzzy(descrizione)
        if result:
            return result
        
        # Step 3: AI fallback
        if self.use_ai:
            result = self._try_ai(descrizione)
            if result:
                return result
        
        # Nessun match
        return MatchResult(
            codice_elanco=None,
            metodo=MatchMethod.NO_MATCH,
            confidence=0.0,
        )
    
    def _try_manual(self, descrizione: str) -> Optional[MatchResult]:
        """Prova match manuale dal dizionario."""
        descrizione_lower = descrizione.lower()
        
        # Cerca match esatto o parziale
        for key, value in self.mapping.items():
            key_lower = key.lower()
            if key_lower in descrizione_lower or descrizione_lower in key_lower:
                return MatchResult(
                    codice_elanco=value,
                    metodo=MatchMethod.MANUAL,
                    confidence=100.0,
                )
        
        return None
    
    def _try_fuzzy(self, descrizione: str) -> Optional[MatchResult]:
        """Prova fuzzy matching sul catalogo."""
        if not self.catalogo:
            return None
        
        # Crea lista (codice, definizione) per fuzzy
        choices = {
            p.codice_elanco: p.definizione
            for p in self.catalogo
        }
        
        # Trova match migliore
        result = process.extractOne(
            descrizione,
            choices,
            score_cutoff=self.fuzzy_threshold,
        )
        
        if result:
            # result = (definizione, score, codice_elanco) quando choices è un dict
            # result[0] = definizione (value del dict), result[2] = codice_elanco (chiave del dict)
            codice_elanco = result[2]
            score = result[1]
            return MatchResult(
                codice_elanco=codice_elanco,
                metodo=MatchMethod.FUZZY,
                confidence=score,
            )
        
        return None
    
    def _try_ai(self, descrizione: str) -> Optional[MatchResult]:
        """Fallback AI con Cloudflare Workers."""
        # TODO: implementare chiamata AI
        # Chiamata a @cf/meta/llama-3.1-8b-instruct
        return None