"""Tests per matcher.py."""

from matcher import Matcher
from models import MatchMethod, MatchResult, Prestazione


def test_match_manual_exact():
    # Setup con mapping manuale
    mapping = {"sostituzione": "A.ES.5.1.10"}
    
    m = Matcher(mapping=mapping, catalogo=[])
    
    result = m.match("Sostituzione apparecchio")
    
    assert result.codice_elanco == "A.ES.5.1.10"
    assert result.metodo == MatchMethod.MANUAL


def test_match_manual_partial():
    mapping = {"pronto intervento": "A.ES.5.2.63"}
    
    m = Matcher(mapping=mapping, catalogo=[])
    
    result = m.match("Pronto intervento riarmato QE")
    
    assert result.codice_elanco == "A.ES.5.2.63"
    assert result.metodo == MatchMethod.MANUAL


def test_match_no_mapping():
    m = Matcher(mapping={}, catalogo=[])
    
    result = m.match("Descrizione sconosciuta")
    
    assert result.metodo == MatchMethod.NO_MATCH


def test_match_fuzzy():
    # Setup con catalogo
    catalogo = [
        Prestazione(
            codice_sap="SAP1",
            codice_elanco="A.EC.1.1.04",
            definizione="Esecuzione canalizzaz. banchina stradale",
            prezzo_unitario=19.92,
            prezzo_scontato=13.94,
            um="m",
        ),
        Prestazione(
            codice_sap="SAP2",
            codice_elanco="A.EC.2.1.10",
            definizione="Posa di sostegni metallici diametro fino a 11,5 cm",
            prezzo_unitario=170.36,
            prezzo_scontato=119.25,
            um="cad",
        ),
    ]
    
    m = Matcher(mapping={}, catalogo=catalogo, fuzzy_threshold=50)  # threshold più basso per test
    
    result = m.match("Esecuzione canalizzazione banchina stradale")
    
    # Il fuzzy dovrebbe trovare un match
    assert result.codice_elanco == "A.EC.1.1.04"
    assert result.metodo == MatchMethod.FUZZY
    assert result.confidence >= 50


def test_match_fuzzy_low_score():
    catalogo = [
        Prestazione(
            codice_sap="SAP1",
            codice_elanco="A.EC.1.1.04",
            definizione="Esecuzione canalizzaz. banchina stradale",
            prezzo_unitario=19.92,
            prezzo_scontato=13.94,
            um="m",
        ),
    ]
    
    m = Matcher(mapping={}, catalogo=catalogo, fuzzy_threshold=85)
    
    # Descrizione molto diversa dovrebbe dare NO_MATCH
    result = m.match("Installazione ascensore")
    
    assert result.metodo == MatchMethod.NO_MATCH