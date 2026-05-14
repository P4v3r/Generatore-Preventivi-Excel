"""Tests per models.py."""

def test_lavoro_dataclass():
    from models import Lavoro
    
    lavoro = Lavoro(
        codice="LAV001",
        data_richiesta="2024-05-10",
        data_esecuzione="2024-05-15",
        comune="Milano",
        via="Via Roma 1",
        descrizione="Installazione 5 pali",
        quantita=5,  # Quantità passata esplicitamente
    )
    
    assert lavoro.codice == "LAV001"
    assert lavoro.comune == "Milano"
    assert lavoro.quantita == 5


def test_prestazione_dataclass():
    from models import Prestazione
    
    prest = Prestazione(
        codice_sap="ITLEIL0001",
        codice_elanco="A.EC.1.1.04",
        definizione="Esecuzione canalizzaz.",
        prezzo_unitario=19.92,
        prezzo_scontato=13.94,
        um="m",
    )
    
    assert prest.codice_elanco == "A.EC.1.1.04"
    assert prest.prezzo_scontato == 13.94


def test_match_result():
    from models import MatchResult, MatchMethod
    
    result = MatchResult(
        codice_elanco="A.EC.2.1.10",
        metodo=MatchMethod.FUZZY,
        confidence=92,
        quantita=5,
    )
    
    assert result.metodo == MatchMethod.FUZZY
    assert result.confidence == 92
    assert result.quantita == 5


def test_match_result_no_match():
    from models import MatchResult, MatchMethod
    
    result = MatchResult(
        codice_elanco=None,
        metodo=MatchMethod.NO_MATCH,
        confidence=0.0,
    )
    
    assert result.codice_elanco is None
    assert result.metodo == MatchMethod.NO_MATCH