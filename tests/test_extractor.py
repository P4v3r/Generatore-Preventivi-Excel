"""Tests per extractor.py."""

from extractor import extract_quantity


def test_extract_integer_simple():
    # Pattern "X unità"
    assert extract_quantity("Installazione 5 pali") == 5
    assert extract_quantity("Posa 10 sostegni") == 10
    assert extract_quantity("Sostituzione 3 lanterne") == 3


def test_extract_with_unit():
    assert extract_quantity("Montaggio 3 apparecchi") == 3
    assert extract_quantity("Installazione 15 bracci") == 15


def test_extract_no_match():
    # Default a 1
    assert extract_quantity("Sostituzione generica") == 1
    assert extract_quantity("Riarmato QE pronto intervento") == 1


def test_extract_various_patterns():
    # Vari pattern
    assert extract_quantity("n. 5 interruttori") == 5
    assert extract_quantity("pezzi: 3") == 3
    assert extract_quantity("quantità: 12") == 12
    assert extract_quantity("5pz") == 5


def test_extract_case_insensitive():
    assert extract_quantity("INSTALLAZIONE 10 PALI") == 10
    assert extract_quantity("posa 7 bracci") == 7