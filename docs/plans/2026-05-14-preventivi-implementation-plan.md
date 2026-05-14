# Generatore Preventivi Excel - Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Generare preventivi Excel automaticamente matchando descrizioni lavori a prestazioni catalogo.

**Architecture:** Script CLI modulare con pipeline matching 3-step (manual → fuzzy → AI), estrazione quantità regex, generazione Excel da template preservando formattazione.

**Tech Stack:** Python 3.14 • pandas • openpyxl • rapidfuzz

---

## Dipendenze Task

```
Task 1 → Task 2 → Task 3 → Task 4 → Task 5 → Task 6 → Task 7
   ↓          ↓          ↓
Task 8     Task 9     Task 10 → Task 11
```

---

## Task 1: Configurazione

**Files:**
- Create: `lavoro/config.py`
- Create: `lavoro/requirements.txt`

**Step 1: Create config.py**

```python
"""Configurazione generatore preventivi."""

from pathlib import Path

# Path
BASE_DIR = Path(__file__).parent
TEMPLATE_DIR = BASE_DIR
OUTPUT_DIR = BASE_DIR / "output"

# Template
TEMPLATE_MILANO = "template-mi.xltx"
TEMPLATE_LIGURIA = "template-lig.xltx"

# Input files
CATALOGO_FILE = BASE_DIR / "catalogo_prestazioni.csv"
LAVORI_FILE = BASE_DIR / "nuovi_lavori.csv"
MAPPING_FILE = BASE_DIR / "mapping_manuali.json"

# Matching
FUZZY_THRESHOLD = 85
USE_AI = False  # Attivare con flag --use-ai

# CSV input columns
CSV_COLS = {
    "codice": "Cod SF - PL",
    "data_richiesta": "Data Richiesta",
    "data_esecuzione": "Data Esecuzione",
    "comune": "Comune",
    "via": "Località",
    "descrizione": "Descrizione",
}

# Celle foglio lavoro
CELLS = {
    "codice": "B2",
    "data_richiesta": "B3",
    "data_esecuzione": "B4",
    "comune": "B5",
    "via": "B6",
    "descrizione": "B7",
}

# Colonna quantità nella tabella prestazioni
COL_QUANTITA = "D"

# Riga inizio prestazioni (PREST_SP marker)
PRESTAZIONI_START_ROW = 233

# Totale
CELL_TOTALE = "I8"

# Foglio Totale
TOTALE_SHEET = "Totale"
TOTALE_HEADER_ROWS = 2  # Prime 2 righe sono intestazione
```

**Step 2: Create requirements.txt**

```
pandas>=2.0
openpyxl>=3.1
rapidfuzz>=3.0
requests>=2.28
```

**Step 3: Commit**

```bash
cd /Users/Privacy/Documents/projects/webapps-ai/lavoro
git add config.py requirements.txt
git commit -m "feat: add config module and requirements"
```

---

## Task 2: Data Models

**Files:**
- Create: `lavoro/models.py`
- Create: `lavoro/tests/test_models.py`

**Step 1: Create test_models.py**

```python
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
    )
    
    assert lavoro.codice == "LAV001"
    assert lavoro.comune == "Milano"

def test_prestazione_dataclass():
    from models import Prestazione
    
    prest = Prestazione(
        codice_sap="ITLEIL0001",
        codice_elanco="A.EC.1.1.04",
        definizione="Esecuzione canalizzaz.",
        prezzo_scontato=13.94,
        um="m",
    )
    
    assert prest.codice_elanco == "A.EC.1.1.04"

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
```

**Step 2: Run test**

Run: `python -m pytest tests/test_models.py -v`
Expected: ERROR (module not found)

**Step 3: Create models.py**

```python
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
```

**Step 4: Run test**

Run: `python -m pytest tests/test_models.py -v`
Expected: PASS

**Step 5: Commit**

```bash
git add models.py tests/test_models.py
git commit -m "feat: add data models (Lavoro, Prestazione, MatchResult)"
```

---

## Task 3: Extractor Quantità

**Files:**
- Create: `lavoro/extractor.py`
- Create: `lavoro/tests/test_extractor.py`

**Step 1: Create test_extractor.py**

```python
"""Tests per extractor.py."""

def test_extract_integer_simple():
    from extractor import extract_quantity
    
    # Pattern "X unità"
    assert extract_quantity("Installazione 5 pali") == 5
    assert extract_quantity("Posa 10 sostegni") == 10
    
def test_extract_with_unit():
    from extractor import extract_quantity
    
    assert extract_quantity("Montaggio 3 apparecchi") == 3
    
def test_extract_no_match():
    from extractor import extract_quantity
    
    # Default a 1
    assert extract_quantity("Sostituzione generica") == 1

def test_extract_various_patterns():
    from extractor import extract_quantity
    
    # Vari pattern
    assert extract_quantity("n. 5 interruttori") == 5
    assert extract_quantity("pezzi: 3") == 3
```

**Step 2: Run test**

Run: `python -m pytest tests/test_extractor.py -v`
Expected: FAIL (extractor not found)

**Step 3: Create extractor.py**

```python
"""Estrazione quantità dalla descrizione lavoro."""

import re
from typing import Optional


# Pattern regex per estrarre quantità
QUANTITY_PATTERNS = [
    r'(\d+)\s*unita',       # "5 unità", "3 UNITA"
    r'(\d+)\s*pz',          # "5 pz", "3 PZ"
    r'n\.\s*(\d+)',         # "n. 5", "N. 5"
    r'quantita\s*(\d+)',    # "quantità 5"
    r'(\d+)\s*pezzi',       # "5 pezzi"
    r'(\d+)\s*pali',        # "5 pali"
    r'(\d+)\s*apparecchi',  # "5 apparecchi"
    r'(\d+)\s*sostegni',    # "5 sostegni"
    r'(\d+)\s*bracci',      # "5 bracci"
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
```

**Step 4: Run test**

Run: `python -m pytest tests/test_extractor.py -v`
Expected: PASS

**Step 5: Commit**

```bash
git add extractor.py tests/test_extractor.py
git commit -m "feat: add quantity extractor from description"
```

---

## Task 4: Matcher (Pipeline 3-Step)

**Files:**
- Create: `lavoro/matcher.py`
- Create: `lavoro/tests/test_matcher.py`

**Step 1: Create test_matcher.py**

```python
"""Tests per matcher.py."""

def test_match_manual_exact():
    from matcher import Matcher
    from models import MatchMethod
    
    # Setup con mapping manuale
    mapping = {"sostituzione": "A.ES.5.1.10"}
    
    m = Matcher(mapping=mapping, catalogo=[])
    
    result = m.match("Sostituzione apparecchio")
    
    assert result.codice_elanco == "A.ES.5.1.10"
    assert result.metodo == MatchMethod.MANUAL

def test_match_no_mapping():
    from matcher import Matcher
    from models import MatchMethod
    
    m = Matcher(mapping={}, catalogo=[])
    
    result = m.match("Descrizione sconosciuta")
    
    assert result.metodo == MatchMethod.NO_MATCH

def test_match_fuzzy():
    from matcher import Matcher
    from models import MatchMethod
    from models import Prestazione
    
    # Setup con catalogo
    catalogo = [
        Prestazione(
            codice_sap="SAP1",
            codice_elanco="A.EC.1.1.04",
            definizione="Esecuzione canalizzaz. banchina",
            prezzo_unitario=19.92,
            prezzo_scontato=13.94,
            um="m",
        ),
    ]
    
    m = Matcher(mapping={}, catalogo=catalogo, fuzzy_threshold=85)
    
    result = m.match("Esecuzione canalizzazione banchina")
    
    assert result.codice_elanco == "A.EC.1.1.04"
    assert result.metodo == MatchMethod.FUZZY
    assert result.confidence >= 85
```

**Step 2: Run test**

Run: `python -m pytest tests/test_matcher.py -v`
Expected: FAIL

**Step 3: Create matcher.py**

```python
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
        choices = [
            (p.codice_elanco, p.definizione)
            for p in self.catalogo
        ]
        
        # Trova match migliore
        result = process.extractOne(
            descrizione,
            {codice: definizione for codice, definizione in choices},
            score_cutoff=self.fuzzy_threshold,
        )
        
        if result:
            codice_elanco, score = result[0], result[1]
            return MatchResult(
                codice_elanco=codice_elanco,
                metodo=MatchMethod.FUZZY,
                confidence=score,
            )
        
        return None
    
    def _try_ai(self, descrizione: str) -> Optional[MatchResult]:
        """Fallback AI con Cloudflare Workers."""
        # TODO: implementare chiamata AI
        return None
```

**Step 4: Run test**

Run: `python -m pytest tests/test_matcher.py -v`
Expected: PASS

**Step 5: Commit**

```bash
git add matcher.py tests/test_matcher.py
git commit -m "feat: add 3-step matching pipeline (manual, fuzzy, AI)"
```

---

## Task 5: Excel Generator

**Files:**
- Create: `lavoro/excel_generator.py`
- Create: `lavoro/tests/test_excel_generator.py`

**Step 1: Create test_excel_generator.py**

```python
"""Tests per excel_generator.py."""

import tempfile
import os

def test_create_output_dir():
    from excel_generator import ensure_output_dir
    
    with tempfile.TemporaryDirectory() as tmpdir:
        output_dir = os.path.join(tmpdir, "output")
        path = ensure_output_dir(output_dir)
        
        assert os.path.exists(path)

def test_get_template_path():
    from excel_generator import get_template_path
    
    # Milano
    path = get_template_path("milano")
    assert "template-mi" in path
    
    # Liguria
    path = get_template_path("liguria")
    assert "template-lig" in path
```

**Step 2: Run test**

Run: `python -m pytest tests/test_excel_generator.py -v`
Expected: FAIL

**Step 3: Create excel_generator.py**

```python
"""Generazione file Excel da template."""

import os
import shutil
from pathlib import Path
from datetime import date
from typing import List
import openpyxl
from openpyxl import load_workbook

from config import (
    TEMPLATE_MILANO,
    TEMPLATE_LIGURIA,
    OUTPUT_DIR,
    CELLS,
    TOTALE_SHEET,
    TOTALE_HEADER_ROWS,
)
from models import Lavoro, MatchResult


def ensure_output_dir() -> Path:
    """Crea directory output se non esiste."""
    output_path = Path(OUTPUT_DIR)
    output_path.mkdir(parents=True, exist_ok=True)
    return output_path


def get_template_path(regione: str) -> Path:
    """Restituisce path template in base alla regione."""
    if regione.lower() in ("milano", "mi"):
        return Path(TEMPLATE_MILANO)
    elif regione.lower() in ("liguria", "lig"):
        return Path(TEMPLATE_LIGURIA)
    else:
        raise ValueError(f"Regione non supportata: {regione}")


def genera_preventivo(
    comune: str,
    lavori: List[Lavoro],
    match_results: dict,
    regione: str,
) -> Path:
    """
    Genera un file Excel preventivo per un comune.
    
    Args:
        comune: Nome comune
        lavori: Lista lavori del comune
        match_results: Dict {codice: MatchResult}
        regione: Template da usare (milano/liguria)
    
    Returns:
        Path del file generato
    """
    # Carica template
    template_path = get_template_path(regione)
    wb = load_workbook(template_path)
    
    # Crea foglio Totale (copia dal template se esiste)
    if TOTALE_SHEET in wb.sheetnames:
        ws_totale = wb[TOTALE_SHEET]
    else:
        ws_totale = wb.create_sheet(TOTALE_SHEET)
    
    next_row = TOTALE_HEADER_ROWS + 1
    
    # Per ogni lavoro, crea un nuovo foglio
    for lavoro in lavori:
        # Clona il foglio template (cerca "LAV_TEMPLATE" o primo foglio dati)
        source_sheet = _get_source_sheet(wb)
        
        if source_sheet:
            new_sheet = wb.copy_worksheet(source_sheet)
            new_sheet.title = lavoro.codice
        else:
            # Crea nuovo foglio vuoto
            new_sheet = wb.create_sheet(lavoro.codice)
        
        # Compila celle foglio lavoro
        _compile_lavoro_sheet(new_sheet, lavoro, match_results.get(lavoro.codice))
        
        # Aggiungi riga a foglio Totale
        _add_totale_row(ws_totale, next_row, lavoro)
        next_row += 1
    
    # Salva file
    output_dir = ensure_output_dir()
    filename = f"Preventivo_{comune}_{date.today().isoformat()}.xlsx"
    output_path = output_dir / filename
    
    wb.save(output_path)
    return output_path


def _get_source_sheet(wb) -> openpyxl.worksheet.worksheet.Worksheet | None:
    """Trova il foglio template da usare come base."""
    # Cerca foglio con "LAV_TEMPLATE"
    for sheet_name in wb.sheetnames:
        if "template" in sheet_name.lower() or sheet_name.startswith("LAV"):
            return wb[sheet_name]
    
    # Altrimenti restituisci il primo foglio non "Totale"
    for sheet_name in wb.sheetnames:
        if sheet_name != TOTALE_SHEET:
            return wb[sheet_name]
    
    return None


def _compile_lavoro_sheet(ws, lavoro: Lavoro, match_result: MatchResult):
    """Compila le celle del foglio lavoro."""
    # Celle base
    ws[CELLS["codice"]] = lavoro.codice
    ws[CELLS["data_richiesta"]] = lavoro.data_richiesta
    ws[CELLS["data_esecuzione"]] = lavoro.data_esecuzione
    ws[CELLS["comune"]] = lavoro.comune
    ws[CELLS["via"]] = lavoro.via
    ws[CELLS["descrizione"]] = lavoro.descrizione
    
    # Se c'è match, trova riga prestazione e inserisci quantità
    if match_result and match_result.codice_elanco:
        _insert_quantita(ws, match_result.codice_elanco, match_result.quantita)


def _insert_quantita(ws, codice_elanco: str, quantita: int):
    """Trova riga con codice elanco e inserisci quantità."""
    # Itera righe per trovare match
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and codice_elanco in str(cell.value):
                # Trovato! Inserisci quantità nella colonna "Quantità"
                # (dovrebbe essere la 3° colonna, ma verificalo dal template)
                quantita_col = _find_quantita_colonna(ws)
                if quantita_col:
                    ws.cell(row=cell.row, column=quantita_col).value = quantita
                return


def _find_quantita_colonna(ws) -> int | None:
    """Trova la colonna Quantità nell'header."""
    # Cerca "Quantità" nella prima riga
    for cell in ws[1]:
        if cell.value and "quantita" in str(cell.value).lower():
            return cell.column
    
    # Prova con la 3° colonna (index 3)
    return 3


def _add_totale_row(ws_totale, row: int, lavoro: Lavoro):
    """Aggiunge riga al foglio Totale."""
    # B1: Codice, C1: Data, D1: Importo (I8 dal foglio lavoro)
    ws_totale[f"B{row}"] = lavoro.codice
    ws_totale[f"C{row}"] = lavoro.data_esecuzione
    ws_totale[f"D{row}"] = f"='{lavoro.codice}'!I8"  # Riferimento a cella I8 del foglio lavoro
```

**Step 4: Run test**

Run: `python -m pytest tests/test_excel_generator.py -v`
Expected: PASS

**Step 5: Commit**

```bash
git add excel_generator.py tests/test_excel_generator.py
git commit -m "feat: add Excel generator with template handling"
```

---

## Task 6: CLI Interface

**Files:**
- Create: `lavoro/cli.py`
- Create: `lavoro/tests/test_cli.py`

**Step 1: Create test_cli.py**

```python
"""Tests per cli.py."""

import argparse

def test_parse_args_defaults():
    from cli import parse_args
    
    args = parse_args(["--input", "test.csv"])
    
    assert args.input == "test.csv"
    assert args.template == "milano"  # Default
    assert args.use_ai == False

def test_parse_args_full():
    from cli import parse_args
    
    args = parse_args([
        "--input", "lavori.csv",
        "--template", "liguria",
        "--use-ai"
    ])
    
    assert args.template == "liguria"
    assert args.use_ai == True

def test_parse_args_help():
    from cli import parse_args
    
    try:
        args = parse_args(["--help"])
    except SystemExit:
        pass
```

**Step 2: Run test**

Run: `python -m pytest tests/test_cli.py -v`
Expected: FAIL

**Step 3: Create cli.py**

```python
"""CLI interface per generatore preventivi."""

import argparse
import sys
from pathlib import Path


def parse_args(args=None):
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Genera preventivi Excel da template regionali",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Esempi:
  %(prog)s --input nuovi_lavori.csv --template milano
  %(prog)s --input lavori.csv --template liguria --use-ai
        """,
    )
    
    parser.add_argument(
        "--input", "-i",
        required=True,
        help="Path al file CSV con nuovi lavori",
    )
    
    parser.add_argument(
        "--template", "-t",
        choices=["milano", "liguria"],
        default="milano",
        help="Template regionale (default: milano)",
    )
    
    parser.add_argument(
        "--use-ai",
        action="store_true",
        help="Attiva fallback AI per match non trovati",
    )
    
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Output verboso",
    )
    
    return parser.parse_args(args)


def validate_args(args) -> bool:
    """Valida gli argomenti."""
    input_path = Path(args.input)
    
    if not input_path.exists():
        print(f"❌ Errore: File non trovato: {input_path}")
        return False
    
    return True


def print_banner():
    """Stampa banner iniziale."""
    print("=" * 50)
    print("  Generatore Preventivi Excel")
    print("=" * 50)
    print()


def print_summary(generati: int, errori: int, log_path: Path):
    """Stampa riepilogo fine esecuzione."""
    print()
    print("=" * 50)
    print(f"  ✅ Preventivi generati: {generati}")
    print(f"  ❌ Errori: {errori}")
    print(f"  📋 Log: {log_path}")
    print("=" * 50)
```

**Step 4: Run test**

Run: `python -m pytest tests/test_cli.py -v`
Expected: PASS

**Step 5: Commit**

```bash
git add cli.py tests/test_cli.py
git commit -m "feat: add CLI interface with argparse"
```

---

## Task 7: Main Orchestrator

**Files:**
- Create: `lavoro/main.py`
- Create: `lavoro/genera_preventivi.py` (entry point)

**Step 1: Create main.py**

```python
"""Orchestrazione principale del generatore."""

import csv
import json
from pathlib import Path
from collections import defaultdict
from datetime import date

import pandas as pd

from config import CATALOGO_FILE, LAVORI_FILE, MAPPING_FILE, CSV_COLS
from models import Lavoro, Prestazione, MatchMethod, MatchResult
from matcher import Matcher
from extractor import extract_quantity
from excel_generator import genera_preventivo


class PreventiviGenerator:
    """Generatore principale."""
    
    def __init__(self, regione: str = "milano", use_ai: bool = False):
        self.regione = regione
        self.use_ai = use_ai
        
        # Carica dati
        self.catalogo = self._load_catalogo()
        self.mapping = self._load_mapping()
        
        # Crea matcher
        self.matcher = Matcher(
            mapping=self.mapping,
            catalogo=self.catalogo,
            use_ai=use_ai,
        )
    
    def run(self, input_file: str) -> tuple[list[Path], list[str]]:
        """
        Esegue la generazione preventivi.
        
        Returns:
            (lista_path_generati, lista_errori)
        """
        # Carica lavori
        lavori = self._load_lavori(input_file)
        
        # Raggruppa per comune
        by_comune = defaultdict(list)
        for lavoro in lavori:
            by_comune[lavoro.comune].append(lavoro)
        
        # Processa ogni comune
        output_files = []
        errors = []
        
        for comune, comuni_lavori in by_comune.items():
            try:
                path = self._process_comune(comune, comuni_lavori)
                output_files.append(path)
            except Exception as e:
                errors.append(f"{comune}: {e}")
        
        return output_files, errors
    
    def _load_catalogo(self) -> list[Prestazione]:
        """Carica catalogo prestazioni."""
        # Parse CSV
        catalogo = []
        
        # Il CSV usa ; come separatore e ha BOM
        df = pd.read_csv(
            CATALOGO_FILE,
            sep=";",
            encoding="utf-8-sig",
            skiprows=1,  # Salta header
        )
        
        for _, row in df.iterrows():
            if pd.isna(row.get("Codice Elenco Compensi")):
                continue
            
            try:
                prezzo_str = str(row.get("Prezzo scontato [€]", "")).replace("€", "").replace(",", ".").strip()
                prezzo = float(prezzo_str) if prezzo_str else 0.0
            except (ValueError, TypeError):
                prezzo = 0.0
            
            catalogo.append(Prestazione(
                codice_sap=str(row.get("Codice SAP/Glovia", "")),
                codice_elanco=str(row.get("Codice Elenco Compensi", "")),
                definizione=str(row.get("DEFINIZIONE", "")),
                prezzo_unitario=0.0,  # Non necessario per ora
                prezzo_scontato=prezzo,
                um=str(row.get("u.m.", "")),
            ))
        
        return catalogo
    
    def _load_mapping(self) -> dict[str, str]:
        """Carica mapping manuale."""
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get("mapping_manuali", {})
    
    def _load_lavori(self, input_file: str) -> list[Lavoro]:
        """Carica nuovi lavori da CSV."""
        df = pd.read_csv(
            input_file,
            sep=";",
            encoding="utf-8-sig",
        )
        
        lavori = []
        for _, row in df.iterrows():
            # Estrai quantità dalla descrizione
            descrizione = str(row.get(CSV_COLS["descrizione"], ""))
            quantita = extract_quantity(descrizione)
            
            lavori.append(Lavoro(
                codice=str(row.get(CSV_COLS["codice"], "")),
                data_richiesta=str(row.get(CSV_COLS["data_richiesta"], "")),
                data_esecuzione=str(row.get(CSV_COLS["data_esecuzione"], "")),
                comune=str(row.get(CSV_COLS["comune"], "")),
                via=str(row.get(CSV_COLS["via"], "")),
                descrizione=descrizione,
                quantita=quantita,
            ))
        
        return lavori
    
    def _process_comune(self, comune: str, lavori: list[Lavoro]) -> Path:
        """Processa tutti i lavori di un comune."""
        # Match ogni lavoro
        match_results = {}
        log_rows = []
        
        for lavoro in lavori:
            match_result = self.matcher.match(lavoro.descrizione)
            match_result.quantita = lavoro.quantita
            
            match_results[lavoro.codice] = match_result
            
            # Log row
            log_rows.append({
                "codice": lavoro.codice,
                "descrizione": lavoro.descrizione,
                "codice_matchato": match_result.codice_elanco or "",
                "metodo": match_result.metodo.value,
                "confidence": match_result.confidence,
                "quantita": match_result.quantita,
            })
        
        # Genera Excel
        output_path = genera_preventivo(
            comune=comune,
            lavori=lavori,
            match_results=match_results,
            regione=self.regione,
        )
        
        # Salva log
        log_path = self._save_log(log_rows)
        
        return output_path
    
    def _save_log(self, rows: list[dict]) -> Path:
        """Salva log matching."""
        log_path = Path("log_matching.csv")
        
        if rows:
            with open(log_path, "w", newline="", encoding="utf-8") as f:
                writer = csv.DictWriter(
                    f,
                    fieldnames=["codice", "descrizione", "codice_matchato", "metodo", "confidence", "quantita"]
                )
                writer.writeheader()
                writer.writerows(rows)
        
        return log_path
```

**Step 2: Create genera_preventivi.py (entry point)**

```python
"""Entry point CLI."""

from cli import parse_args, validate_args, print_banner, print_summary
from main import PreventiviGenerator


def main():
    args = parse_args()
    
    if not validate_args(args):
        return 1
    
    print_banner()
    
    if args.verbose:
        print(f"📂 Input: {args.input}")
        print(f"📋 Template: {args.template}")
        print(f"🤖 AI: {'Attivo' if args.use_ai else 'Disattivo'}")
        print()
    
    # Crea generatore e run
    generator = PreventiviGenerator(
        regione=args.template,
        use_ai=args.use_ai,
    )
    
    output_files, errors = generator.run(args.input)
    
    # Riepilogo
    print_summary(len(output_files), len(errors), Path("log_matching.csv"))
    
    return 0 if not errors else 1


if __name__ == "__main__":
    exit(main())
```

**Step 3: Commit**

```bash
git add main.py genera_preventivi.py
git commit -m "feat: add main orchestrator and CLI entry point"
```

---

## Task 8: Script Init

**Files:**
- Create: `lavoro/run.sh`

```bash
#!/bin/bash
# Genera preventivi con setup automatico

set -e

echo "📦 Installazione dipendenze..."
pip install -r requirements.txt

echo "🚀 Esecuzione generatore..."
python genera_preventivi.py --input nuovi_lavori.csv --template milano

echo "✅ Fatto!"
```

**Step 1: Create run.sh**

```bash
chmod +x run.sh
git add run.sh
git commit -m "chore: add run script"
```

---

## Task 9: README

**Files:**
- Create: `lavoro/README.md`

```markdown
# Generatore Preventivi Excel

Script Python per generare preventivi Excel automatizzando il matching tra descrizioni lavori e catalogo prestazioni.

## Setup

```bash
pip install -r requirements.txt
```

## Utilizzo

```bash
# Milano (default)
python genera_preventivi.py --input nuovi_lavori.csv --template milano

# Liguria
python genera_preventivi.py --input nuovi_lavori.csv --template liguria

# Con AI fallback
python genera_preventivi.py --input nuovi_lavori.csv --template milano --use-ai

# Help
python genera_preventivi.py --help
```

## Output

- `output/Preventivo_{Comune}_{Data}.xlsx`
- `log_matching.csv`
```

**Step 1: Commit**

```bash
git add README.md
git commit -m "docs: add README"
```

---

## Task 10: Test Suite Completa

**Files:**
- Modify: `lavoro/tests/`

Esegui tutti i test per validare integrazione:

```bash
python -m pytest tests/ -v --tb=short
```

Verifica che tutti i moduli funzionino insieme.

---

## Task 11: Test End-to-End con Dati Reali

Testa con i file CSV reali:

```bash
python genera_preventivi.py --input nuovi_lavori.csv --template milano --verbose
```

Verifica:
1. File Excel generato in `output/`
2. `log_matching.csv` creato
3. Fogli nominati correttamente (codici lavoro)
4. Foglio Totale con righe dati

---

## Eseguire il Plan

**Plan completo e salvato in:** `docs/plans/2026-05-14-preventivi-implementation-plan.md`

**Due opzioni di esecuzione:**

**1. Subagent-Driven (questa sessione)**
I dispatch subagent freschi per ogni task, revisione tra i task, iterazione veloce.

**2. Parallel Session (separata)**
Apri nuova sessione con `executing-plans`, esecuzione batch con checkpoints.

Quale approccio preferisci?