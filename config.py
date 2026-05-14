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

# Celle foglio lavoro (colonna C per i valori, DE per codice)
CELLS = {
    "codice": "D2",        # Colonna D (cella unita DE)
    "data_richiesta": "C3", # Colonna C
    "data_esecuzione": "C4", # Colonna C
    "comune": "C5",         # Colonna C (cella unita CDE)
    "via": "C6",            # Colonna C (cella unita CDE)
    "descrizione": "C7",    # Colonna C (cella unita CDE)
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

# Template sheet names
TEMPLATE_WORK_SHEET = "cod"