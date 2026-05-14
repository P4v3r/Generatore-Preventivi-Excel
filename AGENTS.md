# PROJECT KNOWLEDGE BASE

**Generated:** 2026-05-14

## OVERVIEW
Project: **Generatore Preventivi Excel**
Stack: Python 3.14+ • pandas • openpyxl • rapidfuzz

## STRUCTURE
```
lavoro/
├── catalogo_prestazioni.csv   # ~540 voci prestazioni prezzi
├── nuovi_lavori.csv           # Input: nuovi lavori da elaborare
├── mapping_manuali.json       # Mapping descrizione → codice manuale
├── template-mi.xltx           # Template Excel Milano (prezzi scontati)
├── template-lig.xltx          # Template Excel Liguria (prezzi scontati)
├── docs/
│   └── plans/                 # Design documents
└── (genera_preventivi.py)    # Script da creare
```

## COMMANDS
| Action | Command |
|--------|---------|
| Install deps | `pip install pandas openpyxl rapidfuzz` |
| Run | `python genera_preventivi.py --input nuovi_lavori.csv --template milano` |

## CODING STANDARDS
- **Language:** Python 3.14+
- **Style:** PEP 8, type hints, docstrings
- **Modules:** config.py, models.py, matcher.py, extractor.py, excel_generator.py, cli.py

## INPUT FILES

### nuovi_lavori.csv
Colonne: `Cod SF - PL`, `Data Richiesta`, `Data Esecuzione`, `Comune`, `Località`, `Descrizione`

### catalogo_prestazioni.csv
Colonne: `Codice SAP/Glovia`, `Codice Elenco Compensi`, `Quantità`, `u.m.`, `DEFINIZIONE`, `Prezzo Unitario [€]`, `Prezzo scontato [€]`, `Importo [€]`

### mapping_manuali.json
```json
{"descrizione": "codice_elanco"}
```

## OUTPUT FILES
- `Preventivo_{Comune}_{YYYY-MM-DD}.xlsx` - Uno per comune
- `log_matching.csv` - Tracciamento matching

## TEMPLATE EXCEL STRUCTURE
- **Foglio lavoro:** Celle compilabili (DE2, C3, C4, CDE5, CDE6, CDE7), righe prestazioni precompilate
- **Foglio Totale:** Riga 1-2 intestazione, Riga 3+ dati (Codice, Data, Importo)

## WHERE TO LOOK
- **Input:** `lavoro/*.csv`
- **Templates:** `lavoro/template-*.xltx`
- **Design:** `docs/plans/2026-05-14-preventivi-excel-design.md`

## NOTES
- Template sono .xltx (Excel template), NON .xlsx
- Prezzi scontati già inclusi nei template (Milano vs Liguria)
- Fuzzy threshold: 85 per match automatico
- Quantità estratta da descrizione con regex pattern