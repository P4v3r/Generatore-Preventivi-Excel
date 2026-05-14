# Generatore Preventivi Excel

Script Python per generare preventivi Excel automatizzando il matching tra descrizioni lavori e catalogo prestazioni.

## Setup

```bash
pip install -r requirements.txt
```

## Utilizzo

```bash
# Milano (default)
python3 genera_preventivi.py --input nuovi_lavori.csv --template milano

# Liguria
python3 genera_preventivi.py --input nuovi_lavori.csv --template liguria

# Con AI fallback
python3 genera_preventivi.py --input nuovi_lavori.csv --template milano --use-ai

# Help
python3 genera_preventivi.py --help

# Verbose
python3 genera_preventivi.py --input nuovi_lavori.csv --verbose
```

## Output

- `output/Preventivo_{Comune}_{Data}.xlsx` - File Excel per comune
- `log_matching.csv` - Tracciamento matching

## Struttura Progetto

```
lavoro/
├── config.py           # Configurazione
├── models.py           # Data classes
├── matcher.py          # Pipeline matching
├── extractor.py        # Estrazione quantità
├── excel_generator.py # Generazione Excel
├── cli.py              # CLI interface
├── main.py             # Orchestrator
├── genera_preventivi.py # Entry point
├── catalogo_prestazioni.csv
├── nuovi_lavori.csv
├── mapping_manuali.json
├── template-mi.xltx
├── template-lig.xltx
└── output/             # Preventivi generati
```

## Pipeline Matching

1. **Dizionario Manuale** - Match esatto/parziale su `mapping_manuali.json`
2. **Fuzzy Matching** - `rapidfuzz` con threshold 85
3. **AI Fallback** - Cloudflare Workers (attivabile con `--use-ai`)
