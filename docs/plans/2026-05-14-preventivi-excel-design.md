# Design: Generatore Preventivi Excel

**Data:** 2026-05-14  
**Stato:** Bozza per approvazione

---

## 1. PANORAMICA

Script Python CLI per generare preventivi Excel da template regionali (Milano/Liguria), matchando automaticamente le descrizioni dei lavori alle prestazioni del catalogo.

### Input
| File | Descrizione |
|------|-------------|
| `nuovi_lavori.csv` | Lavori da elaborare (Codice, Date, Comune, Via, Descrizione) |
| `catalogo_prestazioni.csv` | ~540 voci prezzi (Codice Elenco, Definizione, Prezzi) |
| `mapping_manuali.json` | Dizionario match diretti |
| `template-mi.xltx` | Template Milano |
| `template-lig.xltx` | Template Liguria |

### Output
| File | Contenuto |
|------|-----------|
| `Preventivo_{Comune}_{YYYY-MM-DD}.xlsx` | Un file per comune |
| `log_matching.csv` | Tracciamento match (Codice, Descrizione, Matchato, Metodo, Confidence) |

---

## 2. ARCHITETTURA

```
genera_preventivi.py
├── config.py              # Costanti, path, configurazione
├── models.py              # Data classes (Lavoro, Prestazione, MatchResult)
├── matcher.py             # Pipeline matching (manual → fuzzy → AI)
├── extractor.py           # Estrazione quantità dalla descrizione
├── excel_generator.py     # Generazione file Excel
├── cli.py                 # Entry point CLI
└── main.py                # Orchestrazione principale
```

### Pipeline Matching (3 step)

```
Descrizione Lavoro
        │
        ▼
┌───────────────────┐
│ 1. Dizionario     │ → Match esatto/parziale su mapping_manuali.json
│    Manuale        │
└─────────┬─────────┘
          │ No match
          ▼
┌───────────────────┐
│ 2. Fuzzy Match    │ → rapidfuzz su catalogo.DEFINIZIONE
│    (score ≥ 85)   │   Score ≥ 85 → accetta
└─────────┬─────────┘
          │ No match / score < 85
          ▼
┌───────────────────┐
│ 3. AI Fallback     │ → Chiamata CF Workers (opzionale, flag USE_AI)
│    (score < 85)   │   Ritorna Codice Elenco
└───────────────────┘
```

---

## 3. LOGICHE SPECIFICHE

### 3.1 Estrazione Quantità

Pattern regex per estrarre quantità dalla descrizione:

```python
patterns = [
    r'(\d+)\s*unita',      # "5 unità", "3 UNITA"
    r'(\d+)\s*pz',         # "5 pz", "3 PZ"
    r'n\.\s*(\d+)',        # "n. 5", "N. 5"
    r'quantita\s*(\d+)',   # "quantità 5"
    r'(\d+)\s*pezzi',      # "5 pezzi"
]
```

- Se nessun pattern match → Quantità = 1
- Loggare in `log_matching.csv` la quantità estratta

### 3.2 Generazione Excel

Per ogni comune con almeno 1 lavoro:

1. **Caricare template regionale** (`template-mi` o `template-lig`)
2. **Creare foglio "Totale"** (prima riga intestazione, vuoto)
3. **Per ogni lavoro del comune:**
   - Clonare il template "LAV_TEMPLATE" in nuovo foglio nominato col Codice
   - Compilare celle:
     - `DE2` = Codice lavoro
     - `C3` = Data richiesta
     - `C4` = Data esecuzione
     - `CDE5` = Comune
     - `CDE6` = Via
     - `CDE7` = Descrizione originale
   - **Trovare riga prestazione** con Codice Elenco matchato
   - **Inserire Quantità** estratta (cella colonna "Quantità")
   - L'Importo si calcola con formula esistente (Q × Prezzo scontato)
4. **Aggiornare foglio Totale:**
   - Aggiungere riga con: B1=Codice, C1=Data, D1=Importo (riferimento a I8)
5. **Salvare** `Preventivo_{Comune}_{Data}.xlsx`

### 3.3 Celle Foglio Lavoro

| Cella | Contenuto |
|-------|-----------|
| DE2 | Codice lavoro |
| C3 | Data richiesta |
| C4 | Data esecuzione |
| CDE5 | Comune |
| CDE6 | Via |
| CDE7 | Descrizione |
| Colonna "Quantità" | Quantità estratta (riga Codice Elenco matchato) |

### 3.4 Log Matching

CSV con colonne:
```csv
Codice,Descrizione_Originale,Codice_Matchato,Metodo,Confidence,Quantita_Estratta
LAV001,Installazione 5 pali,A.EC.2.1.10,fuzzy,92,5
```

---

## 4. CONFIGURAZIONE

### `config.py`

```python
TEMPLATE_DIR = Path(".")
TEMPLATE_MILANO = "template-mi.xltx"
TEMPLATE_LIGURIA = "template-lig.xltx"

CATALOGO_FILE = "catalogo_prestazioni.csv"
LAVORI_FILE = "nuovi_lavori.csv"
MAPPING_FILE = "mapping_manuali.json"

FUZZY_THRESHOLD = 85
USE_AI = False  # Flag per attivare AI fallback

# Colonne CSV input
CSV_COLS = {
    "codice": "Cod SF - PL",
    "data_richiesta": "Data Richiesta",
    "data_esecuzione": "Data Esecuzione",
    "comune": "Comune",
    "via": "Località",
    "descrizione": "Descrizione"
}
```

---

## 5. FLUSSO PRINCIPALE (CLI)

```
$ python genera_preventivi.py --input nuovi_lavori.csv --template milano

1. Parse argomenti CLI
2. Carica catalogo → DataFrame
3. Carica mapping_manuali.json
4. Carica nuovi_lavori.csv
5. Raggruppa per comune
6. Per ogni comune:
   a. Carica template regionale
   b. Crea foglio "Totale"
   c. Per ogni lavoro:
      - Matcha descrizione (3 step)
      - Estrai quantità
      - Crea/clona foglio
      - Compila celle
      - Aggiungi riga a Totale
   d. Salva Preventivo_{Comune}_{data}.xlsx
7. Genera log_matching.csv
8. Stampa riepilogo
```

---

## 6. GESTIONE ERRORI

| Scenario | Comportamento |
|----------|----------------|
| CSV non trovato | Exit con messaggio chiaro |
| Nessun match trovato | Logga come "NO_MATCH", skippa quel lavoro, continua |
| Foglio già esiste | Sovrascrivi |
| Quantità non estraibile | Default a 1, logga warning |
| AI fallisce | Logga errore, skippa, continua |

---

## 7. STRUTTURA FILE OUTPUT

```
lavoro/
├── genera_preventivi.py    # Script principale
├── config.py               # Configurazione
├── models.py               # Data classes
├── matcher.py              # Logica matching
├── extractor.py            # Estrazione quantità
├── excel_generator.py      # Generazione Excel
├── nuovi_lavori.csv        # Input
├── catalogo_prestazioni.csv # Catalogo
├── mapping_manuali.json    # Mapping manuale
├── template-mi.xltx         # Template Milano
├── template-lig.xltx        # Template Liguria
├── output/                 # Output preventivi
│   ├── Preventivo_Milano_2024-05-14.xlsx
│   └── Preventivo_Genoa_2024-05-14.xlsx
└── log_matching.csv        # Log matching
```

---

## 8. DIPENDENZE

```txt
pandas>=2.0
openpyxl>=3.1
rapidfuzz>=3.0
requests>=2.28  # Solo se USE_AI=True
```

---

## 9. COMANDI UTILI

```bash
# Generazione base (senza AI)
python genera_preventivi.py --input nuovi_lavori.csv --template milano

# Con AI fallback
python genera_preventivi.py --input nuovi_lavori.csv --template milano --use-ai

# Help
python genera_preventivi.py --help
```

---

## 10. APPROCCI ALTERNATIVI CONSIDERATI

### Approccio B: Streamlit UI
- **Pro:** User-friendly, nessun terminale
- **Contro:** Richiede browser, più complesso da debuggare
- **Decisione:** Posticipato a fase 2

### Approccio C: Database SQLite
- **Pro:** Storico preventivi
- **Contro:** Overhead per use case semplice
- **Decisione:** Non necessario per ora

---

## DOMANDE APERTE

1. ~~Template Milano/Liguria selezione~~ → Utente specifica via CLI
2. ~~Quantità sempre 1~~ → Estrazione da descrizione
3. ~~Foglio Totale precompilato~~ → Righe da aggiungere manualmente (ma script aggiunge automaticamente)

---

**Prossimo passo:** Se approvi, scrivo implementation plan dettagliato con task.