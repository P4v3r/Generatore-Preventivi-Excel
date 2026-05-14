"""Generazione file Excel da template."""

from pathlib import Path
from datetime import date
from typing import List, Optional
import zipfile
import re
import shutil

import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from config import (
    TEMPLATE_MILANO,
    TEMPLATE_LIGURIA,
    OUTPUT_DIR,
    CELLS,
    COL_QUANTITA,
    TOTALE_SHEET,
    TOTALE_HEADER_ROWS,
    TEMPLATE_WORK_SHEET,
    CELL_TOTALE,
)
from models import Lavoro, MatchResult


# Content type: template (.xltx) vs file normale (.xlsx)
TEMPLATE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
SHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'


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
    
    # Verifica esistenza fogli
    if TEMPLATE_WORK_SHEET not in wb.sheetnames:
        raise ValueError(f"Foglio '{TEMPLATE_WORK_SHEET}' non trovato nel template")
    
    if TOTALE_SHEET not in wb.sheetnames:
        raise ValueError(f"Foglio '{TOTALE_SHEET}' non trovato nel template")
    
    # Gestisci foglio Totale - conta righe esistenti
    ws_totale = wb[TOTALE_SHEET]
    next_totale_row = TOTALE_HEADER_ROWS + 1
    
    # Per ogni lavoro, crea un nuovo foglio
    for lavoro in lavori:
        # Clona il foglio template "cod"
        source_sheet = wb[TEMPLATE_WORK_SHEET]
        new_sheet = wb.copy_worksheet(source_sheet)
        new_sheet.title = lavoro.codice
        
        # Compila celle foglio lavoro
        _compile_lavoro_sheet(new_sheet, lavoro, match_results.get(lavoro.codice))
        
        # Aggiungi riga a foglio Totale
        _add_totale_row(ws_totale, next_totale_row, lavoro)
        next_totale_row += 1
    
    # Salva file
    output_dir = ensure_output_dir()
    filename = f"Preventivo_{comune}_{date.today().isoformat()}.xlsx"
    output_path = output_dir / filename
    
    wb.save(output_path)
    
    # Correzione content type: da template.main+xml a sheet.main+xml
    _fix_content_type(output_path)
    
    return output_path


def _fix_content_type(xlsx_path: Path):
    """Corregg e il content type del workbook da template a sheet.
    
    Il template .xltx ha content type 'template.main+xml' che Excel non accetta
    per file .xlsx normali. Questa funzione lo corregge.
    """
    temp_path = xlsx_path.with_suffix('.temp.xlsx')
    
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                
                if item.filename == '[Content_Types].xml':
                    content = data.decode('utf-8')
                    content = content.replace(TEMPLATE_CONTENT_TYPE, SHEET_CONTENT_TYPE)
                    data = content.encode('utf-8')
                
                zout.writestr(item, data)
    
    shutil.move(str(temp_path), str(xlsx_path))


def _compile_lavoro_sheet(ws: Worksheet, lavoro: Lavoro, match_result: Optional[MatchResult]):
    """Compila le celle del foglio lavoro."""
    # Celle base (colonna B)
    ws[CELLS["codice"]] = lavoro.codice
    ws[CELLS["data_richiesta"]] = lavoro.data_richiesta
    ws[CELLS["data_esecuzione"]] = lavoro.data_esecuzione
    ws[CELLS["comune"]] = lavoro.comune
    ws[CELLS["via"]] = lavoro.via
    ws[CELLS["descrizione"]] = lavoro.descrizione
    
    # Se c'è match, trova riga prestazione e inserisci quantità
    if match_result and match_result.codice_elanco:
        _insert_quantita(ws, match_result.codice_elanco, match_result.quantita)


def _insert_quantita(ws: Worksheet, codice_elanco: str, quantita: int):
    """Trova riga con codice elanco e inserisci quantità nella colonna D."""
    # Colonna D = index 4 in openpyxl (1-based)
    from openpyxl.utils import column_index_from_string
    d_col = column_index_from_string(COL_QUANTITA)  # = 4
    
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        cell = row[2]  # Colonna C (indice 2, 0-based) - cerca il codice
        if cell.value and codice_elanco in str(cell.value):
            # Trovato! Inserisci quantità nella colonna D
            ws.cell(row=cell.row, column=d_col).value = quantita
            return


def _add_totale_row(ws_totale: Worksheet, row: int, lavoro: Lavoro):
    """Aggiunge riga al foglio Totale."""
    # B: Numero ordine, C: Data, D: Importo (riferimento a I8 del foglio lavoro)
    ws_totale[f"B{row}"] = lavoro.codice
    ws_totale[f"C{row}"] = lavoro.data_esecuzione
    ws_totale[f"D{row}"] = f"='{lavoro.codice}'!{CELL_TOTALE}"  # Riferimento a I8 del foglio lavoro