"""Generazione file Excel da template."""

from pathlib import Path
from datetime import date
from typing import List, Optional
import zipfile
import re
import shutil

from config import (
    TEMPLATE_MILANO,
    TEMPLATE_LIGURIA,
    OUTPUT_DIR,
    CELLS,
    COL_QUANTITA,
    TEMPLATE_WORK_SHEET,
)
from models import Lavoro, MatchResult

# Content type: template (.xltx) vs file normale (.xlsx)
TEMPLATE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
SHEET_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'

# Shared strings cache (caricato dal template)
SHARED_STRINGS: list[str] = []


def _load_shared_strings(ss_xml_data: bytes) -> list[str]:
    """Carica le shared strings dal template.
    
    Le shared strings sono usate per memorizzare testo ripetuto nelle celle.
    Ogni cella con t="s" contiene un indice a questa lista.
    """
    if not ss_xml_data:
        return []
    
    import xml.etree.ElementTree as ET
    try:
        root = ET.fromstring(ss_xml_data)
        ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        strings = []
        for si in root.findall('ns:si', ns):
            t = si.find('ns:t', ns)
            if t is not None and t.text:
                strings.append(t.text)
            else:
                # Handle rich text
                texts = []
                for r in si.findall('.//ns:t', ns):
                    if r.text:
                        texts.append(r.text)
                strings.append(''.join(texts) if texts else '')
        return strings
    except ET.ParseError:
        return []


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
    template_path = get_template_path(regione)
    output_dir = ensure_output_dir()
    filename = f"Preventivo_{comune}_{date.today().isoformat()}.xlsx"
    output_path = output_dir / filename
    
    # Costruisci il workbook usando approccio ZIP puro
    _build_workbook_zip(template_path, output_path, lavori, match_results)
    
    # Correzione content type: da template.main+xml a sheet.main+xml
    _fix_content_type(output_path)
    
    return output_path


def _build_workbook_zip(
    template_path: Path,
    output_path: Path,
    lavori: List[Lavoro],
    match_results: dict
):
    """
    Costruisce il workbook dall'XML del template.
    
    Approccio: manipola il ZIP direttamente per preservare drawings, VML, etc.
    """
    # Leggi tutti i file dal template
    with zipfile.ZipFile(template_path, 'r') as ztemplate:
        template_files = {}
        for name in ztemplate.namelist():
            template_files[name] = ztemplate.read(name)
        
        # Carica shared strings per la ricerca dei codici elanco
        global SHARED_STRINGS
        SHARED_STRINGS = _load_shared_strings(template_files.get('xl/sharedStrings.xml', b''))
    
    # Prepara i dati per ogni lavoro
    sheet_data_list = []
    for lavoro in lavori:
        mr = match_results.get(lavoro.codice)
        data = {
            'codice': lavoro.codice,
            'data_richiesta': lavoro.data_richiesta,
            'data_esecuzione': lavoro.data_esecuzione,
            'comune': lavoro.comune,
            'via': lavoro.via,
            'descrizione': lavoro.descrizione,
            'quantita': mr.quantita if mr and mr.codice_elanco else None,
            'codice_elanco': mr.codice_elanco if mr else None,
        }
        sheet_data_list.append(data)
    
    # Prepara i file per il nuovo ZIP
    new_files = {}
    
    # 1. Copia tutti i file dal template eccetto sheet1 (template "cod")
    for filename, data in template_files.items():
        if filename == 'xl/worksheets/sheet1.xml':
            continue  # Sostituiremo con i fogli lavoro
        if filename == 'xl/worksheets/_rels/sheet1.xml.rels':
            continue  # Sostituiremo con le rels dei lavori
        new_files[filename] = data
    
    # 2. Crea i fogli lavoro (partendo da 1)
    # Sheet1 = primo lavoro
    # Sheet2 = Totale
    # Sheet3+ = lavori successivi
    
    # Prendi lo sheet XML dal template come base
    sheet_template_xml = template_files['xl/worksheets/sheet1.xml'].decode('utf-8')
    sheet_template_rels = template_files['xl/worksheets/_rels/sheet1.xml.rels'].decode('utf-8')
    
    # Crea i fogli lavoro
    # Primo lavoro -> sheet1.xml
    modified = _modify_sheet_xml(sheet_template_xml, sheet_data_list[0])
    new_files['xl/worksheets/sheet1.xml'] = modified.encode('utf-8')
    new_files['xl/worksheets/_rels/sheet1.xml.rels'] = sheet_template_rels.encode('utf-8')
    
    # Fogli successivi -> sheet3, sheet4, etc.
    for i in range(1, len(sheet_data_list)):
        sheet_num = 2 + i  # sheet3, sheet4, etc.
        sheet_name = f'xl/worksheets/sheet{sheet_num}.xml'
        rels_name = f'xl/worksheets/_rels/sheet{sheet_num}.xml.rels'
        
        modified = _modify_sheet_xml(sheet_template_xml, sheet_data_list[i])
        new_files[sheet_name] = modified.encode('utf-8')
        new_files[rels_name] = sheet_template_rels.encode('utf-8')
    
    # 3. Aggiorna workbook.xml
    workbook_xml = template_files['xl/workbook.xml'].decode('utf-8')
    
    # Rimuovi externalReferences e definedNames che referenziano "cod"
    workbook_xml = re.sub(r'<externalReferences>.*?</externalReferences>', '', workbook_xml, flags=re.DOTALL)
    workbook_xml = re.sub(r'<definedNames>.*?</definedNames>', '', workbook_xml, flags=re.DOTALL)
    
    # Sostituisci l'intera sezione <sheets> con i nuovi fogli
    # Ordine: fogli lavoro (1, 3, 4...) poi Totale (2)
    sheets_xml = '<sheets>'
    
    # Primo lavoro -> sheetId=1, rId10
    sheets_xml += f'<sheet name="{sheet_data_list[0]["codice"]}" sheetId="1" state="visible" r:id="rId10"/>'
    
    # Fogli successivi
    for i in range(1, len(sheet_data_list)):
        sheetId = 1 + i
        sheets_xml += f'<sheet name="{sheet_data_list[i]["codice"]}" sheetId="{sheetId}" state="visible" r:id="rId{10+i}"/>'
    
    # Totale -> sheetId=n+1, rId4 (esistente)
    total_sheetId = 1 + len(sheet_data_list)
    sheets_xml += f'<sheet name="Totale" sheetId="{total_sheetId}" state="visible" r:id="rId4"/>'
    
    sheets_xml += '</sheets>'
    
    # Sostituisci tutto tra <sheets> e </sheets> nel workbook.xml originale
    workbook_xml = re.sub(r'<sheets>.*?</sheets>', sheets_xml, workbook_xml, flags=re.DOTALL)
    new_files['xl/workbook.xml'] = workbook_xml.encode('utf-8')
    
    # 4. Aggiorna workbook.xml.rels
    workbook_rels = template_files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    
    # Rimuovi la relazione per sheet1 (era "cod")
    workbook_rels = re.sub(
        r'<Relationship[^>]*worksheets/sheet1\.xml"[^>]*/>', 
        '', 
        workbook_rels
    )
    
    # Aggiungi le relazioni per i nuovi fogli
    new_rels = '<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    for i in range(1, len(sheet_data_list)):
        sheet_num = 2 + i  # sheet3, sheet4, etc.
        new_rels += f'<Relationship Id="rId{10+i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{sheet_num}.xml"/>'
    
    workbook_rels = workbook_rels.replace('</Relationships>', new_rels + '</Relationships>')
    new_files['xl/_rels/workbook.xml.rels'] = workbook_rels.encode('utf-8')
    
    # 5. Aggiorna [Content_Types].xml
    content_types = template_files['[Content_Types].xml'].decode('utf-8')
    
    # Rimuovi l'override per sheet1 (era "cod")
    content_types = re.sub(
        r'<Override[^>]*worksheets/sheet1\.xml"[^>]*/>', 
        '', 
        content_types
    )
    
    # Aggiungi override per i nuovi fogli
    # sheet1 è usato per il primo lavoro
    new_overrides = '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    
    for i in range(1, len(sheet_data_list)):
        sheet_num = 2 + i  # sheet3, sheet4, etc.
        new_overrides += f'<Override PartName="/xl/worksheets/sheet{sheet_num}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    
    content_types = content_types.replace('</Types>', new_overrides + '</Types>')
    
    new_files['[Content_Types].xml'] = content_types.encode('utf-8')

    # 6. Lascia il foglio Totale (sheet2.xml) identico al template - NON popolarlo
    # Il foglio Totale deve rimanere vuoto come nel template originale
    # (nessuna modifica necessaria - sheet2.xml è già copiato dal template)

    # 8. Scrivi il nuovo ZIP
    temp_path = output_path.with_suffix('.build.xlsx')
    with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for filename, data in new_files.items():
            zout.writestr(filename, data)
    
    shutil.move(str(temp_path), str(output_path))


def _modify_sheet_xml(sheet_xml: str, data: dict) -> str:
    """
    Modifica solo i valori delle celle nel foglio XML.
    
    NON tocca stili, bordi, drawings, VML, merge cells.
    """
    content = sheet_xml
    
    # D2 - Codice
    content = _set_cell_value(content, 'D2', data['codice'])
    
    # C3 - Data Richiesta
    content = _set_cell_value(content, 'C3', data['data_richiesta'])
    
    # C4 - Data Esecuzione
    content = _set_cell_value(content, 'C4', data['data_esecuzione'])
    
    # C5 - Comune (cella unita CDE5)
    content = _set_cell_value(content, 'C5', data['comune'])
    
    # C6 - Via/Località (cella unita CDE6)
    content = _set_cell_value(content, 'C6', data['via'])
    
    # C7 - Descrizione (cella unita CDE7)
    content = _set_cell_value(content, 'C7', data['descrizione'])
    
    # Quantità in colonna D - se specificata
    if data.get('codice_elanco') and data.get('quantita'):
        content = _insert_quantita_in_sheet(content, data['codice_elanco'], data['quantita'])
    
    return content


def _set_cell_value(sheet_xml: str, cell_ref: str, value) -> str:
    """Imposta il valore di una cella nel foglio XML.
    
    Preserva formule esistenti (<f>) e attributi della cella.
    """
    if value is None:
        return sheet_xml
    
    str_value = str(value)
    
    # Strategy: First check for self-closing cells (<c ... />), 
    # then for cells with content (<c ...>...</c>).
    # IMPORTANT: Self-closing pattern must be checked first because
    # non-self-closing regex can incorrectly match across cells.
    
    # Check for self-closing cell first
    self_closing_pattern = rf'<c r="{re.escape(cell_ref)}"([^>]*)/>'
    self_closing_match = re.search(self_closing_pattern, sheet_xml)
    
    if self_closing_match:
        # Convert self-closing to cell with value
        attrs = self_closing_match.group(1)
        self_closing = self_closing_match.group(0)
        open_tag = f'<c r="{cell_ref}"{attrs}>'
        
        # Estrai attributi
        t_match = re.search(r't="([^"]+)"', open_tag)
        style_match = re.search(r's="(\d+)"', open_tag)
        existing_type = t_match.group(1) if t_match else None
        existing_style = f' s="{style_match.group(1)}"' if style_match else ''
        
        # Costruisci nuova cella
        if isinstance(value, (int, float)):
            new_inner = f'<v>{str_value}</v>'
        else:
            type_attr = ' t="str"' if not existing_type else ''
            new_inner = f'<v>{str_value}</v>'
        
        new_cell = f'<c r="{cell_ref}"{existing_style}{type_attr if not isinstance(value, (int, float)) else ""}>{new_inner}</c>'
        sheet_xml = sheet_xml.replace(self_closing, new_cell, 1)  # Replace only first occurrence
        return sheet_xml
    
    # Now check for non-self-closing cell with content
    # Use a more precise pattern that matches only the specific cell
    # Pattern: <c r="CELLREF" ... > content </c>
    # Must be careful not to match across cell boundaries
    non_self_pattern = rf'<c r="{re.escape(cell_ref)}"([^>]*)>'
    non_self_match = re.search(non_self_pattern, sheet_xml)
    
    if non_self_match:
        # Find the closing </c> after this opening tag
        # Use a non-greedy match that stops at the first </c>
        start_pos = non_self_match.end()  # Position after opening tag
        
        # Find the closing tag - search forward from start_pos
        close_pos = sheet_xml.find('</c>', start_pos)
        if close_pos == -1:
            return sheet_xml  # No closing tag, shouldn't happen
        
        # Extract the full cell content
        open_tag_content = non_self_match.group(1)  # Attributes between r and >
        cell_content = sheet_xml[start_pos:close_pos]
        
        # Build replacement
        t_match = re.search(r't="([^"]+)"', non_self_match.group(0))
        style_match = re.search(r's="(\d+)"', non_self_match.group(0))
        existing_type = t_match.group(1) if t_match else None
        existing_style = f' s="{style_match.group(1)}"' if style_match else ''
        
        # Check if cell has formula
        has_formula = '<f' in cell_content and '</f>' in cell_content
        
        if has_formula:
            # Preserve formula, update value
            formula_match = re.search(r'<f[^>]*>.*?</f>', cell_content, re.DOTALL)
            formula = formula_match.group(0) if formula_match else ''
            new_inner = f'{formula}<v>{str_value}</v>'
        else:
            # No formula - just set value
            if isinstance(value, (int, float)):
                new_inner = f'<v>{str_value}</v>'
            else:
                type_attr = ' t="str"' if not existing_type else ''
                new_inner = f'<v>{str_value}</v>'
        
        old_cell = f'<c r="{cell_ref}"{open_tag_content}>{cell_content}</c>'
        new_cell = f'<c r="{cell_ref}"{existing_style}{' t="str"' if not isinstance(value, (int, float)) and not existing_type else ''}>{new_inner}</c>'
        
        sheet_xml = sheet_xml.replace(old_cell, new_cell, 1)
        return sheet_xml
    
    # Cell doesn't exist - insert it
    row_num = int(re.search(r'(\d+)', cell_ref).group(1))
    new_cell = f'<c r="{cell_ref}"><v>{str_value}</v></c>'
    
    # Find the row
    row_pattern = rf'(<row[^>]*r="{row_num}"[^>]*>)(.*?)(</row>)'
    row_match = re.search(row_pattern, sheet_xml, re.DOTALL)
    
    if row_match:
        row_content = row_match.group(2)
        new_row_content = row_content + new_cell
        new_row = row_match.group(1) + new_row_content + row_match.group(3)
        sheet_xml = sheet_xml.replace(row_match.group(0), new_row)
    
    return sheet_xml


def _insert_quantita_in_sheet(sheet_xml: str, codice_elanco: str, quantita: int) -> str:
    """Trova la riga con il codice elanco e inserisci la quantità nella colonna D.
    
    I codici elanco sono in colonna C del template (colonna "Cod.Elenco" nel template).
    La quantità viene inserita nella colonna D della stessa riga.
    """
    # Cerca la cella in colonna C che contiene il codice elanco (tramite shared string)
    # Le celle C hanno t="s" (shared string) e contengono l'indice nella shared strings table
    
    # Pattern: cerca <c r="C{N}" ... t="s"><v>{INDEX}</v></c>
    c_cell_pattern = r'<c r="C(\d+)"[^>]*t="s"[^>]*><v>(\d+)</v></c>'
    
    # Prima passata: cerca la riga che contiene il codice elanco
    found_row = None
    
    for match in re.finditer(c_cell_pattern, sheet_xml):
        row_num = int(match.group(1))
        str_index = int(match.group(2))
        
        # Verifica se il contenuto della shared string corrisponde al codice elanco
        if str_index < len(SHARED_STRINGS):
            shared_str = SHARED_STRINGS[str_index]
            if codice_elanco == shared_str or codice_elanco.upper() == shared_str.upper():
                found_row = row_num
                break
    
    # Se non trovato con match esatto, prova ricerca case-insensitive
    if found_row is None:
        for match in re.finditer(c_cell_pattern, sheet_xml):
            row_num = int(match.group(1))
            str_index = int(match.group(2))
            
            if str_index < len(SHARED_STRINGS):
                shared_str = SHARED_STRINGS[str_index]
                if codice_elanco.upper() in shared_str.upper() or shared_str.upper() in codice_elanco.upper():
                    found_row = row_num
                    break
    
    if found_row:
        # Inserisci/aggiorna la quantità nella colonna D della stessa riga
        sheet_xml = _set_cell_value(sheet_xml, f'D{found_row}', quantita)
    else:
        # Log per debug: codice non trovato
        import sys
        print(f"DEBUG: Codice elanco '{codice_elanco}' non trovato nel foglio. SHARED_STRINGS ha {len(SHARED_STRINGS)} elementi", file=sys.stderr)
    
    return sheet_xml


def _clear_cell_style(sheet_xml: str, cell_ref: str) -> str:
    """Rimuove lo stile da una cella, mantenendo solo r e t attribute.
    
    Questo elimina bordi/sfondi indesiderati dalle celle compilabili.
    """
    # Pattern per trovare la cella e rimuovere s="..." dal tag di apertura
    cell_pattern = rf'<c r="{re.escape(cell_ref)}"([^>]*)'
    
    def fix_cell(match):
        attrs = match.group(1)
        # Rimuovi s="..." mantenendo altri attributi
        attrs_no_style = re.sub(r's="[^"]*"\s*', '', attrs)
        return f'<c r="{cell_ref}"{attrs_no_style}'
    
    sheet_xml = re.sub(cell_pattern, fix_cell, sheet_xml)
    return sheet_xml


def _fix_content_type(xlsx_path: Path):
    """Corregg e il content type del workbook da template a sheet."""
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