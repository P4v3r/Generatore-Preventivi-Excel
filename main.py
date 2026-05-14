"""Orchestrazione principale del generatore."""

import csv
import json
from pathlib import Path
from collections import defaultdict

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
        
        if not lavori:
            return [], ["Nessun lavoro trovato nel file CSV"]
        
        # Raggruppa per comune
        by_comune = defaultdict(list)
        for lavoro in lavori:
            by_comune[lavoro.comune].append(lavoro)
        
        # Processa ogni comune
        output_files = []
        errors = []
        all_log_rows = []  # Accumula log per tutti i comuni
        
        for comune, comuni_lavori in by_comune.items():
            try:
                path, log_rows = self._process_comune(comune, comuni_lavori)
                output_files.append(path)
                all_log_rows.extend(log_rows)  # Accumula log
            except Exception as e:
                errors.append(f"{comune}: {e}")
        
        # Salva log una volta sola alla fine
        self._save_log(all_log_rows)
        
        return output_files, errors
    
    def _load_catalogo(self) -> list[Prestazione]:
        """Carica catalogo prestazioni."""
        catalogo = []
        
        # Il CSV usa ; come separatore e ha BOM
        df = pd.read_csv(
            CATALOGO_FILE,
            sep=";",
            encoding="utf-8-sig",
        )
        
        # Salta la prima riga che è l'header "PREST_SP"
        for idx, row in df.iterrows():
            if idx == 0:
                continue
            
            codice_elanco = str(row.get("Codice Elenco Compensi", "")).strip()
            
            # Salta righe senza codice o con codici speciali
            if not codice_elanco or codice_elanco == "nan" or "PREST" in codice_elanco:
                continue
            
            # Estrai prezzo scontato
            try:
                prezzo_str = str(row.get("Prezzo scontato [€]", "")).replace("€", "").replace(",", ".").strip()
                if prezzo_str and prezzo_str != "nan" and '#VALUE!' not in prezzo_str:
                    prezzo = float(prezzo_str)
                else:
                    prezzo = 0.0
            except (ValueError, TypeError):
                prezzo = 0.0
            
            catalogo.append(Prestazione(
                codice_sap=str(row.get("Codice SAP/Glovia", "")).strip(),
                codice_elanco=codice_elanco,
                definizione=str(row.get("DEFINIZIONE", "")).strip(),
                prezzo_unitario=0.0,
                prezzo_scontato=prezzo,
                um=str(row.get("u.m.", "")).strip(),
            ))
        
        return catalogo
    
    def _load_mapping(self) -> dict[str, str]:
        """Carica mapping manuale."""
        with open(MAPPING_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        # Il file potrebbe avere la struttura {"mapping_manuali": {...}} oppure direttamente il dict
        if isinstance(data, dict) and "mapping_manuali" in data:
            return data["mapping_manuali"]
        elif isinstance(data, dict):
            return data
        else:
            return {}
    
    def _load_lavori(self, input_file: str) -> list[Lavoro]:
        """Carica nuovi lavori da CSV."""
        df = pd.read_csv(
            input_file,
            sep=";",
            encoding="utf-8-sig",
        )
        
        lavori = []
        for _, row in df.iterrows():
            codice = str(row.get(CSV_COLS["codice"], "")).strip()
            if not codice or codice == "":
                continue
            
            # Estrai quantità dalla descrizione
            descrizione = str(row.get(CSV_COLS["descrizione"], ""))
            quantita = extract_quantity(descrizione)
            
            lavori.append(Lavoro(
                codice=codice,
                data_richiesta=str(row.get(CSV_COLS["data_richiesta"], "")),
                data_esecuzione=str(row.get(CSV_COLS["data_esecuzione"], "")),
                comune=str(row.get(CSV_COLS["comune"], "")),
                via=str(row.get(CSV_COLS["via"], "")),
                descrizione=descrizione,
                quantita=quantita,
            ))
        
        return lavori
    
    def _process_comune(self, comune: str, lavori: list[Lavoro]) -> tuple[Path, list[dict]]:
        """Processa tutti i lavori di un comune.
        
        Returns:
            (path_output, log_rows)
        """
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
        
        return output_path, log_rows
    
    def _save_log(self, rows: list[dict]):
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