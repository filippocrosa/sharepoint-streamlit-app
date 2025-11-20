import time
import io
import zipfile
import logging
from docx import Document
import openpyxl
import re
from datetime import datetime
from openpyxl.cell.cell import Cell
import locale



def check_consistency(excel_file: str, docx_file: str) -> tuple[bool, str, set[str]]:

    """
    Estrae tutte le stringhe racchiuse tra '{{' e '}}' da un file Word.

    Args:
        docx_path: Il percorso del file .docx.

    Returns:
        Una lista di stringhe trovate tra le parentesi graffe.
    """

    error_msg = []

    try:
        # Carica il documento
        document = Document(docx_file)
    except Exception as e:
        logging.warning(f"CHECK_1: Cannot read word file {e}")
        error_msg = "Impossibile aprire il file word."
        return (False, error_msg, set())

    # Espressione Regolare (Regex):
    # \{\{    -> Corrisponde a '{{' (le graffe sono caratteri speciali e vanno 'escapate' con \)
    # (.*?)   -> Cattura qualsiasi carattere (punto) zero o più volte (asterisco) in modo non avido (punto interrogativo). 
    #            Questo è il contenuto che vogliamo estrarre.
    # \}\}    -> Corrisponde a '}}'
    pattern = re.compile(r"\{\{(.*?)\}\}")
    
    placeholders = set() # Usiamo un set per evitare duplicati

    # 1. Itera sui paragrafi (testo principale)
    for paragraph in document.paragraphs:
        # Trova tutte le corrispondenze nel testo del paragrafo
        matches = pattern.findall(paragraph.text)
        for match in matches:
            # Pulisci gli spazi bianchi attorno alla parola e aggiungi al set
            placeholders.add(match.strip())

    # 2. Itera sulle tabelle (testo nascosto nelle celle)
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                # Ricerca in tutti i paragrafi all'interno della cella
                for paragraph in cell.paragraphs:
                    matches = pattern.findall(paragraph.text)
                    for match in matches:
                        placeholders.add(match.strip())

    # logging.debug(f"Placeholder values found in template model: {placeholders}")

    # --- 1. Caricamento ---
    try:
        workbook = openpyxl.load_workbook(excel_file,data_only=True)
        df = workbook.active
    except Exception as e:
        logging.warning(f"CHECK_1: cannot open the excel file: {e}")
        error_msg = "Impossibile aprire il file excel."
        return (False, error_msg, placeholders)
    

    
    # Questo dizionario mapperà: {'NomeColonna': indice_colonna_base_0}
    headers_set = set()

    # Itera solo sulla prima riga (min_row=1, max_row=1)
    for row in df.iter_rows(min_row=1, max_row=1):
        for cell in row:
            # cell.column è 1-based (A=1, B=2)
            # Lo convertiamo in 0-based (A=0, B=1) per usarlo con i tuple
            if cell.value: # Ignora eventuali celle vuote nell'intestazione
                headers_set.add(cell.value)

    # logging.debug(f"Excel colums found: {headers_set}")

    if not headers_set.issuperset(placeholders):
        logging.warning(f"CHECK_1: Found some word placeholder not present in excel list: {placeholders.difference(headers_set)}")
        error_msg= f"Alcuni campi del word non sono stati trovati nell'excel: {placeholders.difference(headers_set)}."
        return  (False, error_msg, placeholders)
    
    return (True, "", placeholders)

def convert_cell(mia_cella: Cell) -> str:
    """
    Prende in input una cella di openpyxl e restituisce una
    rappresentazione stringa formattata.

    - Le valute sono formattate come xxx.xxx,xx
    - Le date sono formattate come gg/mm/aaaa
    - Le celle vuote restituiscono una stringa vuota
    """
    
    valore = mia_cella.value
    formato = mia_cella.number_format
    
    # 1. Cella VUOTA
    if valore is None:
        return ""

    # 2. Cella DATA (Requisito: gg/mm/aaaa)
    # openpyxl converte automaticamente le date Excel in oggetti 'datetime'
    if isinstance(valore, datetime):
        return valore.strftime('%d/%m/%Y')

    # 3. Cella NUMERICA (Intero o Float)
    if isinstance(valore, (int, float)):
        
        # 3a. È una VALUTA? (Requisito: xxx.xxx,xx)
        # Controlliamo cercando i simboli di valuta nel formato Excel
        if '€' in formato or '$' in formato or '£' in formato or 'Currency' in formato:
            
            # "%.2f" forza 2 decimali
            # "grouping=True" usa il separatore delle migliaia ('.')
            # Il separatore decimale (',') è automatico grazie al locale.setlocale()
            return locale.format_string("%.2f", valore, grouping=True)
        
        # 3b. È un altro numero (generico o intero)
        else:
            # Per i numeri generici, usiamo la conversione standard di Python
            # (che userà il '.' come separatore decimale, es. 10.5)
            return str(valore)

    # 4. Cella STRINGA (Testo)
    if isinstance(valore, str):
        return valore
        
    # 5. Cella BOOLEANA o altro (es. Errori Excel come #N/D)
    # Per tutto il resto (True, False, ecc.), restituisce la sua
    # rappresentazione stringa standard.
    return str(valore)

def check_integrity(excel_file, placehoders: set[str]):

    excel_file.seek(0)

    # return dict

    dict_to_return = {}
    dict_errors = {}
        
    # 1. Collecting columns name with their index
    # logging.debug(f"Collecting excel column names and their index.")


    workbook = openpyxl.load_workbook(excel_file,data_only=True)
    df = workbook.active

    # Questo dizionario mapperà: {'NomeColonna': indice_colonna_base_0}
    headers_names = {}


    # Itera solo sulla prima riga (min_row=1, max_row=1)
    for row in df.iter_rows(min_row=1, max_row=1):
        for cell in row:
            # cell.column è 1-based (A=1, B=2)
            # Lo convertiamo in 0-based (A=0, B=1) per usarlo con i tuple
            if cell.value and cell.value in placehoders: 
                headers_names[cell.value] = cell.column-1



    # 2. Ciclo su ogni riga del file Excel
    
    colums_types = {}
    for row in df.iter_rows(min_row=2):

        # Se la riga è vuota warning e skippo
        is_row_empty = all(cell.value is None for cell in row)
        if is_row_empty:
            dict_errors[row[0].row] = "Empy row."
            logging.warning(f"Found empty row on line {row[0].row}. Skipped.")
            continue

        # logging.debug(f"Retrieving values for row {row[0].row}")

        try:
            values_of_row = {}
            for placeholder in placehoders:
                if  len(row) <= headers_names[placeholder] or row[headers_names[placeholder]].value is None :
                    #logger.warning(f"No value  found for placeholder {placeholder} on row {row[0].row}. Skipped row.")
                    #break
                    raise Exception(f"No value  found for placeholder {placeholder} on row {row[0].row}.")
                cell_got = row[headers_names[placeholder]]
                if placeholder in colums_types and colums_types[placeholder] != cell_got.data_type:
                    #logger.warning(f"Mismatch data type for {placeholder} in row {row[0].row}: expected {colums_types[placeholder]} - found {cell_got.data_type}. Skipped row.")
                    raise Exception(f"Mismatch data type for {placeholder} in row {row[0].row}: expected {colums_types[placeholder]} - found {cell_got.data_type}.")
                    #break
                elif placeholder  not in colums_types:
                    logging.debug(f"New data type recorded for {placeholder}: {cell_got.data_type}")
                    colums_types[placeholder] = cell_got.data_type
                values_of_row[placeholder] = convert_cell(cell_got)                            
                logging.debug(f"Row {row[0].row}: inserted {placeholder} = {values_of_row[placeholder]}")
        except Exception as e:
            dict_errors[row[0].row] = e
            # logging.error(f"Error found in excel file. {e}")
            continue
        dict_to_return[row[0].row] = values_of_row


    return (dict_to_return, dict_errors)


def check_integrity_paceholder(excel_file):
    """
    Verifica righe 'strane' nell'Excel.
    Return: list
    - list: Lista di warning (es. "Riga 5 saltata"). Vuota se tutto perfetto.
    """
    logging.info("CHECK_2: Inizio controllo integrità Excel")
    
    # TODO: INCOLLA QUI LA TUA LOGICA
    # Ricordati di resettare il puntatore del file se lo hai letto prima!
    excel_file.seek(0) 
    
    warnings = []
    time.sleep(1)
    
    # Esempio simulato:
    # warnings = ["Riga 10: Valore nullo, verrà saltata", "Riga 40: Formato data errato"]
    
    if warnings:
        logging.warning(f"CHECK_2: Trovati {len(warnings)} warning.")
    else:
        logging.info("CHECK_2: Nessun problema rilevato.")
        
    return warnings

# --- STEP 3: GENERAZIONE FINALE ---
def generate_final_zip(excel_file, word_file):
    """
    Esegue l'elaborazione finale e crea lo zip.
    """
    logging.info("FINAL: Generazione output in corso...")
    excel_file.seek(0)
    word_file.seek(0)
    
    time.sleep(1)
    
    # TODO: Tua logica di creazione file
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zf:
        zf.writestr("report.txt", "Elaborazione completata.")
        # Aggiungi qui i tuoi file veri
        
    return zip_buffer.getvalue()