import locale
import logging
import platform
import streamlit as st

@st.cache_resource
def setup_app_locale():
    """
    Configura il locale a livello di processo all'avvio.
    Tenta diverse configurazioni per Linux (Docker/Cloud) e Windows.
    """
    # Definiamo i target in ordine di priorità
    targets = []
    
    if platform.system() == 'Windows':
        targets = ['it_IT', 'Italian_Italy.1252']
    else:
        # Linux / Mac / Container
        targets = ['it_IT.UTF-8', 'it_IT.utf8']

    # Tentiamo di impostare il locale
    for loc in targets:
        try:
            locale.setlocale(locale.LC_ALL, loc)
            logging.info(f"LOCALE: set to '{loc}'")
            return # Successo! Esci dalla funzione
        except locale.Error:
            logging.warning(f"LOCALE: '{loc}' not available on this system.")

    # Se siamo qui, nessun locale italiano ha funzionato
    # Fallback al default di sistema (spesso Inglese nei server cloud)
    try:
        locale.setlocale(locale.LC_ALL, '')
        logging.warning("LOCALE: default Fallback. Pay attention on curenncy values.")
    except:
        logging.error("LOCALE: impossible to set locale.")