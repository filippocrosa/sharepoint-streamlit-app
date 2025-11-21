# log_utils.py
import logging
import logging.handlers
import streamlit as st
import uuid
import os

# --- 1. L'Adapter (Copia-incolla questo) ---
# Questo serve a iniettare User e Session ID automaticamente
class ContextAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        # Recupera i dati dalla sessione
        user = st.session_state.get('email_user', 'Anonymous')
        sess_id = st.session_state.get('session_id', 'NoSession')
        
        # --- FIX: Dobbiamo mettere i dati dentro la chiave 'extra' ---
        
        # Controlla se ci sono già argomenti 'extra' passati nella chiamata, altrimenti crea un dict vuoto
        if 'extra' not in kwargs:
            kwargs['extra'] = {}
        
        # Inserisce user_id e session_id DENTRO il dizionario 'extra'
        kwargs['extra']['user_id'] = user
        kwargs['extra']['session_id'] = sess_id
        
        # Ritorna il messaggio e i kwargs aggiornati correttemente
        return msg, kwargs

# --- 2. Il Setup (Nascosto e Cachato) ---
# Nota l'underscore _setup: indica che è una funzione "interna"
@st.cache_resource
def _configure_logger():
    logger = logging.getLogger("StreamlitApp")
    logger.setLevel(logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()

    # Rotazione file
    log_file = "webapp_activity.log"
    rotation_handler = logging.handlers.RotatingFileHandler(
        log_file, maxBytes=5*1024*1024, backupCount=5
    )
    
    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s - [%(levelname)s] - [Sid: %(session_id)s] - [User: %(user_id)s] - %(message)s'
    )
    rotation_handler.setFormatter(formatter)
    logger.addHandler(rotation_handler)
    
    # Console (opzionale)
    # console = logging.StreamHandler()
    # console.setFormatter(formatter)
    # logger.addHandler(console)

    return logger

# --- 3. La funzione pubblica che userai ovunque ---
def get_logger():
    """
    Chiama questa funzione in QUALSIASI file per ottenere il logger.
    """
    # Assicura che il logger base sia configurato (lo fa una volta sola grazie alla cache)
    base_logger = _configure_logger()
    
    # Restituisce l'adapter che "legge" la sessione corrente
    return ContextAdapter(base_logger, {})

# --- 4. Inizializzazione Sessione (Utility) ---
def init_logging_session():
    """Da chiamare solo all'inizio di app.py"""
    if 'session_id' not in st.session_state:
        st.session_state['session_id'] = str(uuid.uuid4())[:8]