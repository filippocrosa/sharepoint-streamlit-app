import streamlit as st
from config.settings import Config
from ui.login_screen import render_login
from ui.dashboard import render_dashboard
from config.localization import setup_app_locale
from config.log_utils import get_logger, init_logging_session

# 1. Configurazione Pagina
st.set_page_config(page_title=Config.APP_NAME, page_icon="⚡")

# Inizializza l'ID sessione (Fallo all'inizio!)
init_logging_session()

# Ottieni il logger
logging = get_logger()

# 3. Setup locale
setup_app_locale()

# 4. Inizializzazione Stato Globale
if "is_logged_in" not in st.session_state:
    st.session_state.is_logged_in = False

# 5. Router (Decisione su cosa mostrare)
if not st.session_state.is_logged_in:
    render_login()
else:
    render_dashboard()