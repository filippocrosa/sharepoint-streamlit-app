import streamlit as st
import logging
from config.settings import Config
from ui.login_screen import render_login
from ui.dashboard import render_dashboard
from config.localization import setup_app_locale

# 1. Configurazione Pagina
st.set_page_config(page_title=Config.APP_NAME, page_icon="⚡")

# 2. Setup Logging
logging.basicConfig(
    filename=Config.LOG_FILE, 
    level=logging.INFO,
    format='%(asctime)s - %(message)s'
)

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