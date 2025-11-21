import os

# In produzione, useremo os.environ.get("SEGRETO")
# Per ora usiamo valori statici per lo sviluppo

class Config:
    APP_NAME = "Portale creazione lettere"
    MAX_LOGIN_ATTEMPTS = 3
    
    # Lista utenti abilitati
    ALLOWED_USERS = [
        "cliente@azienda.com",
        "admin@tuaazienda.com",
        "test@test.com",
        "bellina.t@confcooperative.it"
    ]

    # Configurazione Log
    LOG_FILE = "security.log"