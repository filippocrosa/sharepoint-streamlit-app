import random
import logging
from config.settings import Config

def generate_otp() -> str:
    """Genera un codice a 6 cifre."""
    return str(random.randint(100000, 999999))

def send_otp_email(email: str, otp: str) -> bool:
    """
    Gestisce l'invio del codice.
    In dev: Stampa su console.
    In prod: Invia SMTP.
    """
    # Logica di sicurezza
    if email not in Config.ALLOWED_USERS:
        logging.warning(f"AUTH: Tentativo di login non autorizzato per {email}")
        return False

    # Simulazione Invio
    logging.info(f"AUTH: Invio OTP a {email}")
    print(f"\n--- 📧 EMAIL SIMULATA PER {email} ---\nCodice: {otp}\n-----------------------------------\n")
    return True

def verify_otp(input_otp: str, real_otp: str) -> bool:
    """Verifica se il codice combacia."""
    return input_otp == real_otp