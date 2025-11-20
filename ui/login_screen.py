import streamlit as st
import logging  # <--- Mancava questo import!
from core.auth import generate_otp, send_otp_email
from config.settings import Config
import time

def render_login():
    st.subheader("🔐 Accesso Riservato")

    # Inizializza variabili se non esistono
    if "otp_sent" not in st.session_state:
        st.session_state.otp_sent = False

    if not st.session_state.otp_sent:
        # --- STEP 1: Inserimento Email ---
        email_input = st.text_input("Email Aziendale")
        
        if st.button("Ricevi Codice"):
            email_clean = email_input.strip().lower()
            
            # 1. Generiamo l'OTP UNA SOLA VOLTA qui
            nuovo_otp = generate_otp()
            
            # 2. Proviamo a inviarlo
            invio_riuscito = send_otp_email(email_clean, nuovo_otp)
            
            if invio_riuscito:
                # 3. Se l'invio va a buon fine, salviamo QUELLO STESSO codice
                st.session_state.otp_secret = nuovo_otp
                st.session_state.email_user = email_clean
                st.session_state.otp_sent = True
                st.session_state.attempts = 0
                
                # Logghiamo il successo dell'invio
                logging.info(f"LOGIN_FLOW: Codice OTP inviato a {email_clean}")
                st.rerun()
            else:
                st.error("Email non abilitata o errore nell'invio.")
                # Il log di errore è già gestito dentro send_otp_email in core/auth.py
    else:
        # --- STEP 2: Inserimento OTP ---
        st.info(f"Codice inviato a {st.session_state.email_user}. Controlla la console di VS Code.")
        otp_input = st.text_input("Codice OTP", max_chars=6)
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("Verifica"):
                if otp_input == st.session_state.otp_secret:
                    # SUCCESSO
                    st.session_state.is_logged_in = True
                    logging.info(f"LOGIN_SUCCESS: L'utente {st.session_state.email_user} è entrato nel sistema.")
                    st.success("Login effettuato!")
                    time.sleep(0.5)
                    st.rerun()
                else:
                    # ERRORE
                    st.session_state.attempts += 1
                    remaining = Config.MAX_LOGIN_ATTEMPTS - st.session_state.attempts
                    
                    logging.warning(f"LOGIN_FAIL: OTP errato per {st.session_state.email_user}. Tentativo {st.session_state.attempts}/{Config.MAX_LOGIN_ATTEMPTS}")
                    
                    if remaining <= 0:
                        logging.error(f"SECURITY_BLOCK: Utente {st.session_state.email_user} bloccato per troppi tentativi.")
                        st.error("⛔ Troppi tentativi. Sessione bloccata per sicurezza.")
                        st.stop()
                    else:
                        st.error(f"Codice errato. Tentativi rimasti: {remaining}")

        with col2:
            if st.button("Indietro / Cambia Email"):
                st.session_state.otp_sent = False
                st.session_state.attempts = 0
                st.rerun()