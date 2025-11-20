import streamlit as st
import logging
# Importiamo le 3 funzioni separate
from core.processor import check_consistency, check_integrity, generate_final_zip


def reset_state():
    """Reset delle variabili se l'utente cambia i file caricati"""
    st.session_state.fase_elaborazione = 0
    st.session_state.opzioni_disponibili = []
    st.session_state.integrity_warnings = []

def render_dashboard():
    user = st.session_state.get("email_user", "Utente")
    
    with st.sidebar:
        st.write(f"👤 **{user}**")
        if st.button("Logout"):
            st.session_state.clear()
            st.rerun()

    st.title("🚀 Generatore Report")
    st.info("Carica i file per avviare i controlli.")

    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("File Excel", type=["xlsx"])
    with col2:
        word_file = st.file_uploader("File Word", type=["docx"])

    if excel_file and word_file:
        st.markdown("---")
        
        # Usiamo un pulsante per far partire tutto
        if st.button("Avvia Controlli ed Elaborazione", type="primary"):
            
            # Creiamo un contenitore di stato (Spinner evoluto)
            status_container = st.status("Esecuzione controlli in corso...", expanded=True)
            
            # --- FASE 1: CHECK COERENZA ---
            status_container.write("🔍 Controllo coerenza Word/Excel...")
            is_consistent, error_msg, placeholder_found = check_consistency(excel_file, word_file)
            
            if not is_consistent:
                # CASO FALLIMENTO STEP 1
                status_container.update(label="❌ Errore Coerenza Rilevato", state="error", expanded=True)
                st.error("I file non sono coerenti. Correggi i seguenti errori e riprova:")
                
                # Mostriamo gli errori in un box rosso espandibile o una tabella
                with st.expander("Dettagli Errori (Clicca per espandere)", expanded=True):
                    st.markdown(f"- 🔴 {error_msg}")
                
                logging.error(f"CHECK_FAIL: {user} bloccato al check coerenza.")
                return # SI FERMA QUI. L'utente deve ricaricare i file.

            # Se passa, scriviamo ok nello status
            status_container.write("✅ Check coerenza superato.")
            
            # --- FASE 2: CHECK INTEGRITÀ ---
            status_container.write("🛠️ Verifica integrità righe Excel...")
            (dict_rows_OK, dict_rows_FAIL) = check_integrity(excel_file, placeholder_found)
            
            if len(dict_rows_FAIL) != 0:
                status_container.write("⚠️ Riscontrati problemi su alcune righe - verranno saltate.")
                # Salviamo i warning per mostrarli dopo, ma NON fermiamo il processo
            else:
                status_container.write("✅ Check integrità perfetto.")

            # --- FASE 3: GENERAZIONE ZIP ---
            status_container.write("📦 Creazione pacchetto finale...")
            zip_data = generate_final_zip(excel_file, word_file)
            
            # Completamento Status
            status_container.update(label="✅ Elaborazione Completata!", state="complete", expanded=False)
            
            # --- OUTCOME FINALE ALL'UTENTE ---
            st.divider()
            st.success("Processo terminato con successo!")

            # Se c'erano warning nel passo 2, li mostriamo qui ben visibili
            if len(dict_rows_FAIL) != 0:
                with st.expander("⚠️ Attenzione: alcune righe sono state ignorate", expanded=False):
                    st.warning("Il file è stato elaborato, ma queste righe sono state saltate:")
                    st.table(dict_rows_FAIL) # O st.write(integrity_warnings)

            # Bottone Download
            st.download_button(
                label="📥 Scarica Risultati (ZIP)",
                data=zip_data,
                file_name="risultati_verificati.zip",
                mime="application/zip"
            )