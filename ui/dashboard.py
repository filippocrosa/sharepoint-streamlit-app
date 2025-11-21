import streamlit as st
import time
# Importiamo le 3 funzioni separate
from core.processor import check_consistency, check_integrity, generate_final_zip

from config.log_utils import get_logger
logging = get_logger()

def reset_state():
    """
    Questa funzione parte AUTOMATICAMENTE quando l'utente:
    1. Carica un file nuovo.
    2. Clicca la X per rimuovere un file.
    """
    # Resetta lo stato di avanzamento
    st.session_state.fase_elaborazione = 0
    
    # Pulisce le liste dati
    st.session_state.fase_elaborazione = 0
    st.session_state.placeholders = set()
    st.session_state.integrity_warnings = {}
    st.session_state.excel_rows = {}
    
    # FONDAMENTALE: Cancella il risultato finale se c'era
    if "zip_risultato" in st.session_state:
        del st.session_state.zip_risultato
    if "nome_zip_risultato" in st.session_state:
        del st.session_state.nome_zip_risultato
        
    # Opzionale: Logga l'evento per debug
    logging.info("STATE_RESET: L'utente ha modificato i file di input.")

def reset_state():
    """Reset delle variabili se l'utente cambia i file caricati"""
    st.session_state.fase_elaborazione = 0
    st.session_state.placeholders = set()
    st.session_state.integrity_warnings = {}
    st.session_state.excel_rows = {}
    if "zip_risultato" in st.session_state:
        del st.session_state.zip_risultato
    if "nome_zip_risultato" in st.session_state:
        del st.session_state.nome_zip_risultato

def render_dashboard():
    user = st.session_state.get("email_user", "Utente")

    # --- INIZIALIZZAZIONE STATO ---
    if "fase_elaborazione" not in st.session_state:
        st.session_state.fase_elaborazione = 0 
    if "placeholders" not in st.session_state:
        st.session_state.placeholders = set()
    if "integrity_warnings" not in st.session_state:
        st.session_state.integrity_warnings = {}
    if "excel_rows" not in st.session_state:
        st.session_state.excel_rows = {}
    
    with st.sidebar:
        st.write(f"👤 **{user}**")
        if st.button("Logout"):
            logging.info("LOGOUT")
            st.session_state.clear()
            st.rerun()

    st.title("📄 Generatore Lettere")
    
    # ============================================================
    # 1️⃣ SEZIONE UPLOAD
    # ============================================================
    with st.container(border=True):
        st.markdown("### 📂 1. Caricamento Dati")
        st.info("Carica il template Word e il file Excel con i dati dei clienti.")
        
        col1, col2 = st.columns(2)
        with col1:
            excel_file = st.file_uploader("File Excel", type=["xlsx"], on_change=reset_state)
        with col2:
            word_file = st.file_uploader("File Word", type=["docx"], on_change=reset_state)

    # La logica prosegue solo se i file ci sono
    if excel_file and word_file:
        
        st.write("") # Spazio verticale

        # ============================================================
        # 2️⃣ SEZIONE CONTROLLI (Sempre visibile, cambia contenuto)
        # ============================================================
        with st.container(border=True):
            st.markdown("### 🔍 2. Analisi e Controlli")
            
            # --- CASO A: Dobbiamo ancora lanciare i controlli ---
            if st.session_state.fase_elaborazione == 0:
                st.write("Verranno verificati la coerenza tra i placeholder e l'integrità delle righe Excel.")
                
                if st.button("Avvia Controlli ed Elaborazione", type="primary", use_container_width=True):
                    
                    status_container = st.status("Esecuzione pipeline di controllo...", expanded=True)
                    
                    # STEP 1: Coerenza
                    status_container.write("🔍 Controllo coerenza Word/Excel...")
                    time.sleep(1)
                    is_consistent, error_msg, placeholder_found = check_consistency(excel_file, word_file)
                    
                    if not is_consistent:
                        status_container.update(label="❌ Errore Coerenza Rilevato", state="error", expanded=True)
                        st.error("I file non sono coerenti. Correggi i seguenti errori:")
                        st.write(error_msg)
                        logging.error(f"CHECK_FAIL: {user} bloccato al check coerenza.")
                        return # Stop qui
                        
                    status_container.write("✅ Check coerenza superato.")
                    
                    # STEP 2: Integrità
                    status_container.write("🛠️ Verifica integrità righe Excel...")
                    time.sleep(1)
                    (dict_rows_OK, dict_rows_FAIL) = check_integrity(excel_file, placeholder_found)
                    
                    # Salviamo risultati nello stato
                    st.session_state.placeholders = placeholder_found
                    st.session_state.integrity_warnings = dict_rows_FAIL
                    st.session_state.excel_rows = dict_rows_OK

                    status_container.update(label="✅ Controlli Superati!", state="complete", expanded=True)
                    time.sleep(1)
                    
                    # Avanzamento fase e refresh
                    st.session_state.fase_elaborazione = 1
                    st.rerun()

            # --- CASO B: Controlli già fatti (Mostriamo il Report) ---
            else: # fase_elaborazione >= 1
                
                # Qui decidiamo cosa mostrare come "Report Storico" della fase 2
                if len(st.session_state.integrity_warnings) > 0:
                    st.warning(f"✅ Controlli completati con {len(st.session_state.integrity_warnings)} segnalazioni.")
                    
                    # La tabella dei warning vive QUI, nella Fase 2
                    with st.expander("⚠️ Visualizza righe scartate", expanded=True):
                        st.table(st.session_state.integrity_warnings)
                else:
                    st.success("✅ Tutti i controlli sono stati superati senza errori.")

        # ============================================================
        # 3️⃣ SEZIONE CONFIGURAZIONE (Visibile solo dopo i controlli)
        # ============================================================
        if st.session_state.fase_elaborazione >= 1:
            
            st.write("") # Spazio

            with st.container(border=True):
                st.markdown("### ⚙️ 3. Configurazione Output")
                st.info("Seleziona come nominare i file e genera il pacchetto.")

                col_sel, col_btn = st.columns([3, 1])
                
                with col_sel:
                    scelta = st.selectbox(
                        "Scegli quale colonna usare per nominare i file:", 
                        st.session_state.placeholders
                    )
                
                with col_btn:
                    st.write("") 
                    st.write("")
                    if st.button("⬅️ Indietro", key="back_btn"):
                        reset_state()
                        st.rerun()

                st.markdown("---")
                
                # Pulsante Generazione
                if st.button("🚀 Genera Output Finale", type="primary", use_container_width=True):
                    
                    # CSS Blocco
                    st.markdown("""
                        <style>
                            .stApp { pointer-events: none; cursor: progress; }
                            button { cursor: progress !important; }
                        </style>
                    """, unsafe_allow_html=True)

                    with st.spinner("⏳ Creazione documenti in corso..."):
                        try:
                            zip_bytes = generate_final_zip(
                                word_file=word_file,
                                campo_nome_file=scelta,
                                righe_excel=st.session_state.excel_rows, 
                                placeholders_set=st.session_state.placeholders
                            )                        
                            st.session_state.zip_risultato = zip_bytes
                            st.session_state.nome_zip_risultato = f"Lettere_{scelta}_{int(time.time())}.zip"
                            
                            st.session_state.show_toast = True
                            time.sleep(1)
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"Errore generazione: {e}")
                            logging.error(f"GEN_ERROR: {e}")
                            time.sleep(3)
                            st.rerun()

        # ============================================================
        # 4️⃣ SEZIONE DOWNLOAD
        # ============================================================
        if "zip_risultato" in st.session_state:
            
            if st.session_state.get("show_toast"):
                st.toast('Generazione completata!', icon='✅')
                st.session_state.show_toast = False

            st.write("")
            with st.container(border=True):
                st.markdown("### 📥 4. Download")
                st.success("I file sono pronti.")
                
                col_dl_1, col_dl_2, col_dl_3 = st.columns([1, 2, 1])
                with col_dl_2:
                    st.download_button(
                        label="SCARICA ZIP COMPLETO",
                        data=st.session_state.zip_risultato,
                        file_name=st.session_state.nome_zip_risultato,
                        mime="application/zip",
                        type="primary",
                        use_container_width=True
                    )

def render_dashboard_2():
    user = st.session_state.get("email_user", "Utente")

    # --- INIZIALIZZAZIONE STATO LOCALE ---
    if "fase_elaborazione" not in st.session_state:
        st.session_state.fase_elaborazione = 0 # 0=Start, 1=Select, 2=Done
    if "placeholders" not in st.session_state:
        st.session_state.placeholders =set()
    if "integrity_warnings" not in st.session_state:
        st.session_state.integrity_warnings = {}
    if "excel_rows" not in st.session_state:
        st.session_state.excel_rows = {}
    
    with st.sidebar:
        st.write(f"👤 **{user}**")
        if st.button("Logout"):
            logging.info("LOGOUT")
            st.session_state.clear()
            st.rerun()

    st.title("GENERATORE LETTERE")
    st.info("Carica i file per avviare i controlli.")

    col1, col2 = st.columns(2)
    with col1:
        excel_file = st.file_uploader("File Excel", type=["xlsx"], on_change=reset_state)
    with col2:
        word_file = st.file_uploader("File Word", type=["docx"], on_change=reset_state)

    if excel_file and word_file:
        st.markdown("---")
        
        if st.session_state.fase_elaborazione == 0:
            # Usiamo un pulsante per far partire tutto
            if st.button("Avvia Controlli ed Elaborazione", type="primary"):
                
                # Creiamo un contenitore di stato (Spinner evoluto)
                status_container = st.status("Esecuzione pipeline di controllo...", expanded=True)
                
                # --- FASE 1: CHECK COERENZA ---
                status_container.write("🔍 Controllo coerenza Word/Excel...")
                time.sleep(2)
                is_consistent, error_msg, placeholder_found = check_consistency(excel_file, word_file)
                
                if not is_consistent:
                    # CASO FALLIMENTO STEP 1
                    status_container.update(label="❌ Errore Coerenza Rilevato", state="error", expanded=True)
                    st.error("I file non sono coerenti. Correggi i seguenti errori e riprova:")
                    st.write(error_msg)
                    logging.error(f"CHECK_FAIL: {user} bloccato al check coerenza.")
                    return
                    
                # Se passa, scriviamo ok nello status
                status_container.write("✅ Check coerenza superato.")
                
                # --- FASE 2: CHECK INTEGRITÀ ---
                status_container.write("🛠️ Verifica integrità righe Excel...")
                time.sleep(2)
                (dict_rows_OK, dict_rows_FAIL) = check_integrity(excel_file, placeholder_found)
                
                if len(dict_rows_FAIL) != 0:
                    status_container.write("⚠️ Riscontrati problemi su alcune righe - verranno saltate.")
                    # Salviamo i warning per mostrarli dopo, ma NON fermiamo il processo
                else:
                    status_container.write("✅ Check integrità perfetto.")

                st.session_state.placeholders = placeholder_found
                st.session_state.integrity_warnings = dict_rows_FAIL
                st.session_state.excel_rows = dict_rows_OK

                status_container.update(label="✅ Controlli Superati!", state="complete", expanded=True)
                time.sleep(2)
                # AVANZAMENTO DI STATO
                st.session_state.fase_elaborazione = 1
                st.rerun() # Ricarica la pagina per mostrare la nuova interfaccia
        
        if st.session_state.fase_elaborazione >= 1:
            
            # Disegniamo un box riassuntivo verde (o giallo se c'erano warning)
            if len(st.session_state.integrity_warnings) != 0:
                st.warning(f"✅ Controlli superati (con {len(st.session_state.integrity_warnings)} warning).")
            else:
                st.success("✅ Tutti i controlli preliminari sono stati superati.")

        if st.session_state.fase_elaborazione >= 1:

            # Se c'erano warning nel passo 2, li mostriamo qui ben visibili
            if len(st.session_state.integrity_warnings) != 0:
                with st.expander("⚠️ Attenzione: alcune righe sono state ignorate", expanded=False):
                    st.table(st.session_state.integrity_warnings) # O st.write(integrity_warnings)

            st.markdown("---")
            st.subheader("Configurazione Output")
            
            col_sel, col_btn = st.columns([3, 1])
            
            with col_sel:
                scelta = st.selectbox(
                    "Seleziona il campo da utilizzare per nominare i files:", 
                    st.session_state.placeholders
                )
            
            with col_btn:
                # Spaziatura per allineare il bottone alla selectbox
                st.write("") 
                st.write("")
                if st.button("Indietro", key="back_btn"):
                    reset_state()
                    st.rerun()


            if st.button("Genera Output Finale", type="primary", width="content"):
                # print(st.session_state.excel_rows)

                st.markdown("""
                    <style>
                        /* Blocca l'interazione su tutta la pagina */
                        .stApp {
                            pointer-events: none; /* Nessun click passa */
                            cursor: progress;     /* Il mouse diventa una clessidra o rotellina */
                        }
                        
                        /* Opzionale: Se vuoi che i bottoni sembrino "congelati" 
                           senza sfuocare tutto, puoi togliere l'effetto hover */
                        button {
                            cursor: progress !important;
                        }
                    </style>
                """, unsafe_allow_html=True)

                # 1. Mostriamo all'utente che stiamo lavorando (Word ci mette un po')
                with st.spinner("⏳ Apertura Word, conversione in PDF e creazione ZIP... (Attendi qualche secondo)"):
                    
                    # 2. Chiamata al Backend
                    try:

                        zip_bytes =  generate_final_zip(word_file=word_file  ,campo_nome_file=scelta,righe_excel=st.session_state.excel_rows,  placeholders_set=st.session_state.placeholders)                        
                        # 3. SALVIAMO IL RISULTATO NELLO STATO
                        # Così non lo perdiamo se la pagina si ricarica
                        st.session_state.zip_risultato = zip_bytes
                        st.session_state.nome_zip_risultato = f"Lettere_{scelta}_{int(time.time())}.zip"
                        
                        st.success("✅ Elaborazione Completata!")
                        st.toast('Generazione completata con successo!', icon='✅')
                        time.sleep(2)
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"Errore durante la generazione: {e}")
                        logging.error(f"GEN_ERROR: {e}")
                        time.sleep(3) # Diamo tempo di leggere l'errore
                        st.rerun() # Ricarichiamo per sbloccare l'interfaccia



            if "zip_risultato" in st.session_state:

                st.markdown("---")
                st.info("Files pronti per il download.")
                st.download_button(
                    label="📥 SCARICA ZIP COMPLETO",
                    data=st.session_state.zip_risultato,
                    file_name=st.session_state.nome_zip_risultato,
                    mime="application/zip",
                    type="primary",
                    use_container_width=True
                )

    # else:
    #     reset_state()
    #     st.rerun()