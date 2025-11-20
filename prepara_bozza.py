import mimetypes
from email.message import EmailMessage
import os

def crea_bozza_universale(destinatario, oggetto, corpo, percorso_allegato, percorso_output):
    # Creiamo l'oggetto mail standard
    msg = EmailMessage()
    msg['Subject'] = oggetto
    msg['To'] = destinatario
    
    # --- IL TRUCCO FONDAMENTALE ---
    # Questo header dice a Outlook (e Thunderbird, Mail.app, ecc.) 
    # che questa mail è una BOZZA non ancora inviata.
    # Senza questo, la mail si aprirebbe in "sola lettura".
    msg['X-Unsent'] = '1'
    # ------------------------------

    # Impostiamo il corpo del testo
    msg.set_content(corpo)

    # Gestione Allegato
    if percorso_allegato:
        # Indovina il tipo di file (es. application/pdf, image/jpeg)
        ctype, encoding = mimetypes.guess_type(percorso_allegato)
        if ctype is None or encoding is not None:
            # Se non lo riconosce, usiamo un tipo generico binario
            ctype = 'application/octet-stream'
        
        maintype, subtype = ctype.split('/', 1)

        # Legge il file in modalità binaria e lo allega
        with open(percorso_allegato, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(percorso_allegato)
            
            msg.add_attachment(file_data,
                               maintype=maintype,
                               subtype=subtype,
                               filename=file_name)

    # Salvataggio su file .eml
    with open(percorso_output, 'wb') as out_file:
        out_file.write(msg.as_bytes())
    
    print(f"File generato: {percorso_output}")

# --- ESEMPIO DI UTILIZZO ---

destinatario = "mario.rossi@azienda.com"
oggetto = "Preventivo Server"
testo = "Ciao Mario,\n\necco il preventivo di cui parlavamo.\n\nFammi sapere."
file = "C:\\Users\\crosa.f\\progetti\\sharepoint-streamlit-app\\VASCHETTO ANGELA.pdf"  # Assicurati che esista
output = "Bozza_per_Collega.eml" # Estensione .eml è importante

# Esegui (funziona su Windows, Linux, Mac)
crea_bozza_universale(destinatario, oggetto, testo, file, output)