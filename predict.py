import imaplib
import smtplib
import email
import os
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from transformers import pipeline
import openai
import pdfplumber
from io import BytesIO
import pytesseract
from PIL import Image
import re
import cv2
import numpy as np
from spellchecker import SpellChecker

try:
    from docx import Document
except ImportError:
    print("Il modulo python-docx non è installato. Esegui 'pip install python-docx' per installarlo.")
    raise

# Configurazione delle credenziali di Aruba
IMAP_HOST = 'in.postassl.it'
IMAP_PORT = 993
IMAP_USER = 'test@smartdigitalsolutions.it'
IMAP_PASS = 'GenerativeAI_123'

SMTP_HOST = 'out.postassl.it'
SMTP_PORT = 465
SMTP_USER = 'test@smartdigitalsolutions.it'
SMTP_PASS = 'GenerativeAI_123'

# Imposta il token di Hugging Face
os.environ["HUGGINGFACEHUB_API_TOKEN"] = "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"

# Funzione per connettersi al server IMAP
def connect_to_imap():
    print("Connettendo al server IMAP...")
    mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
    mail.login(IMAP_USER, IMAP_PASS)
    print("Connessione IMAP riuscita")
    return mail

# Funzione per connettersi al server SMTP
def connect_to_smtp():
    print("Connettendo al server SMTP...")
    server = smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT)
    server.login(SMTP_USER, SMTP_PASS)
    print("Connessione SMTP riuscita")
    return server

# Funzione per recuperare le email non lette
def fetch_unread_emails(mail):
    mail.select('inbox')
    status, response = mail.search(None, 'UNSEEN')
    unread_msg_nums = response[0].split()
    print("Email non lette recuperate:", len(unread_msg_nums))
    return unread_msg_nums

# Funzione per ottenere il contenuto delle email
def get_email_content(mail, email_id):
    status, data = mail.fetch(email_id, '(RFC822)')
    msg = email.message_from_bytes(data[0][1])
    return msg

# Funzione per rimuovere i bordi del foglio
def remove_borders(image):
    contours, _ = cv2.findContours(image, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnt = max(contours, key=cv2.contourArea)
    x, y, w, h = cv2.boundingRect(cnt)
    cropped_image = image[y:y+h, x:x+w]
    return cropped_image

# Funzione per estrarre il testo dagli allegati PDF
def extract_text_from_pdf(part):
    text = ""
    try:
        pdf_content = part.get_payload(decode=True)
        with pdfplumber.open(BytesIO(pdf_content)) as pdf:
            for page_num, page in enumerate(pdf.pages):
                page_text = page.extract_text()
                if page_text is None:
                    print(f"Avviso: Impossibile estrarre il testo in modo regolare dalla pagina {page_num}. Provo con OCR.")
                    page_image = page.to_image()
                    pil_image = page_image.original
                    page_text = pytesseract.image_to_string(pil_image, lang='ita+eng+fra+spa+deu')
                text += page_text + "\n"
        print("Testo estratto dal PDF:", text)
    except Exception as e:
        print(f"Errore nell'estrazione del testo dal PDF: {e}")
    return text

# Funzione per preprocessare le immagini
def preprocess_image(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    gray = cv2.bilateralFilter(gray, 11, 17, 17)
    gray = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 11, 2)
    return gray

# Funzione per estrarre il testo dalle immagini
def extract_text_from_image(part):
    text = ""
    try:
        image_content = part.get_payload(decode=True)
        image = Image.open(BytesIO(image_content))
        image_cv = np.array(image)
        preprocessed_image = preprocess_image(image_cv)
        text = pytesseract.image_to_string(preprocessed_image, config='--psm 6', lang='ita+eng+fra+spa+deu')
        print("Testo estratto dall'immagine:", text)
    except Exception as e:
        print(f"Errore nell'estrazione del testo dall'immagine: {e}")
    return text

# Funzione per estrarre il testo dagli allegati DOCX
def extract_text_from_docx(part):
    text = ""
    try:
        docx_content = part.get_payload(decode=True)
        document = Document(BytesIO(docx_content))
        for paragraph in document.paragraphs:
            text += paragraph.text + "\n"
        print("Testo estratto dal DOCX:", text)
    except Exception as e:
        print(f"Errore nell'estrazione del testo dal DOCX: {e}")
    return text

# Funzione per estrarre il testo dall'email e dagli allegati
def extract_text_from_email(msg):
    text = ""
    has_attachments = False
    document_text = ""  # Variabile per salvare il testo estratto dagli allegati
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                text += part.get_payload(decode=True).decode('utf-8')
            elif part.get_content_type() == 'application/pdf':
                document_text += extract_text_from_pdf(part)
                has_attachments = True
            elif part.get_content_type() == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
                document_text += extract_text_from_docx(part)
                has_attachments = True
            elif part.get_content_type().startswith('image/'):
                document_text += extract_text_from_image(part)
                has_attachments = True
    else:
        text = msg.get_payload(decode=True).decode('utf-8')
    return text, has_attachments, document_text

# Caricamento del modello di NLP
nlp_model = pipeline('zero-shot-classification', model='facebook/bart-large-mnli')

# Definizione delle categorie e delle parole chiave
categories_keywords = {
    'Fatture': [
        'fattura', 'pagamento', 'scadenza', 'invoice', 'bolletta', 'ricevuta', 'fee', 'bill', 'amount', 'due', 'balance', 
        'credit', 'debt', 'charge', 'statement', 'compensation', 'money', 'debito', 'credito', 'tariffa', 'importo', 
        'somma', 'addebitare', 'estratto conto', 'compenso', 'denaro', 'rata', 'contabile', 'fatturazione'
    ],
    'Progetti': [
        'progetto', 'collaborazione', 'nuovo', 'project', 'task', 'lavoro', 'assignment', 'development', 'team', 'milestone', 
        'objective', 'goal', 'plan', 'schedule', 'deadline', 'deliverable', 'cooperazione', 'squadra', 'obiettivo', 
        'piano', 'programma', 'scadenza', 'consegna', 'sviluppo', 'attività', 'progettazione', 'progettare', 'fase', 'implementazione',
        'sito web', 'web development', 'crm', 'manutenzione', 'website', 'web design', 'digital marketing', 'SEO', 
        'SEM', 'content marketing', 'tecnologia', 'IT', 'e-commerce', 'gestione clienti', 'database', 'applicazioni', 
        'software', 'user experience', 'UX', 'UI', 'analisi dati', 'cloud', 'cybersecurity', 'automazione'
    ],
    'Supporto': [
        'supporto', 'aiuto', 'problema', 'assistenza', 'help', 'issue', 'troubleshooting', 'error', 'bug', 'fix', 
        'support', 'customer', 'service', 'repair', 'maintenance', 'inquiry', 'question', 'sostegno', 'errore', 
        'assistenza clienti', 'riparazione', 'manutenzione', 'domanda', 'richiesta', 'informazioni', 'risoluzione', 
        'guasto', 'difetto', 'malfunzionamento', 'chiedere', 'risolvere', 'intervento', 'tecnico', 'consultazione',
        'ticket', 'chiamata', 'chat', 'email support', 'guida', 'istruzioni', 'docente', 'formazione', 'lezione', 
        'esercitazione', 'supporto tecnico', 'supporto progetti', 'helpdesk', 'manutenzione siti'
    ],
    'Altro': [
        'news', 'update', 'announcement', 'event', 'meeting', 'notification', 'reminder', 'social', 'media', 'network', 
        'personal', 'miscellaneous', 'various', 'general', 'misc', 'other', 'notizia', 'aggiornamento', 'annuncio', 
        'evento', 'riunione', 'notifica', 'promemoria', 'sociale', 'rete', 'personale', 'vario', 'generale', 'altro', 
        'comunicazione', 'invito', 'promozione', 'offerta', 'sconto', 'informazione', 'informazioni', 'avviso', 'calendario',
        'presentazione', 'seminario', 'webinar', 'corso', 'festa', 'celebrazione'
    ]
}

# Funzione per analizzare l'email e determinarne la categoria
def analyze_email(content):
    # Classificazione basata su parole chiave
    keyword_counts = {category: sum(content.lower().count(keyword.lower()) for keyword in keywords) for category, keywords in categories_keywords.items()}
    max_keyword_category = max(keyword_counts, key=keyword_counts.get)

    # Classificazione zero-shot
    categories = list(categories_keywords.keys())
    result = nlp_model(content, candidate_labels=categories)
    zero_shot_category = result['labels'][0]  # Ritorna l'etichetta con la maggiore confidenza

    # Combinare i risultati
    if keyword_counts[max_keyword_category] > 0 and max_keyword_category != zero_shot_category:
        if keyword_counts[max_keyword_category] >= 2 * keyword_counts.get(zero_shot_category, 0):
            return max_keyword_category

    # Fallback alla categoria "Altro" se nessuna categoria specifica è rilevante
    if zero_shot_category in ['Fatture', 'Progetti', 'Supporto'] and keyword_counts[zero_shot_category] == 0:
        return 'Altro'
    
    return zero_shot_category

# Funzione per generare la risposta automatica
openai.api_key = 'sk-emailaccount-PCm6USEfOVb94r9q01PAT3BlbkFJTqbgA4wpIObHSdA8wXti'

def generate_response(role, prompt, has_attachments, max_tokens=600):
    if has_attachments:
        # Se ci sono allegati, non generare risposta automatica
        reply = "Abbiamo ricevuto i documenti allegati e stiamo processando i dati."
    else:
        # Genera risposta automatica solo se non ci sono allegati
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": role},
                {"role": "user", "content": prompt}
            ],
            max_tokens=max_tokens
        )
        reply = response['choices'][0]['message']['content'].strip()

        # Rimuovi eventuali saluti precedenti
        saluti_da_rimuovere = ["Cordiali saluti,", "Il tuo assistente email", "Cordiali saluti,\nil vostro assistente AI di SmartDigitalSolutions"]
        for saluto in reply:
            if saluto in reply:
                reply = reply.replace(saluto, "")

        reply += "\n\nCordiali saluti,\nil vostro assistente AI di SmartDigitalSolutions"

    return reply

# Funzione per inviare l'email e salvare una copia nella cartella Inviati
def send_email(smtp_server, to_address, subject, body, document_text):
    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = to_address
    msg['Subject'] = subject

    # Invia l'email all'utente con i dati estratti
    if document_text:
        body += f"\n\nTesto estratto dai documenti allegati:\n{document_text}"
    
    msg.attach(MIMEText(body, 'plain'))
    smtp_server.sendmail(SMTP_USER, to_address, msg.as_string())
    print("Email inviata a:", to_address)

    # Salva una copia nella cartella Inviati
    save_to_sent_folder(msg)
    print("Email salvata in 'INBOX.Sent'")

# Funzione per salvare una copia dell'email nella cartella Inviati
def save_to_sent_folder(msg):
    try:
        imap_server = connect_to_imap()
        sent_folder = "INBOX.Sent"  # Usa il nome corretto della tua cartella "Inviati"
        result = imap_server.append(sent_folder, "", imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        if result[0] == 'OK':
            print("Email aggiunta a 'INBOX.Sent'")
        else:
            print("Errore nel salvataggio della email in 'INBOX.Sent':", result)
        imap_server.logout()
    except Exception as e:
        print("Errore nel salvataggio della email in 'INBOX.Sent':", e)

# Funzione per spostare l'email in una cartella specifica
def move_email_to_folder(mail, email_id, folder_name):
    prefixed_folder_name = f"INBOX.{folder_name}"
    mail.select('inbox')
    result = mail.copy(email_id, prefixed_folder_name)
    if result[0] == 'OK':
        mail.store(email_id, '+FLAGS', '\\Deleted')
        mail.expunge()
        print(f"Email spostata in '{folder_name}'")
    else:
        print(f"Error copying email to folder '{folder_name}': {result}")

# Funzione principale modificata per eseguire un ciclo continuo
def main():
    while True:
        try:
            imap_conn = connect_to_imap()
            smtp_conn = connect_to_smtp()
            print("Connessione ai server IMAP e SMTP riuscita")

            unread_emails = fetch_unread_emails(imap_conn)
            print("Email non lette recuperate:", len(unread_emails))

            for email_id in unread_emails:
                msg = get_email_content(imap_conn, email_id)
                email_from = email.utils.parseaddr(msg['From'])[1]
                email_subject = msg['Subject']
                email_body, has_attachments, document_text = extract_text_from_email(msg)

                if email_body:  # Assicurati che ci sia del testo nell'email
                    print(f"Elaborazione email da: {email_from} con oggetto: {email_subject}")

                    # Analisi e risposta
                    category = analyze_email(email_body)
                    role = f"Sei un assistente email, rispondi alle domande degli utenti con cortesia e precisione. Categoria: {category}"
                    response_body = generate_response(role, f"Richiesta: {email_body}\nRisposta: {category}", has_attachments)

                    # Invio della risposta
                    send_email(smtp_conn, email_from, email_subject, response_body, document_text)

                    # Sposta l'email nella cartella appropriata
                    move_email_to_folder(imap_conn, email_id, category)

            imap_conn.logout()
            smtp_conn.quit()
            print("Connessioni IMAP e SMTP chiuse")

            # Aspetta per un determinato periodo di tempo prima di ricontrollare (ad esempio, 60 secondi)
            time.sleep(60)
        except Exception as e:
            print(f"Errore: {e}")
            time.sleep(60)  # Attendere prima di riprovare in caso di errore

if __name__ == "__main__":
    main()
