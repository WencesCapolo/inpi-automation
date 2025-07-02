import streamlit as st
import pandas as pd
import json
import os
import sys
import time
import traceback
import logging
from datetime import datetime
import requests
import random
import re
from PyPDF2 import PdfReader
import openai

# Add current directory to path to import your modules
sys.path.append('.')

# Import your existing functions
try:
    from process_actas import get_session_with_cookies, find_formulario_item, construct_document_url, download_pdf_with_retry, extract_email_from_pdf
except ImportError as e:
    st.error(f"Could not import process_actas functions: {e}")

# Page configuration    
st.set_page_config(
    page_title="Estudio Eguía",
    page_icon="⚖️",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f4e79;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    .section-divider {
        border-top: 3px solid #1f4e79;
        margin: 2rem 0;
    }
    .log-container {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 5px;
        border-left: 4px solid #17a2b8;
        font-family: 'Courier New', monospace;
        font-size: 0.9rem;
        max-height: 400px;
        overflow-y: auto;
    }
    .error-log {
        background-color: #f8d7da;
        color: #721c24;
        border-left-color: #dc3545;
    }
    .success-log {
        background-color: #d4edda;
        color: #155724;
        border-left-color: #28a745;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'step' not in st.session_state:
    st.session_state.step = 1
if 'uploaded_data' not in st.session_state:
    st.session_state.uploaded_data = None
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'logs' not in st.session_state:
    st.session_state.logs = []

# Setup logging system
def setup_logger():
    """Setup minimal logging system with console output"""
    logger = logging.getLogger('inpi_automation')
    
    if not logger.handlers:  # Avoid duplicate handlers
        logger.setLevel(logging.DEBUG)
        
        # Console handler
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        
        # Simple formatter for traceability
        formatter = logging.Formatter(
            '%(asctime)s | %(levelname)s | %(funcName)s:%(lineno)d | %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    
    return logger

# Initialize logger
logger = setup_logger()

def add_log(message, log_type="info", include_traceback=False):
    """Enhanced logging with both console and Streamlit UI output"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    
    # Get caller information for better traceability
    frame = sys._getframe(1)
    caller_name = frame.f_code.co_name
    caller_line = frame.f_lineno
    
    # Create enhanced message with context
    if caller_name != '<module>':
        enhanced_message = f"{caller_name}:{caller_line} | {message}"
    else:
        enhanced_message = message
    
    # Add to Streamlit session state (preserve existing functionality)
    st.session_state.logs.append({
        "timestamp": timestamp,
        "message": message,  # Keep original message for UI
        "type": log_type,
        "function": caller_name,
        "line": caller_line
    })
    
    # Log to console with appropriate level
    if log_type == "error":
        logger.error(enhanced_message)
        if include_traceback:
            logger.error("Traceback: %s", traceback.format_exc())
    elif log_type == "warning":
        logger.warning(enhanced_message)
    elif log_type == "success":
        logger.info(f"✅ {enhanced_message}")
    else:  # info and others
        logger.info(enhanced_message)

def log_api_error(operation, url, status_code=None, error_msg=None, response_text=None):
    """Specialized logging for API failures"""
    base_msg = f"API {operation} failed - {url}"
    
    if status_code:
        base_msg += f" (Status: {status_code})"
    
    if error_msg:
        base_msg += f" - {error_msg}"
    
    add_log(base_msg, "error")
    
    if response_text:
        add_log(f"Response details: {response_text[:200]}...", "error")

def log_file_error(operation, filename, error):
    """Specialized logging for file processing errors"""
    add_log(f"File {operation} failed - {filename}: {str(error)}", "error", include_traceback=True)

def log_auth_error(service, status_code, details=None):
    """Specialized logging for authentication/authorization errors"""
    msg = f"Authentication error - {service} (Status: {status_code})"
    if details:
        msg += f" - {details}"
    add_log(msg, "error")

def display_logs():
    """Display logs in a container"""
    if st.session_state.logs:
        log_text = ""
        for log in st.session_state.logs[-50:]:  # Show last 50 logs
            log_text += f"[{log['timestamp']}] {log['message']}\n"
        
        log_class = ""
        if any(log['type'] == 'error' for log in st.session_state.logs[-10:]):
            log_class = "error-log"
        elif any(log['type'] == 'success' for log in st.session_state.logs[-5:]):
            log_class = "success-log"
        
        st.markdown(f'<div class="log-container {log_class}"><pre>{log_text}</pre></div>', unsafe_allow_html=True)

def process_sheet(df, sheet_name):
    """Process a single sheet and extract rows where first column is 'Part.'"""
    try:
        add_log(f"Processing sheet: {sheet_name}")
        
        # Find the row index where 'Agente' appears
        agente_idx = df.iloc[:, 0][df.iloc[:, 0] == 'Agente'].index[0]
        add_log(f"Found 'Agente' header at row {agente_idx}")
        
        # Use the row after 'Agente' as header
        df.columns = df.iloc[agente_idx, :]
        
        # Get data after the header row
        data_df = df.iloc[agente_idx + 1:].copy()
        data_df.columns = ['Agente' if pd.isna(col) else str(col).strip() for col in data_df.columns]
        
        # Filter rows where Agente column equals 'Part.'
        part_rows = data_df[data_df['Agente'] == 'Part.'].copy()
        add_log(f"Found {len(part_rows)} rows with Agente = 'Part.'")
        
        # Convert to dictionary format
        rows_data = []
        for idx, row in part_rows.iterrows():
            row_dict = {}
            for col in data_df.columns:
                val = row[col]
                if pd.isna(val):
                    row_dict[col] = None
                else:
                    row_dict[col] = str(val).strip()
            row_dict['origen'] = sheet_name
            rows_data.append(row_dict)
        
        add_log(f"Successfully processed {len(rows_data)} records from {sheet_name}", "success")
        return rows_data
    
    except Exception as e:
        log_file_error("Excel sheet processing", sheet_name, e)
        return []

def process_inpi_data(data):
    """Process INPI data with real-time progress updates"""
    add_log("Starting INPI data processing...")
    
    # Get session with cookies for API requests
    api_session = get_session_with_cookies()
    if not api_session:
        log_auth_error("INPI API", "session_failed", "Could not establish session with cookies")
        return False
    
    # Get session with cookies for PDF downloads  
    pdf_session = get_session_with_cookies()
    if not pdf_session:
        log_auth_error("INPI PDF", "session_failed", "Could not establish session with cookies")
        return False
    
    # Create progress components
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    success_count = 0
    error_count = 0
    url_found_count = 0
    email_found_count = 0
    
    total_items = len(data['data'])
    add_log(f"Processing {total_items} records...")
    
    for i, item in enumerate(data['data'], 1):
        acta = item.get('Acta')
        if not acta:
            add_log(f"Item {i}: No acta number found, skipping", "error")
            error_count += 1
            continue
        
        # Update progress
        progress = i / total_items
        progress_bar.progress(progress)
        status_text.text(f"Processing record {i}/{total_items}: Acta {acta}")
        
        # Build API URL
        api_url = f"https://portaltramites.inpi.gob.ar/Home/GrillaDigitales?limit=100&offset=0&search=&sort=&order=asc&acta={acta}&direccion=1"
        
        try:
            # Make API request
            response = api_session.get(api_url)
            
            if response.status_code == 200:
                add_log(f"Item {i}: Acta {acta} - API SUCCESS")
                success_count += 1
                
                # Parse response and find Formulario item
                try:
                    response_data = response.json()
                    formulario_data, error_msg = find_formulario_item(response_data)
                    
                    if formulario_data:
                        # Construct document URL
                        document_url = construct_document_url(
                            formulario_data['id_documento_encriptado'],
                            formulario_data['filename']
                        )
                        add_log(f"  -> Document URL found")
                        url_found_count += 1
                        
                        # Download PDF and extract email
                        pdf_content, pdf_error = download_pdf_with_retry(pdf_session, document_url)
                        
                        if pdf_content:
                            email, email_error = extract_email_from_pdf(pdf_content)
                            
                            if email:
                                add_log(f"  -> Email found: {email}")
                                email_found_count += 1
                                # Store email in the item
                                item['email_found'] = email
                            else:
                                add_log(f"  -> No email found: {email_error}")
                        else:
                            log_file_error("PDF download", document_url, pdf_error)
                        
                        # Add delay
                        if i < total_items:
                            time.sleep(random.uniform(0.3, 1.2))
                            
                    else:
                        add_log(f"  -> WARNING: {error_msg}")
                        
                except json.JSONDecodeError as e:
                    log_api_error("INPI", api_url, response.status_code, "Invalid JSON response", str(e))
                except Exception as e:
                    log_api_error("INPI", api_url, response.status_code, str(e))
                    
            else:
                log_api_error("INPI", api_url, response.status_code, f"HTTP request failed for Acta {acta}")
                error_count += 1
                
        except Exception as e:
            log_api_error("INPI", api_url, None, f"Request exception for Acta {acta}: {str(e)}")
            error_count += 1
        
        # Add delay between requests
        if i < total_items:
            time.sleep(random.uniform(0.3, 1.2))
    
    # Final progress update
    progress_bar.progress(1.0)
    status_text.text("Processing completed!")
    
    # Print summary
    add_log(f"=== INPI PROCESSING SUMMARY ===", "success")
    add_log(f"Total items processed: {total_items}", "success")
    add_log(f"Successful requests: {success_count}", "success")
    add_log(f"Failed requests: {error_count}", "success" if error_count == 0 else "error")
    add_log(f"Document URLs found: {url_found_count}", "success")
    add_log(f"Emails extracted: {email_found_count}", "success")
    
    return True

def send_emails(data):
    """Send emails for processed data with automatic batch processing"""
    import gc
    
    add_log("Starting automatic batch email processing...")
    
    # Send webhook with original data before starting email process
    send_webhook(data)
    
    # Configuration
    BATCH_SIZE = 20
    BATCH_DELAY = 1.0  # 1 second between batches
    
    # Get items to process
    items_with_emails = [item for item in data['data'] if item.get('email_found')]
    total_emails = len(items_with_emails)
    
    if not items_with_emails:
        add_log("No emails found to send to", "error")
        return
    
    add_log(f"Found {total_emails} records with emails")
    
    # Generate campana_tag from source filename
    source_filename = data["metadata"]["source_file"]
    campana_tag = None
    match = re.match(r'^(\d+)', source_filename)
    if match:
        numero_boletin = match.group(1)
        campana_tag = f"part-{numero_boletin}"
        add_log(f"Generated campaign tag: {campana_tag}")
    else:
        add_log(f"Warning: Filename '{source_filename}' does not start with a number - no campaign tag generated", "warning")
    
    # Create progress components
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # Initialize counters
    emails_sent = 0
    emails_failed = 0
    batch_number = 1
    
    # Process all batches in continuous loop
    current_idx = 0
    while current_idx < total_emails:
        
        # Calculate current batch
        end_idx = min(current_idx + BATCH_SIZE, total_emails)
        batch_items = items_with_emails[current_idx:end_idx]
        
        add_log(f"Processing batch {batch_number}: emails {current_idx + 1}-{end_idx} of {total_emails}")
        
        # Process current batch
        batch_sent, batch_failed = process_email_batch(
            batch_items, current_idx, total_emails, progress_bar, status_text, campana_tag
        )
        
        # Update counters
        emails_sent += batch_sent
        emails_failed += batch_failed
        current_idx = end_idx
        batch_number += 1
        
        # Memory cleanup after each batch
        gc.collect()
        add_log(f"✅ Batch {batch_number-1} completed: {batch_sent} sent, {batch_failed} failed")
        
        # Add delay between batches (except for last batch)
        if current_idx < total_emails:
            add_log(f"⏳ Next batch starts in {BATCH_DELAY} seconds...")
            time.sleep(BATCH_DELAY)
    
    # Final progress update
    progress_bar.progress(1.0)
    status_text.text("Email sending completed!")
    
    # Print summary
    add_log(f"=== EMAIL SENDING SUMMARY ===", "success")
    add_log(f"Total emails attempted: {min(current_idx, total_emails)}", "success")
    add_log(f"Emails sent successfully: {emails_sent}", "success")
    add_log(f"Emails failed: {emails_failed}", "success" if emails_failed == 0 else "error")


def process_email_batch(batch_items, start_idx, total_emails, progress_bar, status_text, campana_tag):
    """Process a single batch of emails"""
    # OpenAI and Brevo configuration using Streamlit secrets
    openai.api_key = st.secrets["OPENAI_API_KEY"]
    BREVO_API_KEY = st.secrets["BREVO_API_KEY"]
    BREVO_URL = st.secrets["BREVO_URL"]
    WHATSAPP_PHONE = st.secrets["WHATSAPP_PHONE"]
    
    success_count = 0
    failed_count = 0
    
    for i, item in enumerate(batch_items, 1):
        current_email_idx = start_idx + i
        progress = current_email_idx / total_emails
        progress_bar.progress(progress)
        status_text.text(f"Sending email {current_email_idx}/{total_emails}")
        
        try:
            # Build prompt for email generation (same as original)
            prompt = f"""
            Actúa como un abogado especialista en propiedad intelectual en Argentina, que trabaja para el estudio jurídico Eguía, líder en registros de marcas. Escribe un email claro, profesional y persuasivo, destinado a un titular de una marca que recibió una oposición a su solicitud ante el INPI.

            Objetivo: Ofrecer nuestros servicios como representantes legales para acompañarlo en el proceso de defensa y registro exitoso de su marca.

            Datos del caso:
            - Nombre del titular: {item.get("Titulares", "N/A")}
            - Denominación de la marca: {item.get("Denominacion", "N/A")}
            - Clase: {item.get("Clase", "N/A")}
            - Número de acta: {item.get("Acta", "N/A")}
            - Fecha de publicación: {item.get("Fecha", "N/A")}
            - Cantidad de oposiciones: {item.get("Oposiciones", "N/A")}

            Instrucciones:
            - Comienza con un saludo personalizado (usa el nombre completo del titular).
            - Informa con precisión que su marca "{item.get("Denominacion", "N/A")}", clase {item.get("Clase", "N/A")}, ha recibido una oposición en el proceso de registro ante el INPI.
            - Explica brevemente qué significa una oposición y qué implicancias tiene (puede afectar el registro de su marca).
            - Presenta al Estudio Eguía como un equipo experto en defensa de marcas con amplia experiencia en resolver oposiciones.
            - Ofrece una consulta gratuita para analizar el caso sin compromiso.
            - Muestra empatía y transmite seguridad profesional.
            - Firma como "Estudio Eguía – Marcas y Patentes".
            - No escribas un asunto.

            Tono: Profesional, cercano, claro, sin tecnicismos innecesarios. Evita sonar como spam. La redacción debe invitar al titular a responder o agendar una llamada.
            """
            
            # Generate email content
            add_log(f"Generating email content for {item.get('Titulares', 'N/A')}")
            
            try:
                response = openai.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[
                        {"role": "system", "content": "Eres un abogado experto en propiedad intelectual."},
                        {"role": "user", "content": prompt}
                    ],
                    temperature=0.7,
                    max_tokens=700,
                    timeout=30
                )
                
                email_content = response.choices[0].message.content
                add_log(f"Email content generated successfully")
                
            except Exception as openai_error:
                log_api_error("OpenAI", "https://api.openai.com/v1/chat/completions", None, 
                            f"Email generation failed for {item.get('Titulares', 'N/A')}", str(openai_error))
                failed_count += 1
                continue
            
            # Send email via Brevo
            headers = {
                "accept": "application/json",
                "content-type": "application/json",
                "api-key": BREVO_API_KEY
            }
            
            # Create simplified HTML email (optimized for memory)
            html_content = f"""
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                    .logo {{ text-align: center; margin-bottom: 30px; }}
                    .content {{ margin: 20px 0; }}
                    .whatsapp-cta {{ text-align: center; margin: 30px 0; }}
                    .whatsapp-btn {{ 
                        display: inline-block;
                        background-color: #25D366;
                        color: white;
                        padding: 15px 30px;
                        text-decoration: none;
                        border-radius: 25px;
                        font-weight: bold;
                        font-size: 16px;
                        transition: background-color 0.3s;
                    }}
                    .whatsapp-btn:hover {{ 
                        background-color: #1da851;
                    }}
                    .footer {{ 
                        margin-top: 40px; 
                        padding-top: 20px; 
                        border-top: 2px solid #1f4e79;
                        font-size: 12px;
                        color: #666;
                    }}
                    .footer strong {{ color: #1f4e79; }}
                </style>
            </head>
            <body>
                <div class="logo">
                    <img src="https://eguia.com.ar/wp-content/uploads/2024/05/Eguia-Logo-png.webp" 
                         alt="Estudio Eguía Logo" 
                         style="max-width: 250px; height: auto;">
                </div>
                
                <div class="content">
                    {email_content.replace(chr(10), '<br>') if email_content else ''}
                    
                    <div class="whatsapp-cta">
                        <a href="https://wa.me/{WHATSAPP_PHONE}?text=Hola%21%20Me%20contactaron%20por%20una%20oposici%C3%B3n%20a%20mi%20marca%2C%20quisiera%20saber%20m%C3%A1s%20informaci%C3%B3n.%20Mi%20nombre%20es%3A%0A"
                           class="whatsapp-btn" 
                           target="_blank">     
                            📱 Contactar por WhatsApp
                        </a>
                    </div>
                </div>
                
                <div class="footer">
                    <strong>Nicolas Eguía Cima</strong><br>
                    Dirección<br><br>
                    
                    <strong>Móvil:</strong> +54 9 351 5114133<br>
                    <strong>Teléfono:</strong> +54 0351 4812200<br><br>
                    
                    Tristán Malbrán 4011 - Piso 2 Of. 1<br>
                    Cerro de las Rosas - CP: 5009ACE - Córdoba - Argentina<br><br>
                    
                    <strong>Redes:</strong> @eguiamarcasypatentes<br><br>
                    
                    <strong>Nosotros:</strong> eguia.com.ar<br><br>
                    
                    <em>CÓRDOBA - ROSARIO - MENDOZA - BUENOS AIRES - LA RIOJA - TUCUMÁN</em>
                </div>
            </body>
            </html>
            """

            payload = {
                "sender": {
                    "name": "Estudio Eguía",
                    "email": "nicolas@eguia.com.ar"
                },
                "replyTo": {
                    "name": "Estudio Eguía",
                    "email": "fgarzon@eguia.com.ar"
                },
                "to": [
                    {
                        "email": item.get('email_found'),
                        "name": item.get("Titulares", "N/A")
                    }
                ],
                "subject": f"Oposición a su marca '{item.get('Denominacion', 'N/A')}' - Estudio Eguía",
                "htmlContent": html_content
            }
            
            # Add campaign tag if available
            if campana_tag:
                payload["tags"] = [campana_tag]
            
            email_response = requests.post(BREVO_URL, headers=headers, data=json.dumps(payload), timeout=30)
            
            if email_response.status_code == 201:
                add_log(f"✅ Email sent successfully to {item.get('email_found')}", "success")
                success_count += 1
            else:
                # Check for IP authorization errors specifically
                if email_response.status_code in [401, 403]:
                    error_text = email_response.text.lower()
                    if 'ip' in error_text and ('not authorized' in error_text or 'forbidden' in error_text or 'unauthorized' in error_text):
                        log_auth_error("Brevo", email_response.status_code, "IP address not authorized - add your IP to Brevo authorized list")
                    else:
                        log_auth_error("Brevo", email_response.status_code, "Authentication/Authorization failed")
                else:
                    log_api_error("Brevo Email", BREVO_URL, email_response.status_code, 
                                f"Failed to send email to {item.get('email_found')}", 
                                email_response.text)
                
                failed_count += 1
                
        except Exception as e:
            add_log(f"❌ Error sending email for {item.get('Titulares', 'N/A')}: {str(e)}", "error", include_traceback=True)
            failed_count += 1
        
        # Add delay between emails (reduced from original)
        if i < len(batch_items):
            time.sleep(random.uniform(0.5, 1.0))
    
    return success_count, failed_count

def generate_comprehensive_json(data):
    """Generate comprehensive JSON with all XLS data + emails"""
    try:
        # Create comprehensive export
        export_data = {
            "metadata": {
                "source_file": data["metadata"]["source_file"],
                "processing_date": datetime.now().isoformat(),
                "total_records": len(data["data"]),
                "records_with_emails": len([item for item in data["data"] if item.get("email_found")]),
                "sheets_processed": data["metadata"]["sheets_processed"]
            },
            "records": data["data"]
        }
        
        # Save to file
        filename = f"comprehensive_data_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(export_data, f, ensure_ascii=False, indent=2)
        
        add_log(f"✅ Comprehensive JSON exported: {filename}", "success")
        return filename, export_data
        
    except Exception as e:
        log_file_error("JSON export generation", "comprehensive_data_export", e)
        return None, None

def send_webhook(data):
    """Send webhook notification with comprehensive JSON data"""
    try:
        # Get webhook URL from secrets
        webhook_url = st.secrets.get("WEBHOOK_URL")
        if not webhook_url:
            add_log("❌ Webhook URL not configured in secrets", "error")
            return False
        
        add_log("Sending webhook notification...")
        
        # Generate the comprehensive data
        _, export_data = generate_comprehensive_json(data)
        if not export_data:
            add_log("❌ Failed to generate webhook payload", "error")
            return False
        
        # Send HTTP POST request with JSON payload
        headers = {
            "Content-Type": "application/json",
            "User-Agent": "Estudio-Eguia-INPI-Automation/1.0"
        }
        
        response = requests.post(
            webhook_url,
            headers=headers,
            data=json.dumps(export_data, ensure_ascii=False),
            timeout=30
        )
        
        if response.status_code in [200, 201, 202, 204]:
            add_log(f"✅ Webhook sent successfully (Status: {response.status_code})", "success")
            return True
        else:
            log_api_error("Webhook", webhook_url, response.status_code, 
                         "Webhook request failed", response.text[:200])
            return False
            
    except requests.exceptions.Timeout:
        add_log("❌ Webhook request timed out after 30 seconds", "error")
        return False
    except requests.exceptions.ConnectionError:
        add_log("❌ Webhook connection failed - check URL and network", "error")
        return False
    except Exception as e:
        add_log(f"❌ Webhook error: {str(e)}", "error", include_traceback=True)
        return False

# Header
st.markdown('<h1 class="main-header">⚖️ INPI Automatización</h1>', unsafe_allow_html=True)
st.markdown('<h3 style="text-align: center; color: #6c757d;">Estudio Eguía - Marcas y Patentes</h3>', unsafe_allow_html=True)

# Progress indicator
progress_steps = ["📁 Cargar", "🔍 Buscar", "📧 Enviar"]
cols = st.columns(3)
for idx, (col, step_name) in enumerate(zip(cols, progress_steps)):
    with col:
        if idx + 1 < st.session_state.step:
            st.success(f"✅ {step_name}")
        elif idx + 1 == st.session_state.step:
            st.info(f"▶️ {step_name}")
        else:
            st.write(f"⏳ {step_name}")

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# Step 1: Upload File
if st.session_state.step == 1:
    st.header("📁 Paso 1: Cargar Excel")
    
    uploaded_file = st.file_uploader(
        "El archivo Excel que contenga las hojas OPOSICIONES y VISTAS",
        type=['xls'],
        help="Cargue el archivo Excel desde INPI",
        key="excel_uploader_step1"
    )
    
    if uploaded_file is not None:
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Nombre del Archivo", uploaded_file.name)
        with col2:
            st.metric("Tamaño del Archivo", f"{uploaded_file.size / 1024:.1f} KB")
        
        if st.button("🔄 Procesar Archivo", type="primary"):
            try:
                add_log("=== STARTING FILE PROCESSING ===")
                add_log(f"Processing file: {uploaded_file.name}")
                
                # Save uploaded file temporarily
                temp_file_path = f"temp_{uploaded_file.name}"
                with open(temp_file_path, "wb") as f:
                    f.write(uploaded_file.getvalue())
                
                # Process using existing logic
                excel = pd.ExcelFile(temp_file_path, engine='xlrd')
                sheet_names = excel.sheet_names
                add_log(f"Found sheets: {sheet_names}")
                
                target_sheets = ["OPOSICIONES", "VISTAS"]
                all_part_rows = []
                
                # Create metadata
                metadata = {
                    "source_file": uploaded_file.name,
                    "processing_date": datetime.now().isoformat(),
                    "sheets_processed": []
                }
                
                for sheet in target_sheets:
                    matching_sheets = [s for s in sheet_names if s.upper() == sheet.upper()]
                    
                    if matching_sheets:
                        actual_sheet_name = matching_sheets[0]
                        add_log(f"Processing sheet: {actual_sheet_name}")
                        
                        df = pd.read_excel(excel, sheet_name=actual_sheet_name)
                        sheet_rows = process_sheet(df, actual_sheet_name)
                        
                        if sheet_rows:
                            all_part_rows.extend(sheet_rows)
                            metadata["sheets_processed"].append({
                                "name": actual_sheet_name,
                                "rows_found": len(sheet_rows)
                            })
                    else:
                        add_log(f"WARNING: No sheet found matching '{sheet}'")
                
                if all_part_rows:
                    # Store processed data
                    st.session_state.uploaded_data = {
                        "metadata": metadata,
                        "data": all_part_rows
                    }
                    
                    # Save to JSON for compatibility with other scripts
                    with open('part_data.json', 'w', encoding='utf-8') as f:
                        json.dump(st.session_state.uploaded_data, f, ensure_ascii=False, indent=2)
                    
                    add_log(f"=== FILE PROCESSING COMPLETED ===", "success")
                    add_log(f"Total records found: {len(all_part_rows)}", "success")
                    
                    # Move to next step
                    st.session_state.step = 2
                    st.rerun()
                else:
                    add_log("❌ No records with Agente = 'Part.' found in the file", "error")
                
                # Clean up temp file
                os.remove(temp_file_path)
                
            except Exception as e:
                log_file_error("Excel file processing", uploaded_file.name, e)
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)

# Step 2: Process INPI Data
elif st.session_state.step == 2:
    st.header("🔍 Paso 2: Buscar emails en INPI")
    
    if st.session_state.uploaded_data:
        data = st.session_state.uploaded_data
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Registros", len(data["data"]))
        with col2:
            st.metric("Archivo Fuente", data["metadata"]["source_file"])
        with col3:
            st.metric("Hojas Procesadas", len(data["metadata"]["sheets_processed"]))
        
        st.info("Este proceso buscará el email de cada registro en el portal de INPI. Esto puede tomar varios minutos.")
        
        if st.button("🚀 Iniciar búsqueda en INPI", type="primary"):
            if process_inpi_data(data):
                st.session_state.processed_data = data
                
                # Generate comprehensive JSON export
                json_filename, _ = generate_comprehensive_json(data)
                if json_filename:
                    with open(json_filename, 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="📥 Descargar JSON Completo",
                            data=f.read(),
                            file_name=json_filename,
                            mime="application/json"
                        )
                
                st.session_state.step = 3
                st.rerun()
        
# Step 3: Send Emails
elif st.session_state.step == 3:
    st.header("📧 Paso 3: Enviar Emails")
    
    if st.session_state.processed_data:
        data = st.session_state.processed_data
        
        items_with_emails = [item for item in data['data'] if item.get('email_found')]
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Total Registros", len(data["data"]))
        with col2:
            st.metric("Registros con Emails", len(items_with_emails))
        
        if len(items_with_emails) > 0:
            st.info(f"Listo para enviar emails a {len(items_with_emails)} destinatarios")
            
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📧 Enviar Todos los Emails", type="primary"):
                    send_emails(data)
                
                # Generate and download comprehensive JSON
                json_filename, _ = generate_comprehensive_json(data)
                if json_filename:
                    with open(json_filename, 'r', encoding='utf-8') as f:
                        st.download_button(
                            label="📥 Descargar JSON Completo",
                            data=f.read(),
                            file_name=json_filename,
                            mime="application/json"
                        )
        else:
            st.warning("No se encontraron direcciones de email. El procesamiento de INPI puede ser necesario primero.")
        
        # Reset button
        if st.button("🔄 Reiniciar"):
            # Clear session state
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()
