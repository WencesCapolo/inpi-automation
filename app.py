import streamlit as st
import pandas as pd
import json
import io
import os
import sys
import time
import traceback
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
    page_title="INPI Automatización - Estudio Eguía",
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

def add_log(message, log_type="info"):
    """Add a log message with timestamp"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.logs.append({
        "timestamp": timestamp,
        "message": message,
        "type": log_type
    })

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
        error_msg = f"Error processing sheet {sheet_name}: {str(e)}"
        add_log(error_msg, "error")
        add_log(f"Full traceback: {traceback.format_exc()}", "error")
        return []

def process_inpi_data(data):
    """Process INPI data with real-time progress updates"""
    add_log("Starting INPI data processing...")
    
    # Get session with cookies for API requests
    api_session = get_session_with_cookies()
    if not api_session:
        add_log("Failed to get session with cookies for API", "error")
        return False
    
    # Get session with cookies for PDF downloads  
    pdf_session = get_session_with_cookies()
    if not pdf_session:
        add_log("Failed to get session with cookies for PDF downloads", "error")
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
                            add_log(f"  -> PDF download failed: {pdf_error}", "error")
                        
                        # Add delay
                        if i < total_items:
                            time.sleep(random.uniform(0.3, 1.2))
                            
                    else:
                        add_log(f"  -> WARNING: {error_msg}")
                        
                except json.JSONDecodeError:
                    add_log(f"  -> ERROR: Invalid JSON response", "error")
                except Exception as e:
                    add_log(f"  -> ERROR: {str(e)}", "error")
                    
            else:
                add_log(f"Item {i}: Acta {acta} - FAILED (Status: {response.status_code})", "error")
                error_count += 1
                
        except Exception as e:
            add_log(f"Item {i}: Acta {acta} - ERROR: {str(e)}", "error")
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
    """Send emails for processed data"""
    add_log("Starting email sending process...")
    
    # OpenAI and Brevo configuration using Streamlit secrets
    openai.api_key = st.secrets["OPENAI_API_KEY"]
    BREVO_API_KEY = st.secrets["BREVO_API_KEY"]
    BREVO_URL = st.secrets["BREVO_URL"]
    WHATSAPP_PHONE = st.secrets["WHATSAPP_PHONE"]
    
    # Filter items that have emails
    items_with_emails = [item for item in data['data'] if item.get('email_found')]
    
    if not items_with_emails:
        add_log("No emails found to send to", "error")
        return
    
    add_log(f"Found {len(items_with_emails)} records with emails")
    
    # Create progress components
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    emails_sent = 0
    emails_failed = 0
    
    for i, item in enumerate(items_with_emails, 1):
        progress = i / len(items_with_emails)
        progress_bar.progress(progress)
        status_text.text(f"Sending email {i}/{len(items_with_emails)}")
        
        try:
            # Build prompt for email generation
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

            Tono: Profesional, cercano, claro, sin tecnicismos innecesarios. Evita sonar como spam. La redacción debe invitar al titular a responder o agendar una llamada.
            """
            
            # Generate email content
            add_log(f"Generating email content for {item.get('Titulares', 'N/A')}")
            
            response = openai.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": "Eres un abogado experto en propiedad intelectual."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
                max_tokens=700
            )
            
            email_content = response.choices[0].message.content
            add_log(f"Email content generated successfully")
            
            # Send email via Brevo
            headers = {
                "accept": "application/json",
                "content-type": "application/json",
                "api-key": BREVO_API_KEY
            }
            
            # Create HTML email with logo and footer
            html_content = f"""
            <html>
            <head>
                <meta charset="UTF-8">
                <style>
                    body {{ font-family: Arial, sans-serif; line-height: 1.6; color: #333; }}
                    .logo {{ text-align: center; margin-bottom: 30px; }}
                    .content {{ margin: 20px 0; }}
                    .whatsapp-cta {{ 
                        text-align: center; 
                        margin: 30px 0; 
                    }}
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
                        <a href="https://api.whatsapp.com/send?phone={WHATSAPP_PHONE}" 
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
                "to": [
                    {
                        "email": item.get('email_found'),
                        "name": item.get("Titulares", "N/A")
                    }
                ],
                "subject": f"Oposición a su marca '{item.get('Denominacion', 'N/A')}' - Estudio Eguía",
                "htmlContent": html_content
            }
            
            email_response = requests.post(BREVO_URL, headers=headers, data=json.dumps(payload))
            
            if email_response.status_code == 201:
                add_log(f"✅ Email sent successfully to {item.get('email_found')}", "success")
                emails_sent += 1
            else:
                add_log(f"❌ Failed to send email to {item.get('email_found')} - Status: {email_response.status_code}", "error")
                add_log(f"Error details: {email_response.text}", "error")
                emails_failed += 1
                
        except Exception as e:
            add_log(f"❌ Error sending email for {item.get('Titulares', 'N/A')}: {str(e)}", "error")
            add_log(f"Full traceback: {traceback.format_exc()}", "error")
            emails_failed += 1
        
        # Add delay between emails
        if i < len(items_with_emails):
            time.sleep(random.uniform(1.0, 2.0))
    
    # Final progress update
    progress_bar.progress(1.0)
    status_text.text("Email sending completed!")
    
    # Print summary
    add_log(f"=== EMAIL SENDING SUMMARY ===", "success")
    add_log(f"Total emails attempted: {len(items_with_emails)}", "success")
    add_log(f"Emails sent successfully: {emails_sent}", "success")
    add_log(f"Emails failed: {emails_failed}", "success" if emails_failed == 0 else "error")

# Header
st.markdown('<h1 class="main-header">⚖️ INPI Automatización</h1>', unsafe_allow_html=True)
st.markdown('<h3 style="text-align: center; color: #6c757d;">Estudio Eguía - Marcas y Patentes</h3>', unsafe_allow_html=True)

# Progress indicator
progress_steps = ["📁 Cargar", "🔍 Procesar", "📧 Enviar"]
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
    st.header("📁 Paso 1: Cargar Archivo XLS")
    
    uploaded_file = st.file_uploader(
        "Elija un archivo XLS que contenga las hojas OPOSICIONES y VISTAS",
        type=['xls'],
        help="Cargue el archivo Excel desde INPI"
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
                add_log(f"❌ Error processing file: {str(e)}", "error")
                add_log(f"Full traceback: {traceback.format_exc()}", "error")
                if os.path.exists(temp_file_path):
                    os.remove(temp_file_path)

# Step 2: Process INPI Data
elif st.session_state.step == 2:
    st.header("🔍 Paso 2: Procesar Datos de INPI")
    
    if st.session_state.uploaded_data:
        data = st.session_state.uploaded_data
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Registros", len(data["data"]))
        with col2:
            st.metric("Archivo Fuente", data["metadata"]["source_file"])
        with col3:
            st.metric("Hojas Procesadas", len(data["metadata"]["sheets_processed"]))
        
        st.info("Este proceso buscará datos adicionales en el portal de INPI para cada registro. Esto puede tomar varios minutos.")
        
        if st.button("🚀 Iniciar Procesamiento de INPI", type="primary"):
            if process_inpi_data(data):
                st.session_state.processed_data = data
                st.session_state.step = 3
                st.rerun()
        
        # Button to skip to email step (for testing)
        if st.button("⏭️ Saltar Procesamiento de INPI"):
            st.session_state.processed_data = st.session_state.uploaded_data
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
            
            if st.button("📧 Enviar Todos los Emails", type="primary"):
                send_emails(data)
        else:
            st.warning("No se encontraron direcciones de email. El procesamiento de INPI puede ser necesario primero.")
        
        # Reset button
        if st.button("🔄 Reiniciar"):
            # Clear session state
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

# Always show logs at the bottom
st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
st.header("📋 Registros de Procesamiento")
display_logs()

# Clear logs button
if st.session_state.logs:
    if st.button("🗑️ Clear Logs"):
        st.session_state.logs = []
        st.rerun()
