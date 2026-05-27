import logging
import os
import time
import requests
import msal
import azure.functions as func
from openai import OpenAI
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict, List, Tuple


# Configuration constants
REQUEST_TIMEOUT = 30
MAX_EMAILS = 100
MAX_PROMPT_CHARS = 8000  # Reduced for token efficiency
BATCH_SIZE = 20
MAX_BATCH_RETRIES = 3
BATCH_RETRY_DELAY = 2  # seconds
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"


def _validate_environment() -> Dict[str, str]:
    """Validate and retrieve required environment variables."""
    required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "OPENAI_API_KEY", "User_email"]
    optional_vars = ["OPENAI_MODEL"]
    config = {}
    
    for var in required_vars:
        value = os.environ.get(var)
        if not value:
            raise ValueError(f"Missing required environment variable: {var}")
        config[var] = value
    
    for var in optional_vars:
        config[var] = os.environ.get(var, "")
    
    if not config["OPENAI_MODEL"]:
        config["OPENAI_MODEL"] = DEFAULT_OPENAI_MODEL
    
    return config


def _get_access_token(tenant_id: str, client_id: str, client_secret: str) -> str:
    """Acquire access token via MSAL with error handling."""
    try:
        app_msal = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret
        )
        token_response = app_msal.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        
        if "error" in token_response:
            raise Exception(f"Token acquisition failed: {token_response.get('error_description')}")
        
        return token_response["access_token"]
    except Exception as e:
        logging.error(f"Authentication failed: {str(e)}")
        raise


def _fetch_unread_emails(user_email: str, access_token: str) -> List[Dict]:
    """Fetch unread emails from the last 4 days with pagination support."""
    since = (
        datetime.now(timezone.utc)
        - timedelta(days=4)
    ).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    url = (
        f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
        f"?$top={MAX_EMAILS}"
        f"&$filter=(receivedDateTime ge {since}) and (isRead eq false)"
        f"&$select=id,subject,bodyPreview,from,receivedDateTime"
        f"&$orderby=receivedDateTime desc"
    )

    try:
        response = requests.get(
            url,
            headers={"Authorization": f"Bearer {access_token}"},
            timeout=REQUEST_TIMEOUT
        )
        response.raise_for_status()
        return response.json().get("value", [])
    except requests.exceptions.RequestException as e:
        logging.error(f"Failed to fetch emails: {str(e)}")
        raise


def _build_email_context(emails: List[Dict]) -> Tuple[str, int]:
    """Build email context for OpenAI analysis with aggressive truncation.
    
    Returns:
        Tuple of (email_context_string, number_of_emails_included)
    """
    email_lines = []
    current_size = 0
    emails_included = 0
    
    # Truncate email body to reduce tokens
    MAX_EMAIL_BODY = 200  # chars per email
    
    for idx, email in enumerate(emails):
        try:
            subject = email.get("subject", "No subject")
            body = email.get("bodyPreview", "No preview")[:MAX_EMAIL_BODY]
            
            # Minimal format to reduce tokens
            email_text = f"{subject}\n{body}\n\n"
            
            if current_size + len(email_text) > MAX_PROMPT_CHARS:
                # Accurate count of skipped emails
                emails_skipped = len(emails) - idx
                email_lines.append(f"\n... ({emails_skipped} more emails excluded due to context limit)")
                break
            
            email_lines.append(email_text)
            current_size += len(email_text)
            emails_included += 1
        except Exception as e:
            logging.warning(f"Error processing email {idx}: {str(e)}")
            continue
    
    context = "".join(email_lines)
    logging.info(f"Email context: {emails_included}/{len(emails)} emails included, {current_size} chars")
    return context, emails_included


def _get_openai_client(openai_key: str) -> OpenAI:
    """Create a public OpenAI client using the official OpenAI endpoint."""
    return OpenAI(api_key=openai_key, base_url="https://api.openai.com/v1")


def _analyze_emails_with_openai(openai_key: str, email_context: str, model: str) -> str:
    """Analyze emails using the public OpenAI API in Spanish with metrics and deeper insights."""
    prompt = (
        "ANALISIS EJECUTIVO PROFUNDO DE CORREOS\n\n"
        "A continuacion hay mensajes de correo que requieren un analisis detallado y analitico. "
        "Genera un resumen EN ESPAÑOL perfectamente estructurado incluyendo:\n\n"
        "1. **CONTEXTO Y METRICAS**: Total de correos, remitentes principales, temas recurrentes, tendencias detectadas.\n"
        "2. **ANALISIS DE PRIORIDADES**: Items criticos con justificacion de por que son prioritarios, impacto estimado, urgencia.\n"
        "3. **RIESGOS Y BLOQUEOS**: Identificar obstaculos, riesgos operacionales, dependencias, impacto si no se resuelven.\n"
        "4. **PLAZOS Y DEADLINES**: Listar fechas criticas con semaforo (Rojo/Amarillo/Verde). Urgencia relativa.\n"
        "5. **ACCIONES RECOMENDADAS**: Acciones concretas por orden de impacto, responsables sugeridos, timing.\n"
        "6. **INTELIGENCIA**: Patrones, cambios en comunicacion, stakeholders clave, oportunidades latentes.\n\n"
        "Se especifico, usa numeros y metricas. NO seas generico. Incluye analisis de causas raiz cuando sea relevante. "
        "Estructura con encabezados claros. No escribas en ingles.\n\n"
        "CORREOS:"
        f"{email_context}"
    )

    try:
        client = _get_openai_client(openai_key)
        completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "Eres un analista de negocios senior que genera resumenes ejecutivos en espanol detallados, analiticos y accionables. Incluye datos, metricas, riesgos y recomendaciones concretas. Se profesional pero directo."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.6,
            max_tokens=1500
        )

        usage = getattr(completion, "usage", None)
        if usage is not None:
            prompt_tokens = getattr(usage, "prompt_tokens", 0)
            completion_tokens = getattr(usage, "completion_tokens", 0)
            logging.info(
                f"OpenAI tokens - Input: {prompt_tokens}, Output: {completion_tokens}, "
                f"Total approximate cost: ~${(prompt_tokens * 0.15 + completion_tokens * 0.60) / 1_000_000:.6f}"
            )

        return completion.choices[0].message.content
    except Exception as e:
        logging.error(f"OpenAI analysis failed: {str(e)}")
        raise


def _send_brief_email(user_email: str, access_token: str, summary: str) -> None:
    """Send executive brief email with professional formatting."""
    send_email_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"
    
    # Format summary for HTML
    html_summary = summary.replace("\n\n", "</p><p>").replace("\n", "<br>")
    timestamp = datetime.now().strftime("%d de %B, %Y - %H:%M UTC")
    
    email_payload = {
        "message": {
            "subject": f"Executive Daily Brief - {datetime.now().strftime('%d/%m/%Y')}",
            "body": {
                "contentType": "HTML",
                "content": (
                    "<!DOCTYPE html><html><head><meta charset='UTF-8'></head><body>"
                    "<div style='font-family: Segoe UI, Arial, sans-serif; max-width: 900px; margin: 0; background: #f5f5f5; padding: 20px;'>"
                    "<div style='background: white; border-radius: 8px; padding: 30px; box-shadow: 0 2px 4px rgba(0,0,0,0.1);'>"
                    "<h1 style='color: #1f4788; margin-top: 0; border-bottom: 3px solid #0078d4; padding-bottom: 15px; font-size: 24px;'>Resumen Ejecutivo Diario</h1>"
                    f"<p style='color: #666; font-size: 13px; margin: 10px 0;'>Generado: {timestamp}</p>"
                    "<hr style='border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;'>"
                    "<div style='color: #333; line-height: 1.8; font-size: 14px;'>"
                    f"<p>{html_summary}</p>"
                    "</div>"
                    "<hr style='border: none; border-top: 1px solid #e0e0e0; margin: 20px 0;'>"
                    "<footer style='color: #999; font-size: 12px; text-align: center; margin-top: 30px;'>"
                    "<p>Este resumen fue generado automaticamente por Executive Daily Brief.</p>"
                    "</footer>"
                    "</div></div></body></html>"
                )
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": user_email
                    }
                }
            ]
        },
        "saveToSentItems": True
    }

    try:
        response = requests.post(
            send_email_url,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            },
            json=email_payload,
            timeout=REQUEST_TIMEOUT
        )
        response.raise_for_status()
        logging.info("Executive Daily Brief sent successfully.")
    except requests.exceptions.RequestException as e:
        logging.error(f"Failed to send brief email: {str(e)}")
        raise


def _mark_emails_as_read_batch(user_email: str, access_token: str, email_ids: List[str]) -> None:
    """Mark emails as read using batch operations with retry logic."""
    if not email_ids:
        return
    
    total_failed = []
    logging.info(f"Intentando marcar {len(email_ids)} correos como leídos.")
    
    # Process in batches to avoid overwhelming the API
    for batch_num, i in enumerate(range(0, len(email_ids), BATCH_SIZE), 1):
        batch = email_ids[i:i + BATCH_SIZE]
        batch_failures = []
        
        # Retry logic for transient failures
        for attempt in range(MAX_BATCH_RETRIES):
            if attempt > 0:
                logging.info(f"Reintentando batch {batch_num} (intento {attempt + 1}/{MAX_BATCH_RETRIES})...")
                time.sleep(BATCH_RETRY_DELAY)
            
            batch_requests = [
                {
                    "id": str(idx),
                    "method": "PATCH",
                    "url": f"/users/{user_email}/messages/{msg_id}",
                    "body": {"isRead": True}
                }
                for idx, msg_id in enumerate(batch)
            ]
            
            batch_payload = {"requests": batch_requests}
            
            try:
                response = requests.post(
                    "https://graph.microsoft.com/v1.0/$batch",
                    headers={
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/json"
                    },
                    json=batch_payload,
                    timeout=REQUEST_TIMEOUT
                )
                response.raise_for_status()
                
                batch_response = response.json()
                batch_failures = []
                for item in batch_response.get("responses", []):
                    if item.get("status", 200) >= 400:
                        req_id = int(item.get("id", -1))
                        if 0 <= req_id < len(batch):
                            batch_failures.append(batch[req_id])
                
                if not batch_failures:
                    logging.info(f"Marked {len(batch)} emails as read (batch {batch_num}/{(len(email_ids) + BATCH_SIZE - 1) // BATCH_SIZE}).")
                    break
                elif attempt < MAX_BATCH_RETRIES - 1:
                    logging.warning(f"Batch {batch_num}: {len(batch_failures)} emails fallaron, reintentando...")
                    batch = batch_failures
                    continue
                    
            except requests.exceptions.RequestException as e:
                if attempt < MAX_BATCH_RETRIES - 1:
                    logging.warning(f"Batch {batch_num} request failed: {str(e)}, retrying...")
                    continue
                else:
                    logging.error(f"Batch {batch_num} failed after {MAX_BATCH_RETRIES} attempts: {str(e)}")
                    batch_failures = batch
        
        if batch_failures:
            total_failed.extend(batch_failures)
    
    if total_failed:
        logging.warning(f"Fallo al marcar {len(total_failed)} correos como leídos. Probando parche individual...")
        for msg_id in total_failed:
            try:
                patch_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{msg_id}"
                patch_response = requests.patch(
                    patch_url,
                    headers={
                        "Authorization": f"Bearer {access_token}",
                        "Content-Type": "application/json"
                    },
                    json={"isRead": True},
                    timeout=REQUEST_TIMEOUT
                )
                patch_response.raise_for_status()
                logging.info(f"Correo {msg_id} marcado como leído individualmente.")
            except requests.exceptions.RequestException as e:
                logging.error(f"No se pudo marcar el correo {msg_id} como leído: {str(e)}")
        logging.warning(f"Proceso de marcado como leído completado con {len(total_failed)} objetos fallidos en batch.")


def main(mytimer: func.TimerRequest) -> None:
    """Main execution function for Executive Daily Brief."""
    try:
        logging.info("Starting Executive Daily Brief")
        
        # Validate configuration
        config = _validate_environment()
        
        # Authenticate
        access_token = _get_access_token(
            config["TENANT_ID"],
            config["CLIENT_ID"],
            config["CLIENT_SECRET"]
        )
        
        # Fetch emails
        emails = _fetch_unread_emails(config["User_email"], access_token)
        
        if not emails:
            logging.info("No unread emails found in the last 4 days.")
            return
        
        logging.info(f"Processing {len(emails)} unread emails.")
        
        # Build context
        email_context, emails_included = _build_email_context(emails)
        
        if not email_context.strip():
            logging.info("No email content to analyze.")
            return
        
        # Analyze with OpenAI
        summary = _analyze_emails_with_openai(
            config["OPENAI_API_KEY"],
            email_context,
            config["OPENAI_MODEL"]
        )
        
        # Send brief
        _send_brief_email(config["User_email"], access_token, summary)
        
        # Mark as read (batch operation)
        email_ids = [email.get("id") for email in emails if email.get("id")]
        _mark_emails_as_read_batch(config["User_email"], access_token, email_ids)
        
        logging.info(f"Executive Daily Brief completed successfully. Processed {len(emails)} emails.")
        
    except Exception as e:
        logging.error(f"Executive Daily Brief failed: {str(e)}")
        raise