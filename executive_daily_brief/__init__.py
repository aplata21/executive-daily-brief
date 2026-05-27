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


def _validate_environment() -> Dict[str, str]:
    """Validate and retrieve required environment variables."""
    required_vars = ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "OPENAI_API_KEY", "User_email"]
    config = {}
    
    for var in required_vars:
        value = os.environ.get(var)
        if not value:
            raise ValueError(f"Missing required environment variable: {var}")
        config[var] = value
    
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


def _analyze_emails_with_openai(openai_key: str, email_context: str) -> str:
    """Analyze emails using OpenAI.
    
    Note: Prompt caching (ephemeral or standard) won't benefit daily-scheduled
    functions since cache expires before next execution. Using simpler approach.
    """
    prompt = f"""Analyze these executive emails concisely. Identify:
- Priorities and high-impact items
- Blockers and risks  
- Deadlines and action items

EMAILS:
{email_context}"""

    try:
        client = OpenAI(api_key=openai_key)
        completion = client.chat.completions.create(
            model="gpt-4-mini",
            messages=[{"role": "user", "content": prompt}],
            temperature=0.5,
            max_tokens=800
        )
        
        # Log token usage for cost monitoring
        usage = completion.usage
        logging.info(
            f"OpenAI tokens - Input: {usage.prompt_tokens}, Output: {usage.completion_tokens}, "
            f"Total cost: ~${(usage.prompt_tokens * 0.15 + usage.completion_tokens * 0.60) / 1_000_000:.6f}"
        )
        
        return completion.choices[0].message.content
    except Exception as e:
        logging.error(f"OpenAI analysis failed: {str(e)}")
        raise


def _send_brief_email(user_email: str, access_token: str, summary: str) -> None:
    """Send executive brief email."""
    send_email_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"
    
    email_payload = {
        "message": {
            "subject": "Executive Daily Brief",
            "body": {
                "contentType": "HTML",
                "content": f"<pre>{summary}</pre>"
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
    
    # Process in batches to avoid overwhelming the API
    for batch_num, i in enumerate(range(0, len(email_ids), BATCH_SIZE), 1):
        batch = email_ids[i:i + BATCH_SIZE]
        batch_failures = []
        
        # Retry logic for transient failures
        for attempt in range(MAX_BATCH_RETRIES):
            if attempt > 0:
                logging.info(f"Retrying batch {batch_num} (attempt {attempt + 1}/{MAX_BATCH_RETRIES})...")
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
                
                # Check for individual failures in batch response
                batch_response = response.json()
                batch_failures = []
                for item in batch_response.get("responses", []):
                    if item.get("status") >= 400:
                        req_id = int(item.get("id", -1))
                        if 0 <= req_id < len(batch):
                            batch_failures.append(batch[req_id])
                
                if not batch_failures:
                    # Batch succeeded completely
                    logging.info(f"Marked {len(batch)} emails as read (batch {batch_num}/{(len(email_ids) + BATCH_SIZE - 1) // BATCH_SIZE}).")
                    break
                elif attempt < MAX_BATCH_RETRIES - 1:
                    # Some failed, retry
                    logging.warning(f"Batch {batch_num}: {len(batch_failures)} emails failed, retrying...")
                    batch = batch_failures  # Retry only failed ones
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
        logging.warning(f"Failed to mark {len(total_failed)} emails as read after retries. IDs: {total_failed[:5]}{'...' if len(total_failed) > 5 else ''}")


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
        summary = _analyze_emails_with_openai(config["OPENAI_API_KEY"], email_context)
        
        # Send brief
        _send_brief_email(config["User_email"], access_token, summary)
        
        # Mark as read (batch operation)
        email_ids = [email.get("id") for email in emails if email.get("id")]
        _mark_emails_as_read_batch(config["User_email"], access_token, email_ids)
        
        logging.info(f"Executive Daily Brief completed successfully. Processed {len(emails)} emails.")
        
    except Exception as e:
        logging.error(f"Executive Daily Brief failed: {str(e)}")
        raise