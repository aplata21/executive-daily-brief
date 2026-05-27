import logging
import os
import requests
import msal
import azure.functions as func
from openai import OpenAI
from datetime import datetime, timedelta, timezone


def main(mytimer: func.TimerRequest) -> None:
    logging.info("Starting Executive Daily Brief")

    tenant_id = os.environ["TENANT_ID"]
    client_id = os.environ["CLIENT_ID"]
    client_secret = os.environ["CLIENT_SECRET"]
    openai_key = os.environ["OPENAI_API_KEY"]
    user_email = os.environ["User_email"]

    app_msal = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret
    )

    token_response = app_msal.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    access_token = token_response["access_token"]

    since = (
    datetime.now(timezone.utc)
    - timedelta(days=4)
).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    url = (
    f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
    f"?$top=50"
    f"&$filter=(receivedDateTime ge {since}) and (isRead eq false)"
    f"&$select=id,subject,bodyPreview,from,receivedDateTime,isRead"
)

    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"}
    )
    response.raise_for_status()

    emails = response.json().get("value", [])

    if not emails:
        logging.info("No unread emails found in the last 4 days.")
        return

    context = ""

    for email in emails:
        context += f"""
From: {email.get("from", {}).get("emailAddress", {}).get("address")}

Subject: {email.get("subject")}

Body: {email.get("bodyPreview")}

------------------------
"""

    client = OpenAI(api_key=openai_key)

    completion = client.responses.create(
        model="gpt-4.1-mini",
        input=f"""
Act as a senior executive assistant specialized in strategic decision-making.

Analyze these unread emails from the last 4 days.

Identify:
- priorities
- blockers
- risks
- deadlines
- dependencies
- follow-ups
- action items

Ignore:
- signatures
- newsletters
- repetitive content
- low-value noise

Return:
1. Executive summary
2. High priority items
3. Risks
4. Deadlines
5. Blockers
6. Dependencies
7. Follow-ups
8. Suggested actions

EMAILS:

{context}
"""
    )

    summary = completion.output_text

    logging.info(summary)
    print(summary)

    send_email_url = f"https://graph.microsoft.com/v1.0/users/{user_email}/sendMail"

    email_payload = {
        "message": {
            "subject": "Executive Daily Brief",
            "body": {
                "contentType": "Text",
                "content": summary
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

    send_email_response = requests.post(
        send_email_url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        },
        json=email_payload
    )

    send_email_response.raise_for_status()

    logging.info("Executive Daily Brief email sent successfully.")

    for email in emails:
        message_id = email.get("id")

        if not message_id:
            continue

        mark_read_url = (
            f"https://graph.microsoft.com/v1.0/users/{user_email}/messages/{message_id}"
        )

        mark_read_response = requests.patch(
            mark_read_url,
            headers={
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json"
            },
            json={"isRead": True}
        )

        mark_read_response.raise_for_status()

    logging.info(f"Marked {len(emails)} emails as read.")