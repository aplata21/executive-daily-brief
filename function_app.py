import azure.functions as func
import logging
import os
import requests
import msal
from openai import OpenAI
from datetime import datetime, timedelta, timezone

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

@app.timer_trigger(schedule="0 0 11 * * *", arg_name="mytimer", run_on_startup=False)
def executive_daily_brief(mytimer: func.TimerRequest) -> None:

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

    since = (datetime.now(timezone.utc) - timedelta(days=4)).isoformat()

    url = (
        f"https://graph.microsoft.com/v1.0/users/{user_email}/mailFolders/inbox/messages"
        f"?$top=50"
        f"&$filter=receivedDateTime ge {since}"
        f"&$select=id,subject,bodyPreview,from"
    )

    response = requests.get(
        url,
        headers={"Authorization": f"Bearer {access_token}"}
    )

    emails = response.json().get("value", [])

    context = ""

    for email in emails:
        context += f'''
From: {email.get("from", {}).get("emailAddress", {}).get("address")}

Subject: {email.get("subject")}

Body: {email.get("bodyPreview")}

------------------------
'''

    client = OpenAI(api_key=openai_key)

    completion = client.responses.create(
        model="gpt-4.1-mini",
        input=f"""
Act as a senior executive assistant specialized in strategic decision-making.

Analyze these emails from the last 4 days.

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
6. Follow-ups
7. Suggested actions

EMAILS:

{context}
"""
    )

    summary = completion.output_text

    logging.info(summary)

    print(summary)