
import os
import msal
import requests
from datetime import datetime, timedelta, timezone

client_id = os.getenv("MS_CLIENT_ID")

app = msal.PublicClientApplication(
    client_id=client_id,
    authority="https://login.microsoftonline.com/consumers",
)

flow = app.initiate_device_flow(scopes=["Calendars.ReadWrite"])

if "user_code" not in flow:
    raise RuntimeError(str(flow))

print(flow["message"])
result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise RuntimeError(
        result.get("error_description", result.get("error"))
    )

headers = {
    "Authorization": f"Bearer {result['access_token']}",
    "Content-Type": "application/json",
}

# Evento de teste para amanhã, às 15h no horário de Brasília.
start_local = datetime.now(timezone.utc).astimezone(
    timezone(timedelta(hours=-3))
).replace(hour=15, minute=0, second=0, microsecond=0) + timedelta(days=1)

end_local = start_local + timedelta(minutes=30)

def to_utc_string(dt):
    return dt.astimezone(timezone.utc).isoformat(
        timespec="seconds"
    ).replace("+00:00", "Z")

event = {
    "subject": "TESTE - CS2 Match Calendar",
    "body": {
        "contentType": "Text",
        "content": "Evento de teste da integração Python + Outlook."
    },
    "start": {
        "dateTime": to_utc_string(start_local),
        "timeZone": "UTC"
    },
    "end": {
        "dateTime": to_utc_string(end_local),
        "timeZone": "UTC"
    }
}

response = requests.post(
    "https://graph.microsoft.com/v1.0/me/events",
    headers=headers,
    json=event,
    timeout=20,
)

print("Status HTTP:", response.status_code)

response.raise_for_status()

created = response.json()
print("Evento criado:", created.get("subject"))
print("ID do evento:", created.get("id"))
print("Horário local:", start_local.strftime("%d/%m/%Y às %H:%M"))
