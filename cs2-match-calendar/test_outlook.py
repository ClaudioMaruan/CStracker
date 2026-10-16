
import os
import msal
import requests

client_id = os.getenv("MS_CLIENT_ID")

if not client_id:
    raise RuntimeError("A variável MS_CLIENT_ID não foi configurada.")

app = msal.PublicClientApplication(
    client_id=client_id,
    authority="https://login.microsoftonline.com/consumers",
)

scopes = ["Calendars.ReadWrite"]

flow = app.initiate_device_flow(scopes=scopes)

if "user_code" not in flow:
    raise RuntimeError(
        "Não foi possível iniciar o login: "
        + str(flow.get("error_description", flow))
    )

print(flow["message"])

result = app.acquire_token_by_device_flow(flow)

if "access_token" not in result:
    raise RuntimeError(
        "Login não concluído: "
        + str(result.get("error_description", result.get("error")))
    )

response = requests.get(
    "https://graph.microsoft.com/v1.0/me/calendars",
    headers={
        "Authorization": f"Bearer {result['access_token']}"
    },
    timeout=20,
)

response.raise_for_status()

calendars = response.json().get("value", [])

print("\nConexão com o Outlook estabelecida!")
print(f"Calendários encontrados: {len(calendars)}")

for calendar in calendars:
    print("-", calendar.get("name", "Sem nome"))
