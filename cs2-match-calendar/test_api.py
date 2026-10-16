
import os
import requests
from datetime import datetime

API_KEY = os.getenv("CITO_API_KEY")

if not API_KEY:
    raise RuntimeError("CITO_API_KEY não configurada.")

URL = "https://api.citoapi.com/api/v1/cs2/matches/upcoming"

# Troque estes exemplos pelos seus times favoritos.
FAVORITE_TEAMS = [
    "FURIA",
    "Legacy",
]

response = requests.get(
    URL,
    headers={"x-api-key": API_KEY},
    timeout=20,
)
response.raise_for_status()

matches = response.json().get("data", [])

# Considera apenas partidas futuras com status upcoming.
upcoming = [
    match for match in matches
    if match.get("status") == "upcoming"
    and match.get("startsAt")
]

# Filtra partidas em que pelo menos um time favorito participa.
filtered = []

for match in upcoming:
    team1 = match.get("team1Name") or ""
    team2 = match.get("team2Name") or ""

    if any(
        favorite.casefold() in team.casefold()
        for favorite in FAVORITE_TEAMS
        for team in (team1, team2)
    ):
        filtered.append(match)

# Ordena pela data da partida.
filtered.sort(key=lambda match: match["startsAt"])

from datetime import timezone, timedelta

local_tz = timezone(timedelta(hours=-3))

print(f"Partidas recebidas: {len(matches)}")
print(f"Partidas futuras encontradas: {len(filtered)}")

for match in filtered:
    start = datetime.fromisoformat(
        match["startsAt"].replace("Z", "+00:00")
    ).astimezone(local_tz)

    print("\n" + "=" * 45)
    print(f"{match.get('team1Name')} x {match.get('team2Name')}")
    print(f"Campeonato: {match.get('eventName') or 'Não informado'}")
    print("Horário local:", start.strftime("%d/%m/%Y às %H:%M"))
    print("ID:", match.get("id"))

if not filtered:
    print(
        "\nNenhuma partida correspondente foi encontrada "
        "entre os resultados recebidos."
    )

