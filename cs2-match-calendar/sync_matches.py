
import os
import sys
from datetime import datetime, timedelta, timezone

import msal
import requests


# CONFIGURAÇÕES
CITO_API_URL = "https://api.citoapi.com/api/v1/cs2/matches/upcoming"
GRAPH_URL = "https://graph.microsoft.com/v1.0"

FAVORITE_TEAMS = {"FURIA", "Legacy"}
SCOPES = ["Calendars.ReadWrite"]

# Duração estimada quando a API não informa o fim da partida.
DEFAULT_DURATION_HOURS = 3

# Identificador usado para reconhecer eventos criados por este script.
EVENT_MARKER = "CITO_MATCH_ID:"


def get_environment_variable(name):
    value = os.getenv(name)
    if not value:
        raise RuntimeError(
            f"A variável {name} não está configurada neste terminal."
        )
    return value


def parse_datetime(value):
    """Converte uma data ISO da API para um datetime com fuso horário."""
    if not value:
        return None

    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))

    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)

    return parsed


def is_favorite(team_name):
    """Compara nomes sem diferenciar maiúsculas e minúsculas."""
    if not team_name:
        return False

    team_name = team_name.strip().casefold()

    return any(
        team_name == favorite.casefold()
        or team_name.startswith(favorite.casefold() + " ")
        for favorite in FAVORITE_TEAMS
    )


def get_upcoming_matches(api_key):
    response = requests.get(
        CITO_API_URL,
        headers={"x-api-key": api_key},
        timeout=30,
    )
    response.raise_for_status()

    payload = response.json()

    if not payload.get("success", False):
        raise RuntimeError("A CitoAPI não confirmou o sucesso da consulta.")

    matches = payload.get("data", [])

    now = datetime.now(timezone.utc)
    selected = []

    for match in matches:
        if match.get("status") != "upcoming":
            continue

        start = parse_datetime(match.get("startsAt"))
        if not start or start < now:
            continue

        team1 = match.get("team1Name")
        team2 = match.get("team2Name")

        if not (is_favorite(team1) or is_favorite(team2)):
            continue

        selected.append(match)

    return selected


def authenticate_outlook(client_id):
    app = msal.PublicClientApplication(
        client_id=client_id,
        authority="https://login.microsoftonline.com/consumers",
    )

    flow = app.initiate_device_flow(scopes=SCOPES)

    if "user_code" not in flow:
        raise RuntimeError(
            f"Não foi possível iniciar o login: {flow}"
        )

    print("\nEntre na sua conta Microsoft para autorizar o calendário:")
    print(flow["message"])

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise RuntimeError(
            "A autenticação não foi concluída. "
            + result.get("error_description", "Tente novamente.")
        )

    return result["access_token"]



def get_existing_events(token, start, end):
    """Retorna os eventos sincronizados, indexados pelo ID da partida."""
    headers = {"Authorization": f"Bearer {token}"}

    url = f"{GRAPH_URL}/me/calendarView"
    params = {
        "startDateTime": start.isoformat(),
        "endDateTime": end.isoformat(),
        "$select": "id,subject,body,start,end",
        "$top": 100,
    }

    existing_events = {}

    while url:
        response = requests.get(
            url,
            headers=headers,
            params=params,
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()

        for event in payload.get("value", []):
            body = event.get("body", {}).get("content", "")

            if EVENT_MARKER not in body:
                continue

            marker_value = body.split(EVENT_MARKER, 1)[1]
            match_id = marker_value.split("<", 1)[0].split()[0].strip()

            existing_events[match_id] = event

        url = payload.get("@odata.nextLink")
        params = None

    return existing_events




def build_event_payload(match):
    start = parse_datetime(match["startsAt"])
    end = parse_datetime(match.get("endsAt"))

    if not end or end <= start:
        end = start + timedelta(hours=DEFAULT_DURATION_HOURS)

    team1 = match.get("team1Name") or "Time a confirmar"
    team2 = match.get("team2Name") or "Adversário a confirmar"
    event_name = match.get("eventName") or "Campeonato de CS2"
    match_id = match["id"]

    subject = f"CS2 | {team1} x {team2}"

    body_text = (
        f"Campeonato: {event_name}<br>"
        f"Times: {team1} x {team2}<br>"
        f"ID da partida: {match_id}<br>"
        f"{EVENT_MARKER} {match_id}"
    )

    payload = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body_text,
        },
        "start": {
            "dateTime": start.astimezone(timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": end.astimezone(timezone.utc).strftime(
                "%Y-%m-%dT%H:%M:%S"
            ),
            "timeZone": "UTC",
        },
        "categories": ["CS2"],
    }

    return payload


def normalize_graph_datetime(value):
    parsed = parse_datetime(value)
    return parsed.astimezone(timezone.utc).replace(
        microsecond=0
    )


def create_outlook_event(token, match):
    payload = build_event_payload(match)

    response = requests.post(
        f"{GRAPH_URL}/me/events",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if response.status_code != 201:
        print(
            f"ERRO ao criar {payload['subject']}: "
            f"{response.status_code} - {response.text}"
        )
        return False

    print(f"CRIADO: {payload['subject']}")
    return True


def update_outlook_event(token, match, existing_event):
    payload = build_event_payload(match)

    current_start = normalize_graph_datetime(
        existing_event["start"]["dateTime"]
    )
    current_end = normalize_graph_datetime(
        existing_event["end"]["dateTime"]
    )

    new_start = normalize_graph_datetime(
        payload["start"]["dateTime"]
    )
    new_end = normalize_graph_datetime(
        payload["end"]["dateTime"]
    )

    current_body = existing_event.get("body", {}).get("content", "")
    event_name = match.get("eventName") or "Campeonato de CS2"

    changed = (
        existing_event.get("subject") != payload["subject"]
        or current_start != new_start
        or current_end != new_end
        or event_name not in current_body
    )

    if not changed:
        print(f"SEM ALTERAÇÕES: {payload['subject']}")
        return False

    event_id = existing_event["id"]

    response = requests.patch(
        f"{GRAPH_URL}/me/events/{event_id}",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        json=payload,
        timeout=30,
    )

    if response.status_code != 200:
        print(
            f"ERRO ao atualizar {payload['subject']}: "
            f"{response.status_code} - {response.text}"
        )
        return False

    print(f"ATUALIZADO: {payload['subject']}")
    return True



def main():
    api_key = get_environment_variable("CITO_API_KEY")
    client_id = get_environment_variable("MS_CLIENT_ID")

    print("Buscando partidas na CitoAPI...")
    matches = get_upcoming_matches(api_key)

    print(f"Partidas encontradas de FURIA/Legacy: {len(matches)}")

    if not matches:
        print("Nenhuma partida nova para processar.")
        return

    # Consulta os eventos futuros, incluindo uma margem de duração.
    now = datetime.now(timezone.utc)
    latest_start = max(
        parse_datetime(match["startsAt"]) for match in matches
    )
    calendar_end = max(
        now + timedelta(days=30),
        latest_start + timedelta(hours=DEFAULT_DURATION_HOURS + 1),
    )

    token = authenticate_outlook(client_id)

    
    existing_events = get_existing_events(token, now, calendar_end)

    created = 0
    updated = 0
    unchanged = 0

    for match in matches:
        match_id = match.get("id")

        if not match_id:
            print("Partida ignorada: não tem ID.")
            continue

        existing_event = existing_events.get(match_id)

        if existing_event:
            was_updated = update_outlook_event(
                token,
                match,
                existing_event,
            )

            if was_updated:
                updated += 1
            else:
                unchanged += 1

        else:
            if create_outlook_event(token, match):
                created += 1

    print("\nSincronização concluída!")
    print(f"Eventos criados: {created}")
    print(f"Eventos atualizados: {updated}")
    print(f"Eventos sem alterações ou não atualizados: {unchanged}")



if __name__ == "__main__":
    try:
        main()
    except requests.RequestException as error:
        print(f"Erro de comunicação com uma API: {error}")
        sys.exit(1)
    except (RuntimeError, KeyError, ValueError) as error:
        print(f"Erro: {error}")
        sys.exit(1)
