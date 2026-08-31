"""Contact search on `events_attended` (CONTAINS_TOKEN prefilter + token split)."""

from __future__ import annotations

import re

import requests

from hubspot.retry import request_json

_SEARCH_URL = "https://api.hubapi.com/crm/v3/objects/contacts/search"
_FILTER_GROUP_LIMIT = 5
_PAGE_SIZE = 100
_WS = re.compile(r"\s+")


def normalize_event_token(value: str) -> str:
    """Strip ends and collapse internal whitespace — both sides of a match use this."""
    return _WS.sub(" ", (value or "").strip())


def tokens_for_contact(events_attended: str | None) -> set[str]:
    """Split HubSpot `events_attended` on `;` and normalize each token."""
    if not events_attended:
        return set()
    return {
        normalize_event_token(part)
        for part in events_attended.split(";")
        if part.strip()
    }


def search_contacts_by_events_attended(
    session: requests.Session, appendages: list[str]
) -> list[dict]:
    """Prefilter contacts whose `events_attended` may include any appendage.

    HubSpot `CONTAINS_TOKEN` is a prefilter only. Callers must confirm with
    `tokens_for_contact` / `normalize_event_token` equality.
    """
    unique: list[str] = []
    seen: set[str] = set()
    for appendage in appendages:
        token = normalize_event_token(appendage)
        if token and token not in seen:
            seen.add(token)
            unique.append(token)

    by_id: dict[str, dict] = {}
    for start in range(0, len(unique), _FILTER_GROUP_LIMIT):
        chunk = unique[start : start + _FILTER_GROUP_LIMIT]
        for contact in _search_chunk(session, chunk):
            by_id[str(contact["id"])] = contact
    return list(by_id.values())


def _search_chunk(session: requests.Session, tokens: list[str]) -> list[dict]:
    results: list[dict] = []
    after: str | None = None
    while True:
        body: dict = {
            "filterGroups": [
                {
                    "filters": [
                        {
                            "propertyName": "events_attended",
                            "operator": "CONTAINS_TOKEN",
                            "value": token,
                        }
                    ]
                }
                for token in tokens
            ],
            "properties": ["events_attended"],
            "limit": _PAGE_SIZE,
        }
        if after:
            body["after"] = after
        payload = request_json(session, "POST", _SEARCH_URL, json=body)
        results.extend(payload.get("results") or [])
        after = (payload.get("paging") or {}).get("next", {}).get("after")
        if not after:
            break
    return results
