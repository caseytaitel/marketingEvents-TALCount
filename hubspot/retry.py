"""HTTP helper with retries for HubSpot 429 and 5xx responses."""

from __future__ import annotations

import time
from typing import Any

import requests

_MAX_ATTEMPTS = 6
_BACKOFF_SECONDS = (1, 2, 4, 8, 16)
_TIMEOUT = 60


def request_json(
    session: requests.Session,
    method: str,
    url: str,
    *,
    json: dict[str, Any] | None = None,
) -> dict[str, Any]:
    last_error: Exception | None = None
    for attempt in range(_MAX_ATTEMPTS):
        try:
            response = session.request(
                method, url, json=json, timeout=_TIMEOUT
            )
        except requests.RequestException as exc:
            last_error = exc
            if attempt == _MAX_ATTEMPTS - 1:
                raise
            time.sleep(_BACKOFF_SECONDS[min(attempt, len(_BACKOFF_SECONDS) - 1)])
            continue

        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After")
            try:
                wait = float(retry_after) if retry_after else None
            except ValueError:
                wait = None
            if wait is None:
                wait = _BACKOFF_SECONDS[min(attempt, len(_BACKOFF_SECONDS) - 1)]
            time.sleep(wait)
            continue

        if response.status_code >= 500:
            if attempt == _MAX_ATTEMPTS - 1:
                raise RuntimeError(
                    f"HubSpot {response.status_code} {method} {url}: {response.text[:800]}"
                )
            time.sleep(_BACKOFF_SECONDS[min(attempt, len(_BACKOFF_SECONDS) - 1)])
            continue

        if not response.ok:
            raise RuntimeError(
                f"HubSpot {response.status_code} {method} {url}: {response.text[:800]}"
            )

        if not response.content:
            return {}
        payload = response.json()
        if not isinstance(payload, dict):
            raise RuntimeError(f"HubSpot returned a non-object JSON body from {url}")
        return payload

    raise RuntimeError(f"HubSpot request failed after retries: {method} {url}") from last_error
