"""Load HUBSPOT_TOKEN from .env.local and build an authenticated session."""

from __future__ import annotations

import os
from pathlib import Path

import requests
from dotenv import load_dotenv

_REPO_ROOT = Path(__file__).resolve().parent.parent
_ENV_LOCAL = _REPO_ROOT / ".env.local"


def load_token() -> str:
    """Load `.env.local` from the repo root and return `HUBSPOT_TOKEN`."""
    load_dotenv(_ENV_LOCAL)
    token = (os.environ.get("HUBSPOT_TOKEN") or "").strip()
    if not token:
        raise SystemExit(
            f"HUBSPOT_TOKEN is missing. Copy .env.example to {_ENV_LOCAL.name} and set it."
        )
    return token


def make_session(token: str) -> requests.Session:
    session = requests.Session()
    session.headers.update(
        {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }
    )
    return session
