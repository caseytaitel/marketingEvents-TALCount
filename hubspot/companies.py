"""Primary Company (associations v4 typeId 1) and TAL company batch read."""

from __future__ import annotations

import requests

from hubspot.retry import request_json

_ASSOCIATIONS_URL = (
    "https://api.hubapi.com/crm/v4/associations/contacts/companies/batch/read"
)
_COMPANY_BATCH_URL = "https://api.hubapi.com/crm/v3/objects/companies/batch/read"
_PRIMARY_TYPE_ID = 1
_PRIMARY_CATEGORY = "HUBSPOT_DEFINED"
_ASSOC_BATCH = 100
_COMPANY_BATCH = 100


def primary_companies_for_contacts(
    session: requests.Session, contact_ids: list[str]
) -> dict[str, str]:
    """Map contact id → Primary Company id (`HUBSPOT_DEFINED`, `typeId == 1`).

    Contacts with no Primary Company are omitted (they count in D, not E).
    """
    primary: dict[str, str] = {}
    unique_ids = list(dict.fromkeys(str(cid) for cid in contact_ids))
    for start in range(0, len(unique_ids), _ASSOC_BATCH):
        chunk = unique_ids[start : start + _ASSOC_BATCH]
        payload = request_json(
            session,
            "POST",
            _ASSOCIATIONS_URL,
            json={"inputs": [{"id": cid} for cid in chunk]},
        )
        for row in payload.get("results") or []:
            from_id = str((row.get("from") or {}).get("id") or "")
            if not from_id:
                continue
            company_id = _primary_company_id(row.get("to") or [])
            if company_id:
                primary[from_id] = company_id
    return primary


def _primary_company_id(targets: list[dict]) -> str | None:
    for target in targets:
        types = target.get("associationTypes") or []
        if any(
            assoc.get("category") == _PRIMARY_CATEGORY
            and str(assoc.get("typeId")) == str(_PRIMARY_TYPE_ID)
            for assoc in types
        ):
            to_id = target.get("toObjectId")
            if to_id is not None:
                return str(to_id)
    return None


def tal_flags_for_companies(
    session: requests.Session, company_ids: list[str]
) -> dict[str, bool]:
    """Map company id → True iff `hs_is_target_account` is the string `"true"`."""
    flags: dict[str, bool] = {}
    unique_ids = list(dict.fromkeys(str(cid) for cid in company_ids))
    for start in range(0, len(unique_ids), _COMPANY_BATCH):
        chunk = unique_ids[start : start + _COMPANY_BATCH]
        payload = request_json(
            session,
            "POST",
            _COMPANY_BATCH_URL,
            json={
                "properties": ["hs_is_target_account"],
                "inputs": [{"id": cid} for cid in chunk],
            },
        )
        for row in payload.get("results") or []:
            company_id = str(row.get("id") or "")
            if not company_id:
                continue
            raw = (row.get("properties") or {}).get("hs_is_target_account")
            flags[company_id] = raw == "true"
    return flags
