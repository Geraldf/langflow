from __future__ import annotations

import time
from typing import Any

import httpx
from msal import ConfidentialClientApplication

GRAPH_DEFAULT_SCOPE = "https://graph.microsoft.com/.default"


class GraphAuthError(RuntimeError):
    pass


def get_graph_token(
    tenant_id: str,
    client_id: str,
    client_secret: str,
    scope: str = GRAPH_DEFAULT_SCOPE,
) -> dict[str, str]:
    """Acquire an application token for Microsoft Graph using client credentials.

    Returns a dict with keys: access_token, expires_at, token_type
    """
    if not all([tenant_id, client_id, client_secret]):
        message = "Missing tenant_id, client_id, or client_secret"
        raise GraphAuthError(message)

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    result = app.acquire_token_for_client(scopes=[scope])
    if "access_token" not in result:
        error_detail: Any = result.get("error_description") or result
        message = f"Failed to acquire token: {error_detail}"
        raise GraphAuthError(message)

    # Normalize output
    expires_in = int(result.get("expires_in", 0))
    now = int(time.time())
    return {
        "access_token": result["access_token"],
        "token_type": result.get("token_type", "Bearer"),
        "expires_at": str(now + expires_in) if expires_in else "0",
    }


def create_graph_client(
    access_token: str,
    base_url: str = "https://graph.microsoft.com/v1.0",
    timeout_seconds: float = 30.0,
) -> httpx.Client:
    """Create a preconfigured httpx Client for Microsoft Graph using the token."""
    if not access_token:
        message = "access_token is required to create Graph client"
        raise GraphAuthError(message)

    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json",
        "Content-Type": "application/json",
    }
    return httpx.Client(base_url=base_url, headers=headers, timeout=timeout_seconds)
