from __future__ import annotations

from typing import Any

import pandas as pd
from jose import jwt
from jose.exceptions import JWTError

from lfx.components.microsoft.auth_helper import get_graph_token
from lfx.custom.custom_component.component import Component
from lfx.io import DropdownInput, MessageTextInput, SecretStrInput
from lfx.schema.dataframe import DataFrame
from lfx.template.field.base import Output


class MicrosoftGraphPermissionsComponent(Component):
    """Decode the Microsoft Graph access token and output its roles/scopes.

    For application tokens (client credentials), roles are in the `roles` claim.
    For delegated tokens, permissions are space-separated in the `scp` claim.
    """

    display_name: str = "Microsoft Graph Permissions"
    description: str = "Decodes the Graph access token and lists roles/scopes."
    icon: str = "Microsoft"
    # Minimal, commonly used Graph permission → representative REST endpoint mapping
    # Note: Some endpoints are contextual and may require identifiers (e.g., {id}).
    APP_ROLE_TO_ENDPOINT: dict[str, str] = {
        "User.Read.All": "/users",
        "User.ReadWrite.All": "/users",
        "Directory.Read.All": "/directoryObjects",
        "Directory.ReadWrite.All": "/directoryObjects",
        "Group.Read.All": "/groups",
        "Group.ReadWrite.All": "/groups",
        "Mail.Read": "/users/{id}/messages",
        "Mail.ReadBasic.All": "/users/{id}/messages",
        "Mail.ReadWrite": "/users/{id}/messages",
        "Calendars.Read": "/users/{id}/events",
        "Calendars.ReadWrite": "/users/{id}/events",
        "Files.Read.All": "/drives",
        "Files.ReadWrite.All": "/drives",
        "Sites.Read.All": "/sites",
        "Sites.ReadWrite.All": "/sites",
        "Reports.Read.All": "/reports",
        "Presence.Read.All": "/communications/presences",
        "Presence.ReadWrite.All": "/communications/presences",
        "Chat.Read.All": "/chats",
        "Channel.Read.All": "/teams/{id}/channels",
        "Team.ReadBasic.All": "/teams",
    }

    SCOPE_TO_ENDPOINT: dict[str, str] = {
        "User.Read": "/me",
        "User.ReadBasic.All": "/users",
        "User.Read.All": "/users",
        "User.ReadWrite": "/me",
        "User.ReadWrite.All": "/users",
        "Mail.Read": "/me/messages",
        "Mail.ReadWrite": "/me/messages",
        "Mail.Send": "/me/sendMail",
        "Mail.ReadBasic": "/me/messages",
        "Calendars.Read": "/me/events",
        "Calendars.ReadWrite": "/me/events",
        "Files.Read": "/me/drive",
        "Files.Read.All": "/drives",
        "Files.ReadWrite": "/me/drive",
        "Files.ReadWrite.All": "/drives",
        "Sites.Read.All": "/sites",
        "Channel.ReadBasic.All": "/teams",
        "Chat.Read": "/me/chats",
        "Presence.Read": "/me/presence",
    }

    inputs = [
        MessageTextInput(
            name="tenant_id",
            display_name="Tenant ID",
            info="Azure AD tenant (GUID or domain).",
            required=True,
        ),
        MessageTextInput(
            name="client_id",
            display_name="Application (Client) ID",
            required=True,
        ),
        SecretStrInput(
            name="client_secret",
            display_name="Client Secret",
            required=True,
        ),
        DropdownInput(
            name="graph_version",
            display_name="Graph Version",
            options=["v1.0", "beta"],
            value="v1.0",
        ),
    ]

    outputs = [
        Output(name="permissions", display_name="Permissions", method="list_token_roles"),
    ]

    def endpoint_for_permission(self, name: str, permission_type: str) -> str | None:
        if permission_type == "application":
            r = self.APP_ROLE_TO_ENDPOINT.get(name)
            self.log(r, "Roles")
            return self.APP_ROLE_TO_ENDPOINT.get(name)
        return self.SCOPE_TO_ENDPOINT.get(name)

    def list_token_roles(self) -> DataFrame:
        token = get_graph_token(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )
        try:
            access_token = token["access_token"]
            claims: dict[str, Any] = jwt.get_unverified_claims(access_token)

            # Prefer application roles; fall back to delegated scopes
            roles = claims.get("roles")
            permission_type = "application" if roles else "delegated"
            if not roles:
                scp = claims.get("scp")  # space-separated scopes for delegated tokens
                roles = scp.split(" ") if isinstance(scp, str) else []

            rows = []
            for name in roles if isinstance(roles, list) else []:
                endpoint = self.endpoint_for_permission(name, permission_type)
                rows.append(
                    {
                        "permission_type": permission_type,
                        "permission": name,
                        "endpoint": endpoint or "",
                    }
                )
            if not rows:
                rows = [{"info": "No roles/scopes present in token"}]

            return DataFrame(pd.DataFrame(rows))
        except JWTError as e:
            return DataFrame(pd.DataFrame({"error": [str(e)]}))
