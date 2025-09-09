from __future__ import annotations

from typing import Any

import pandas as pd
import requests

from lfx.components.microsoft.auth_helper import create_graph_client, get_graph_token
from lfx.custom.custom_component.component import Component
from lfx.io import DropdownInput, MessageTextInput, SecretStrInput
from lfx.schema.dataframe import DataFrame
from lfx.template.field.base import Output


class MicrosoftGraphComponent(Component):
    """Simple component to call a Microsoft Graph endpoint using app-only auth."""

    display_name: str = "Microsoft Graph"
    description: str = "Calls a Microsoft Graph endpoint using Client Credentials."
    icon: str = "Microsoft"

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
        MessageTextInput(
            name="user_id",
            display_name="User ID",
            info="the user id to get the details of",
            value="",
            required=True,
        ),
        MessageTextInput(
            name="endpoint_path",
            display_name="Endpoint Path",
            info="Relative path, e.g. /users?$top=5",
            value="/users?$top=5",
            tool_mode=True,
            required=True,
        ),
    ]

    outputs = [
        Output(name="response", display_name="Response", method="call_graph"),
    ]

    def call_graph(self) -> DataFrame:
        base_url = f"https://graph.microsoft.com/{self.graph_version}"
        token = get_graph_token(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )
        client = create_graph_client(token["access_token"], base_url=base_url)

        try:
            resp = client.get(self.endpoint_path)
            resp.raise_for_status()
            data: Any = resp.json()

            # Normalize typical Graph responses
            if isinstance(data, dict) and "value" in data and isinstance(data["value"], list):
                rows: list[dict[str, Any]] = data["value"]
            elif isinstance(data, list):
                rows = data
            else:
                rows = [data]

            return DataFrame(pd.DataFrame(rows))
        except requests.exceptions.RequestException as e:
            return DataFrame(pd.DataFrame({"error": [str(e)]}))
