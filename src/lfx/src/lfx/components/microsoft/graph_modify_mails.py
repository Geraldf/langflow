from __future__ import annotations

from typing import Any

import pandas as pd  # pyright: ignore[reportMissingImports]
import requests  # pyright: ignore[reportMissingModuleSource]

from lfx.components.microsoft.auth_helper import create_graph_client, get_graph_token, graph_patch
from lfx.custom.custom_component.component import Component
from lfx.io import DataInput, DropdownInput, MessageTextInput, SecretStrInput
from lfx.schema.dataframe import DataFrame
from lfx.template.field.base import Output


class MicrosoftGraphModifyMailsComponent(Component):
    """Simple component to call a Microsoft Graph endpoint using app-only auth."""

    display_name: str = "Modify Mails Category with Microsoft Graph"
    description: str = "Modify mails category from a Microsoft Graph endpoint using Client Credentials."
    icon: str = "mail-check"

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
            name="auure_user_id",
            display_name="User ID",
            info="the user id to get the details of",
            value="",
            required=True,
        ),
        DataInput(
            name="message_id",
            display_name="Message ID",
            info="The input Message ID to operate on.",
            required=True,
            input_types=["Data"],
            is_list=True,
        ),
        MessageTextInput(
            name="category",
            display_name="please list the category you want to modify",
            info="the category you want to modify can be a comma separated list",
            value="",
            required=True,
        ),
    ]

    outputs = [
        Output(name="mails", display_name="Mails", method="call_graph"),
    ]

    def call_graph(self) -> DataFrame:
        base_url = f"https://graph.microsoft.com/{self.graph_version}"
        token = get_graph_token(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )

        client = create_graph_client(token["access_token"], base_url=base_url)

        # Extract message ID from Data object
        # message_id = self.message_id.text if hasattr(self.message_id, "text") else str(self.message_id)
        message_id = self.message_id[0].data["id"]
        try:
            resp = graph_patch(
                client,
                f"/users/{self.auure_user_id}/messages/{message_id}",
                json={"categories": self.category.split(",") if "," in self.category else [self.category]},
            )
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
