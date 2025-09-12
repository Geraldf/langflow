from __future__ import annotations

from typing import Any

import pandas as pd  # pyright: ignore[reportMissingImports]
import requests  # pyright: ignore[reportMissingModuleSource]

from lfx.components.microsoft.auth_helper import create_graph_client, get_graph_token
from lfx.custom.custom_component.component import Component
from lfx.inputs import IntInput, SortableListInput
from lfx.io import DropdownInput, MessageTextInput, SecretStrInput
from lfx.schema.dataframe import DataFrame
from lfx.template.field.base import Output


class MicrosoftGraphReadMailsComponent(Component):
    """Simple component to call a Microsoft Graph endpoint using app-only auth."""

    display_name: str = "Read Mails with Microsoft Graph"
    description: str = "Reads mails from a Microsoft Graph endpoint using Client Credentials."
    icon: str = "mails"

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
        SortableListInput(
            name="read_from_folder",
            display_name="Read From Folder",
            placeholder="Select Folder",
            options=[
                {"name": "Inbox", "icon": "plus"},
                {"name": "Outbox", "icon": "minus"},
                {"name": "Drafts", "icon": "filter"},
                {"name": "SentItems", "icon": "arrow-up"},
                {"name": "Trash", "icon": "pencil"},
                {"name": "JunkEmail", "icon": "replace"},
            ],
            limit=1,
        ),
        IntInput(
            name="max_results",
            display_name="Max Results",
            info="Maximum number of emails to load.",
            required=True,
            value=100,
        ),
        MessageTextInput(
            name="auure_user_id",
            display_name="User ID",
            info="the user id to get the details of",
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

        try:
            self.log(f"Reading mails for user2: {self.read_from_folder[0]['name']}")
            resp = client.get(
                f"/users/{self.auure_user_id}/mailFolders/{self.read_from_folder[0]['name']}/messages?$top={self.max_results}"
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
