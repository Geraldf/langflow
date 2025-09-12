from __future__ import annotations

from typing import Any

import pandas as pd  # pyright: ignore[reportMissingImports]
import requests  # pyright: ignore[reportMissingModuleSource]

from lfx.components.microsoft.auth_helper import create_graph_client, get_graph_token
from lfx.custom.custom_component.component import Component
from lfx.io import DropdownInput, MessageTextInput, SecretStrInput
from lfx.schema.dataframe import DataFrame
from lfx.schema.message import Message
from lfx.template.field.base import Output


class MicrosoftGraphUserSelectorComponent(Component):
    """Fetch Microsoft Graph users and allow selection from a dropdown.

    Retrieves users from the tenant and provides a selection interface.
    Returns the selected user's details as a DataFrame.
    """

    display_name: str = "Get User ID with Microsoft Graph "
    description: str = "Fetch Graph users and select one from a dropdown."
    icon: str = "user-check"

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
            name="user_filter",
            display_name="Enter the users e-mail",
            info="The users e-mail to search for.",
            value="",
            required=True,
        ),
    ]

    outputs = [
        Output(name="selected_user", display_name="Selected User", method="get_selected_user"),
    ]

    # def _pre_run_setup(self):
    #     if not hasattr(self, "_initialized"):
    #         self._initialized = True
    #         self.all_users = self.get_all_users()
    #         self.log(f"All users: {self.all_users}")
    #         self.dropdown_options = self.all_users

    #         # Populate the 'selected_user_id' input options
    #         try:
    #             selected_user_input = next(i for i in self.inputs if getattr(i, "name", "") == "selected_user_id")
    #             selected_user_input.options = self.all_users if isinstance(self.all_users, list) else ["none"]
    #             if selected_user_input.options and selected_user_input.options[0] != "none":
    #                 selected_user_input.value = selected_user_input.options[0]
    #         except StopIteration:
    #             pass

    # def update_build_config(self, build_config: dict, field_value: str, field_name: str | None = None) -> dict:
    #     if field_name == "selected_user_id":
    #         self.all_users = self.get_all_users()
    #         if field_value == "none":
    #             build_config["selected_user_id"]["options"] = self.all_users
    #     return build_config

    def get_all_users(self) -> DataFrame:
        """Fetch all users from Microsoft Graph."""
        base_url = f"https://graph.microsoft.com/{self.graph_version}"
        token = get_graph_token(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )
        client = create_graph_client(token["access_token"], base_url=base_url)

        try:
            # Build the query
            query_params = (
                f"$top={self.max_users}&"
                f"$select=id,displayName,userPrincipalName,mail,givenName,surname,jobTitle,department"
            )

            self.log(f"Query params: {query_params}")
            resp = client.get(f"/users?{query_params}")
            resp.raise_for_status()
            data: Any = resp.json()

            users = data.get("value", []) if isinstance(data, dict) else []
            self.log(f"Users: {users}")

            # Update the dropdown options
            user_options = []
            for user in users:
                mail = user.get("mail", "")
                user_options.append(f"{mail}")

            # Update the dropdown input options
            self.inputs[-1].options = user_options  # Update the selected_user_id dropdown
        except requests.exceptions.RequestException as e:
            return DataFrame(pd.DataFrame({"error": [str(e)]}))
        else:
            return user_options

    def get_selected_user(self) -> Message:
        """Fetch all users from Microsoft Graph."""
        base_url = f"https://graph.microsoft.com/{self.graph_version}"
        token = get_graph_token(
            tenant_id=self.tenant_id,
            client_id=self.client_id,
            client_secret=self.client_secret,
        )
        client = create_graph_client(token["access_token"], base_url=base_url)

        try:
            # Build the query
            query_params = f"$filter=mail eq '{self.user_filter}'"
            resp = client.get(f"/users?{query_params}")
            resp.raise_for_status()
            data: Any = resp.json()

            users = data.get("value", []) if isinstance(data, dict) else []
            self.log(f"Users: {users}")
        except requests.exceptions.RequestException as e:
            return DataFrame(pd.DataFrame({"error": [str(e)]}))
        else:
            return Message(text=users[0]["id"])
