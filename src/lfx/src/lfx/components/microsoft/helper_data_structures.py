# from lfx.field_typing import Data


from bs4 import BeautifulSoup  # pyright: ignore[reportMissingModuleSource]

from lfx.custom.custom_component.component import Component
from lfx.io import DataInput, Output
from lfx.schema.message import Message


class HelperOnDataStructures(Component):
    display_name = "Format E-Mails"
    description = "Preprocesses email JSON (subject + body[html]) into plain text."
    documentation: str = "https://docs.langflow.org/components-custom-components"
    icon = "code"
    name = "HelperOnDataStructures"

    inputs = [
        DataInput(name="data_value", display_name="Data", info="Data object to filter.", required=True, is_list=False),
    ]

    outputs = [
        Output(display_name="Output", name="output", method="build_output"),
    ]

    def html_to_plaintext(self, html_text):
        """Konvertiert HTML zu Plain Text mit Beautiful Soup."""
        soup = BeautifulSoup(html_text, "html.parser")
        return soup.get_text()

    def build_output(self) -> Message:
        d = self.data_value
        md = d.model_dump()
        body_list = md["data"]["results"]
        if len(body_list) > 1:
            self.log(f"can only handle one body, you provided {len(body_list)}")
        if body_list[0]["body"]["contentType"] == "html":
            body_text = self.html_to_plaintext(body_list[0]["body"]["content"])
        else:
            body_text = body_list[0]["body"]["content"]
        subject_text = body_list[0]["subject"]
        from_text = body_list[0]["from"]["emailAddress"]["address"]

        # hier bauen wir den Output-Text
        output_text = f"From: {from_text}\nSubject: {subject_text}\nBody_Text:{chr(10)}{body_text}"

        # korrektes Message-Objekt zurückgeben
        return Message(text=output_text)
