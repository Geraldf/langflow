# from lfx.field_typing import Data
import json
import re
from html import unescape

from lfx.custom.custom_component.component import Component
from lfx.io import DataInput, MessageTextInput, Output
from lfx.schema.message import Message


class HelperOnDataStructures(Component):
    display_name = "Format E-Mails"
    description = "Preprocesses email JSON (subject + body[html]) into plain text."
    documentation: str = "https://docs.langflow.org/components-custom-components"
    icon = "code"
    name = "HelperOnDataStructures"

    inputs = [
        MessageTextInput(
            name="input_value",
            display_name="Input Value",
            info="This is a custom component Input",
            value="Hello, World!",
            tool_mode=True,
        ),
        DataInput(name="data", display_name="Data", info="Data object to filter.", required=True, is_list=True),
    ]

    outputs = [
        Output(display_name="Output", name="output", method="build_output"),
    ]

    def html_to_text(self, html: str) -> str:
        if not html:
            return ""
        # Scripts/Styles entfernen
        html = re.sub("(?is)<script.*?>.*?</script>", "", html)
        html = re.sub("(?is)<style.*?>.*?</style>", "", html)
        # <br> / </p> in Zeilenumbrüche umwandeln
        html = html.replace("<br>", chr(10)).replace("<br/>", chr(10)).replace("<br />", chr(10))
        html = html.replace("</p>", chr(10) * 2).replace("</P>", chr(10) * 2)
        # Alle übrigen Tags entfernen
        text = re.sub("(?s)<.*?>", "", html)
        text = unescape(text)
        # Mehrfach-Spaces reduzieren
        text = re.sub(" +", " ", text)
        # 3+ Leerzeilen auf 2 reduzieren
        return re.sub("(" + chr(10) + "){3,}", chr(10) * 2, text).strip()

    def _to_string(self, maybe_message):
        try:
            return maybe_message.to_string()
        except AttributeError:
            try:
                return maybe_message.text
            except AttributeError:
                return maybe_message if isinstance(maybe_message, str) else str(maybe_message)

    def build_output(self) -> Message:
        raw = self._to_string(self.input_value)  # JSON-String vom TextInput
        try:
            obj = json.loads(raw)
        except json.JSONDecodeError:
            obj = {"subject": "", "body": {"contentType": "text", "content": raw}}

        subject = obj.get("subject", "")
        body = obj.get("body") or {}
        ct = (body.get("contentType") or "").lower()
        content = body.get("content") or ""
        body_text = (
            self.html_to_text(content)
            if "html" in ct
            else (content.strip() if isinstance(content, str) else str(content))
        )

        # hier bauen wir den Output-Text
        output_text = f"Subject: {subject}{chr(10)}{chr(10)}Body_Text:{chr(10)}{body_text}"

        # korrektes Message-Objekt zurückgeben
        return Message(text=output_text)
