"""
dok.xml_writer
~~~~~~~~~~~~~~
A tiny XML builder. No dependencies.
Produces well-formed XML strings into an in-memory buffer.

Used only by DocxWriter. Nothing else should touch raw XML.
"""

from __future__ import annotations
import io


def _escape(text: str) -> str:
    return (text
            .replace("&",  "&amp;")
            .replace("<",  "&lt;")
            .replace(">",  "&gt;")
            .replace('"',  "&quot;"))


def _attrs(d: dict) -> str:
    if not d:
        return ""
    return "".join(f' {k}="{_escape(str(v))}"' for k, v in d.items())


class XmlWriter:
    """
    Builds an XML document into a string buffer.

    Usage:
        w = XmlWriter()
        w.declaration()
        w.open("w:document", {"xmlns:w": "..."})
        w.open("w:body")
        w.tag("w:p")
        w.close("w:body")
        w.close("w:document")
        xml = w.getvalue()
    """

    def __init__(self) -> None:
        self._buf:   list[str] = []
        self._stack: list[str] = []

    def declaration(self) -> None:
        """Emit <?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
        self._buf.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')

    def open(self, tag: str, attrs: dict | None = None) -> None:
        """Open a tag and push it onto the stack."""
        self._buf.append(f"<{tag}{_attrs(attrs or {})}>")
        self._stack.append(tag)

    def close(self, tag: str | None = None) -> None:
        """Close the topmost open tag."""
        top = self._stack.pop()
        if tag and tag != top:
            raise ValueError(f"close({tag!r}) but top is {top!r}")
        self._buf.append(f"</{top}>")

    def tag(self, name: str, attrs: dict | None = None) -> None:
        """Emit a self-closing tag: <name attrs/>"""
        self._buf.append(f"<{name}{_attrs(attrs or {})}/>"  )

    def text(self, content: str) -> None:
        """Emit escaped text content."""
        self._buf.append(_escape(content))

    def raw(self, xml: str) -> None:
        """Emit raw XML (caller is responsible for correctness)."""
        self._buf.append(xml)

    def getvalue(self) -> str:
        if self._stack:
            raise RuntimeError(f"Unclosed tags: {self._stack}")
        return "".join(self._buf)

    # ------------------------------------------------------------------
    # Convenience: write a text run with preserved spaces
    # ------------------------------------------------------------------

    def w_t(self, text: str) -> None:
        """
        Emit <w:t> with xml:space="preserve" when needed.
        Always use this instead of raw tag() for text content.
        """
        needs_preserve = text.startswith(" ") or text.endswith(" ")
        attrs = {"xml:space": "preserve"} if needs_preserve else {}
        self.open("w:t", attrs)
        self.text(text)
        self.close("w:t")