"""Shared low-level operations for generated Word documents."""

from __future__ import annotations

from io import BytesIO
from zipfile import ZIP_DEFLATED, ZipFile

import lxml.etree as etree

_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def strip_comments(docx_bytes: bytes) -> bytes:
    """Remove all Word comment parts, references and relationships."""

    output = BytesIO()
    parser = etree.XMLParser(resolve_entities=False, no_network=True)
    with ZipFile(BytesIO(docx_bytes), "r") as source, ZipFile(
        output, "w", ZIP_DEFLATED
    ) as target:
        for item in source.infolist():
            name = item.filename
            if (
                name.startswith("word/comments")
                or name.startswith("word/_rels/comments")
                or name in {"word/people.xml", "word/_rels/people.xml.rels"}
            ):
                continue
            data = source.read(name)
            if name == "word/document.xml":
                root = etree.fromstring(data, parser)
                namespaces = {"w": _WORD_NS}
                for tag in ("commentRangeStart", "commentRangeEnd"):
                    for node in root.xpath(f".//w:{tag}", namespaces=namespaces):
                        node.getparent().remove(node)
                for node in root.xpath(".//w:commentReference", namespaces=namespaces):
                    run = node.getparent()
                    if run is not None and run.getparent() is not None:
                        run.getparent().remove(run)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            elif name.endswith(".rels") or name == "[Content_Types].xml":
                root = etree.fromstring(data, parser)
                for node in list(root):
                    relation_type = node.get("Type", "").casefold()
                    part_name = node.get("PartName", "").casefold()
                    if any(
                        marker in relation_type or marker in part_name
                        for marker in ("comments", "people")
                    ):
                        root.remove(node)
                data = etree.tostring(
                    root, xml_declaration=True, encoding="UTF-8", standalone=True
                )
            target.writestr(item, data)
    return output.getvalue()
