#!/usr/bin/env python3
"""Convert a macro-enabled deck to PPAM and add the PPT Tools Ribbon."""

from __future__ import annotations

import argparse
import os
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path


RELATIONSHIP_TYPE = (
    "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
)
RELATIONSHIP_TARGET = "customUI/customUI14.xml"
CONTENT_TYPES_PATH = "[Content_Types].xml"
ADDIN_CONTENT_TYPE = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"


def inject_custom_ui(ppam_path: Path, custom_ui_path: Path) -> None:
    with zipfile.ZipFile(ppam_path, "r") as source:
        root_relationships = ET.fromstring(source.read("_rels/.rels"))
        namespace = root_relationships.tag.split("}", 1)[0].removeprefix("{")
        content_types = ET.fromstring(source.read(CONTENT_TYPES_PATH))
        content_namespace = content_types.tag.split("}", 1)[0].removeprefix("{")

        for override in content_types:
            if override.get("PartName") == "/ppt/presentation.xml":
                override.set("ContentType", ADDIN_CONTENT_TYPE)
                break
        else:
            raise ValueError("The package has no /ppt/presentation.xml content type")

        for relationship in root_relationships:
            if relationship.get("Type") == RELATIONSHIP_TYPE:
                relationship.set("Target", RELATIONSHIP_TARGET)
                break
        else:
            used_ids = {item.get("Id") for item in root_relationships}
            next_id = 1
            while f"rId{next_id}" in used_ids:
                next_id += 1
            ET.SubElement(
                root_relationships,
                f"{{{namespace}}}Relationship",
                {
                    "Id": f"rId{next_id}",
                    "Type": RELATIONSHIP_TYPE,
                    "Target": RELATIONSHIP_TARGET,
                },
            )

        ET.register_namespace("", namespace)
        relationships_xml = ET.tostring(
            root_relationships,
            encoding="utf-8",
            xml_declaration=True,
        )
        ET.register_namespace("", content_namespace)
        content_types_xml = ET.tostring(
            content_types,
            encoding="utf-8",
            xml_declaration=True,
        )
        custom_ui_xml = custom_ui_path.read_bytes()

        with tempfile.NamedTemporaryFile(
            prefix=ppam_path.stem + "-",
            suffix=".ppam",
            dir=ppam_path.parent,
            delete=False,
        ) as temporary_file:
            temporary_path = Path(temporary_file.name)

        try:
            with zipfile.ZipFile(
                temporary_path,
                "w",
                compression=zipfile.ZIP_DEFLATED,
            ) as destination:
                for item in source.infolist():
                    if item.filename in {
                        CONTENT_TYPES_PATH,
                        "_rels/.rels",
                        RELATIONSHIP_TARGET,
                    }:
                        continue
                    destination.writestr(item, source.read(item.filename))

                destination.writestr(CONTENT_TYPES_PATH, content_types_xml)
                destination.writestr("_rels/.rels", relationships_xml)
                destination.writestr(RELATIONSHIP_TARGET, custom_ui_xml)

            os.replace(temporary_path, ppam_path)
        finally:
            temporary_path.unlink(missing_ok=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("ppam", type=Path)
    parser.add_argument("custom_ui", type=Path)
    args = parser.parse_args()

    inject_custom_ui(args.ppam.resolve(), args.custom_ui.resolve())


if __name__ == "__main__":
    main()
