from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from typing import Iterable

from src.model.relay_model import Parameter


def emit_xml(parameters: Iterable[Parameter], out_path: str) -> None:
    root = ET.Element("DeviceData")
    settings = ET.SubElement(root, "Settings")
    function_group = ET.SubElement(settings, "FunctionGroup", {"Name": "AutoImported"})
    page = ET.SubElement(function_group, "SettingPage", {"Name": "Parameters"})

    for p in parameters:
        attrs = {
            "DAdr": p.address,
            "Name": p.name_full or p.name_short,
            "Value": p.value_norm,
            "Datatype": p.datatype,
            "Unit": p.unit,
            "SourceFile": p.source_file,
        }
        node = ET.SubElement(page, "Parameter", attrs)
        if p.enum_map:
            values = ET.SubElement(node, "Value")
            for code, label in sorted(p.enum_map.items(), key=lambda kv: kv[0]):
                ET.SubElement(values, "Comment", {"Number": str(code), "Name": str(label)})

    tree = ET.ElementTree(root)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
