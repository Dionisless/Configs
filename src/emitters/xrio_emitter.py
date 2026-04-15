from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from typing import Iterable

from src.model.relay_model import Parameter


def emit_xrio(parameters: Iterable[Parameter], out_path: str) -> None:
    root = ET.Element("XRio")
    custom = ET.SubElement(root, "CUSTOM")
    block = ET.SubElement(custom, "Block", {"Id": "ADD_PARAM"})

    for p in parameters:
        pnode = ET.SubElement(
            block,
            "Parameter",
            {
                "ForeignId": p.address,
                "Name": p.name_full or p.name_short,
                "Value": p.value_norm,
                "Datatype": p.datatype,
                "Unit": p.unit,
            },
        )
        if p.enum_map:
            enum_list = ET.SubElement(pnode, "EnumList")
            for code, label in sorted(p.enum_map.items(), key=lambda kv: kv[0]):
                ET.SubElement(enum_list, "Value", {"Code": f"TXT_{code}", "Name": str(label)})

    tree = ET.ElementTree(root)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    tree.write(out_path, encoding="utf-8", xml_declaration=True)
