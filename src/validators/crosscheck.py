from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass

from src.model.normalize import normalize_address, normalize_pair_for_compare


@dataclass
class CrosscheckResult:
    missing_in_xml: list[str]
    missing_in_xrio: list[str]
    value_mismatch: list[str]
    datatype_mismatch: list[str]

    def summary(self) -> dict:
        return {
            "missing_in_xml": len(self.missing_in_xml),
            "missing_in_xrio": len(self.missing_in_xrio),
            "value_mismatch": len(self.value_mismatch),
            "datatype_mismatch": len(self.datatype_mismatch),
        }


def _read_xml_params(xml_path: str) -> dict[str, tuple[str, str]]:
    tree = ET.parse(xml_path)
    root = tree.getroot()
    params = {}
    for p in root.findall('.//Parameter'):
        addr = normalize_address(p.attrib.get('DAdr') or p.attrib.get('ForeignId') or '')
        value = p.attrib.get('Value', '')
        dtype = p.attrib.get('Datatype', '')
        if addr:
            params[addr] = normalize_pair_for_compare(value, dtype)
    return params


def compare_xml_xrio(xml_path: str, xrio_path: str) -> CrosscheckResult:
    xml = _read_xml_params(xml_path)
    xrio = _read_xml_params(xrio_path)

    xml_set = set(xml)
    xrio_set = set(xrio)

    missing_in_xml = sorted(xrio_set - xml_set)
    missing_in_xrio = sorted(xml_set - xrio_set)
    value_mismatch = []
    datatype_mismatch = []

    for addr in sorted(xml_set & xrio_set):
        x_val, x_dtype = xml[addr]
        r_val, r_dtype = xrio[addr]
        if x_val != r_val:
            value_mismatch.append(addr)
        if x_dtype != r_dtype:
            datatype_mismatch.append(addr)

    return CrosscheckResult(
        missing_in_xml=missing_in_xml,
        missing_in_xrio=missing_in_xrio,
        value_mismatch=value_mismatch,
        datatype_mismatch=datatype_mismatch,
    )
