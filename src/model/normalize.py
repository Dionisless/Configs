from __future__ import annotations

import re


def normalize_address(value: str) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    text = text.replace("0X", "")
    return text


def normalize_datatype(value: str) -> str:
    if not value:
        return "STRING"
    text = str(value).strip().upper()
    aliases = {
        "INT": "INTEGER",
        "ENUM": "ENUM",
        "FLOAT": "FLOAT",
        "REAL": "FLOAT",
        "BOOL": "BOOLEAN",
        "BOOLEAN": "BOOLEAN",
    }
    return aliases.get(text, text)


def normalize_value(value: str, trim_units: bool = True) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() == "oo":
        return "INF"

    if text.upper().startswith("TXT_"):
        text = text[4:]

    if trim_units:
        text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_pair_for_compare(value: str, datatype: str) -> tuple[str, str]:
    return normalize_value(value), normalize_datatype(datatype)
