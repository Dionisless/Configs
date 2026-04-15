from __future__ import annotations

import os
import re
from typing import List

from src.model.normalize import normalize_address, normalize_datatype, normalize_value
from src.model.relay_model import Parameter

ADDRESS_RE = re.compile(rb"(?:0x)?([0-9A-Fa-f]{4,8})")
TEXT_RE = re.compile(rb"[ -~\xd0-\xff]{4,}")


def _guess_datatype(raw_value: str) -> str:
    if raw_value.isdigit() or re.fullmatch(r"-?\d+", raw_value):
        return "INTEGER"
    if re.fullmatch(r"-?\d+[.,]\d+", raw_value):
        return "FLOAT"
    if raw_value.lower() in {"true", "false", "0", "1"}:
        return "BOOLEAN"
    return "STRING"


def decode_prot_bytes(blob: bytes, source_file: str = "") -> List[Parameter]:
    addresses = [normalize_address(m.group(1).decode("ascii", errors="ignore")) for m in ADDRESS_RE.finditer(blob)]
    texts = [m.group(0).decode("latin1", errors="ignore").strip("\x00 ") for m in TEXT_RE.finditer(blob)]

    params: List[Parameter] = []
    for idx, address in enumerate(addresses[:1000]):
        if not address:
            continue
        name = texts[idx] if idx < len(texts) else f"Param_{address}"
        raw = texts[idx + 1] if idx + 1 < len(texts) else "0"
        datatype = normalize_datatype(_guess_datatype(raw))
        params.append(
            Parameter(
                address=address,
                name_short=name[:40],
                name_full=name,
                value_raw=raw,
                value_norm=normalize_value(raw),
                datatype=datatype,
                enum_map={},
                unit="",
                section_path="auto/discovered",
                source_file=os.path.basename(source_file),
            )
        )

    unique = {}
    for p in params:
        unique[p.address] = p
    return list(unique.values())


def decode_prot_file(path: str) -> List[Parameter]:
    with open(path, "rb") as f:
        return decode_prot_bytes(f.read(), source_file=path)
