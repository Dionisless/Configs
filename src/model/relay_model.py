from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List


@dataclass
class Parameter:
    address: str
    name_short: str
    name_full: str
    value_raw: str
    value_norm: str
    datatype: str
    enum_map: Dict[str, str] = field(default_factory=dict)
    unit: str = ""
    section_path: str = ""
    source_file: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class RelayModel:
    parameters: List[Parameter] = field(default_factory=list)

    def add(self, parameter: Parameter) -> None:
        self.parameters.append(parameter)

    def by_address(self) -> Dict[str, Parameter]:
        out: Dict[str, Parameter] = {}
        for p in self.parameters:
            out[p.address] = p
        return out
