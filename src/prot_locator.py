from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import List


@dataclass
class InputManifest:
    root: str
    prot_files: List[str] = field(default_factory=list)
    par_files: List[str] = field(default_factory=list)


def discover_prot_par(root_dir: str) -> InputManifest:
    manifest = InputManifest(root=root_dir)
    for cur, _, files in os.walk(root_dir):
        for name in files:
            upper = name.upper()
            full = os.path.join(cur, name)
            if upper.startswith("PROT"):
                manifest.prot_files.append(full)
            elif upper.startswith("PAR"):
                manifest.par_files.append(full)

    manifest.prot_files.sort()
    manifest.par_files.sort()
    return manifest
