from __future__ import annotations

import os
import shutil
import tempfile
import zipfile
from dataclasses import dataclass


@dataclass
class ExtractionResult:
    work_dir: str
    source_type: str


def _extract_zip(input_path: str, target_dir: str) -> bool:
    if not zipfile.is_zipfile(input_path):
        return False
    with zipfile.ZipFile(input_path) as zf:
        zf.extractall(target_dir)
    return True


def _extract_rar(input_path: str, target_dir: str) -> bool:
    try:
        import rarfile  # type: ignore
    except Exception:
        return False

    try:
        with rarfile.RarFile(input_path) as rf:
            rf.extractall(target_dir)
        return True
    except Exception:
        return False


def extract_container(input_path: str, out_root: str | None = None) -> ExtractionResult:
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)

    root = out_root or tempfile.mkdtemp(prefix="digsi_readable_")
    work_dir = os.path.join(root, "input")
    os.makedirs(work_dir, exist_ok=True)

    if os.path.isdir(input_path):
        input_abs = os.path.abspath(input_path)
        root_abs = os.path.abspath(root)
        work_abs = os.path.abspath(work_dir)
        for name in os.listdir(input_path):
            src = os.path.join(input_path, name)
            src_abs = os.path.abspath(src)
            if src_abs == root_abs or src_abs.startswith(root_abs + os.sep):
                continue
            if src_abs == work_abs or src_abs.startswith(work_abs + os.sep):
                continue
            dst = os.path.join(work_dir, name)
            if os.path.isdir(src):
                shutil.copytree(src, dst, dirs_exist_ok=True)
            else:
                shutil.copy2(src, dst)
        return ExtractionResult(work_dir=work_dir, source_type="directory")

    if _extract_zip(input_path, work_dir):
        return ExtractionResult(work_dir=work_dir, source_type="zip")

    if _extract_rar(input_path, work_dir):
        return ExtractionResult(work_dir=work_dir, source_type="rar")

    shutil.copy2(input_path, os.path.join(work_dir, os.path.basename(input_path)))
    return ExtractionResult(work_dir=work_dir, source_type="single-file")
