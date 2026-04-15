from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import zipfile
from datetime import datetime, timezone


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def copy_if_exists(src: str, dst: str) -> bool:
    if not os.path.exists(src):
        return False
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    if os.path.isdir(src):
        shutil.copytree(src, dst, dirs_exist_ok=True)
    else:
        shutil.copy2(src, dst)
    return True


def create_portable_bundle(dist_dir: str, output_dir: str, zip_name: str = 'digsi_readable-portable.zip') -> dict:
    bundle_root = os.path.join(output_dir, 'digsi_readable_portable')
    if os.path.exists(bundle_root):
        shutil.rmtree(bundle_root)
    os.makedirs(bundle_root, exist_ok=True)

    included = []

    candidates = [
        (os.path.join(dist_dir, 'digsi_readable.exe'), 'digsi_readable.exe'),
        (os.path.join(dist_dir, 'digsi_readable'), 'digsi_readable'),
        (os.path.join(dist_dir, 'pipeline.exe'), 'pipeline.exe'),
        ('compare_tool.exe', 'compare_tool.exe'),
    ]

    for src_rel, dst_name in candidates:
        src = src_rel if os.path.isabs(src_rel) else os.path.join(os.getcwd(), src_rel)
        dst = os.path.join(bundle_root, dst_name)
        if copy_if_exists(src, dst):
            included.append(dst_name)

    if copy_if_exists(os.path.join(os.getcwd(), 'SiemensPie'), os.path.join(bundle_root, 'SiemensPie')):
        included.append('SiemensPie/')

    if 'digsi_readable.exe' not in included and 'digsi_readable' not in included:
        raise FileNotFoundError('digsi_readable executable not found in dist. Build it first via PyInstaller.')

    launch_bat = os.path.join(bundle_root, 'run_digsi_readable.bat')
    with open(launch_bat, 'w', encoding='utf-8') as f:
        f.write('@echo off\n')
        f.write('setlocal\n')
        f.write('cd /d %~dp0\n')
        f.write('if "%~1"=="" (\n')
        f.write('  echo Usage: run_digsi_readable.bat ^<input_config^> [-o out_dir] [--report]\n')
        f.write('  exit /b 1\n')
        f.write(')\n')
        f.write('if exist digsi_readable.exe (\n')
        f.write('  digsi_readable.exe %*\n')
        f.write(') else (\n')
        f.write('  echo digsi_readable.exe not found in portable folder\n')
        f.write('  exit /b 1\n')
        f.write(')\n')
    included.append('run_digsi_readable.bat')

    readme = os.path.join(bundle_root, 'PORTABLE_README.txt')
    with open(readme, 'w', encoding='utf-8') as f:
        f.write('DIGSI Readable Portable Bundle\n')
        f.write('================================\n')
        f.write('1) Скопируйте папку целиком на ПК Windows.\n')
        f.write('2) Запускайте run_digsi_readable.bat <input_config> -o out --report\n')
        f.write('3) Интернет и Python не требуются, если в папке есть digsi_readable.exe.\n')
        f.write('4) Для SiemensPie сценариев проверьте наличие папки SiemensPie рядом с exe.\n')
    included.append('PORTABLE_README.txt')

    manifest = {
        'generated_at_utc': datetime.now(timezone.utc).isoformat(),
        'included': included,
        'checksums': {},
    }

    for root, _, files in os.walk(bundle_root):
        for name in files:
            full = os.path.join(root, name)
            rel = os.path.relpath(full, bundle_root)
            manifest['checksums'][rel] = sha256_file(full)

    manifest_path = os.path.join(bundle_root, 'portable_manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    zip_path = os.path.join(output_dir, zip_name)
    with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
        for root, _, files in os.walk(bundle_root):
            for name in files:
                full = os.path.join(root, name)
                rel = os.path.relpath(full, output_dir)
                zf.write(full, rel)

    return {
        'bundle_root': bundle_root,
        'zip_path': zip_path,
        'manifest': manifest_path,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description='Build portable DIGSI bundle')
    parser.add_argument('--dist-dir', default='dist')
    parser.add_argument('--output-dir', default='portable')
    parser.add_argument('--zip-name', default='digsi_readable-portable.zip')
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    result = create_portable_bundle(args.dist_dir, args.output_dir, args.zip_name)
    print('Portable bundle created:')
    print(f"  folder:   {result['bundle_root']}")
    print(f"  zip:      {result['zip_path']}")
    print(f"  manifest: {result['manifest']}")


if __name__ == '__main__':
    main()
