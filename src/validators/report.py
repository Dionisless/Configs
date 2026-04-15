from __future__ import annotations

import json
import os
from datetime import datetime, timezone

from src.validators.crosscheck import CrosscheckResult


def write_report(result: CrosscheckResult, out_dir: str) -> tuple[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    json_path = os.path.join(out_dir, 'report.json')
    txt_path = os.path.join(out_dir, 'report.txt')

    payload = {
        'generated_at_utc': datetime.now(timezone.utc).isoformat(),
        'missing_in_xml': result.missing_in_xml,
        'missing_in_xrio': result.missing_in_xrio,
        'value_mismatch': result.value_mismatch,
        'datatype_mismatch': result.datatype_mismatch,
        'summary': result.summary(),
    }

    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write('DIGSI readable validation report\n')
        f.write('=' * 40 + '\n')
        for key, value in payload['summary'].items():
            f.write(f'{key}: {value}\n')

    return json_path, txt_path
