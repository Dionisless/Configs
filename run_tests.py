from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import datetime, timezone


ARTIFACTS_DIR = 'artifacts'
os.makedirs(ARTIFACTS_DIR, exist_ok=True)

commands = [
    ['python', '-m', 'unittest', 'discover', '-s', 'tests', '-v'],
    ['python', 'digsi_readable.py', 'Примеры', '-o', os.path.join(ARTIFACTS_DIR, 'integration_out'), '--report'],
]

results = []
for cmd in commands:
    proc = subprocess.run(cmd, capture_output=True, text=True)
    name = '_'.join(cmd[:3]).replace('/', '_').replace(' ', '_')
    log_file = os.path.join(ARTIFACTS_DIR, f'{name}.log')
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write(proc.stdout)
        f.write('\n--- STDERR ---\n')
        f.write(proc.stderr)
    results.append({
        'command': cmd,
        'returncode': proc.returncode,
        'log': log_file,
    })

report_file = os.path.join(ARTIFACTS_DIR, 'test_report.json')
with open(report_file, 'w', encoding='utf-8') as f:
    json.dump({
        'generated_at_utc': datetime.now(timezone.utc).isoformat(),
        'results': results,
    }, f, ensure_ascii=False, indent=2)

failed = any(r['returncode'] != 0 for r in results)
print(f'Test report: {report_file}')
for r in results:
    print(f"{'OK' if r['returncode'] == 0 else 'FAIL'}: {' '.join(r['command'])} -> {r['log']}")

sys.exit(1 if failed else 0)
