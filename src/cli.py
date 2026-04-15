from __future__ import annotations

import argparse
import json
import os

from src.container_reader import extract_container
from src.emitters.xml_emitter import emit_xml
from src.emitters.xrio_emitter import emit_xrio
from src.model.normalize import normalize_value, normalize_datatype, normalize_address
from src.model.relay_model import Parameter, RelayModel
from src.prot_locator import discover_prot_par
from src.prot_parser.prot_decoder import decode_prot_file
from src.validators.crosscheck import compare_xml_xrio, CrosscheckResult
from src.validators.report import write_report


def _build_model(manifest) -> RelayModel:
    model = RelayModel()
    all_sources = manifest.prot_files + manifest.par_files

    for file_path in all_sources:
        decoded = decode_prot_file(file_path)
        for p in decoded:
            model.add(
                Parameter(
                    address=normalize_address(p.address),
                    name_short=p.name_short,
                    name_full=p.name_full,
                    value_raw=p.value_raw,
                    value_norm=normalize_value(p.value_norm),
                    datatype=normalize_datatype(p.datatype),
                    enum_map=p.enum_map,
                    unit=p.unit,
                    section_path=p.section_path,
                    source_file=p.source_file,
                )
            )

    dedup = RelayModel()
    seen = set()
    for p in model.parameters:
        if p.address and p.address not in seen:
            dedup.add(p)
            seen.add(p.address)
    return dedup


def run_pipeline(input_config: str, out_dir: str, with_report: bool = False) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    extraction = extract_container(input_config, out_root=out_dir)
    manifest = discover_prot_par(extraction.work_dir)

    model = _build_model(manifest)

    output_xml = os.path.join(out_dir, 'output.xml')
    output_xrio = os.path.join(out_dir, 'output.xrio')

    emit_xml(model.parameters, output_xml)
    emit_xrio(model.parameters, output_xrio)

    result: CrosscheckResult = compare_xml_xrio(output_xml, output_xrio)
    report_files = write_report(result, out_dir) if with_report else (None, None)

    manifest_path = os.path.join(out_dir, 'manifest.json')
    with open(manifest_path, 'w', encoding='utf-8') as f:
        json.dump(
            {
                'source_type': extraction.source_type,
                'root': extraction.work_dir,
                'prot_files': manifest.prot_files,
                'par_files': manifest.par_files,
                'parameter_count': len(model.parameters),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    return {
        'output_xml': output_xml,
        'output_xrio': output_xrio,
        'manifest': manifest_path,
        'report_json': report_files[0],
        'report_txt': report_files[1],
        'summary': result.summary(),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog='digsi_readable', description='DIGSI config -> XML/XRio pipeline')
    parser.add_argument('input_config', help='Path to config container/archive or single file')
    parser.add_argument('-o', '--out-dir', default='out', help='Output directory')
    parser.add_argument('--report', action='store_true', help='Write report.json/report.txt')
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    result = run_pipeline(args.input_config, args.out_dir, with_report=args.report)

    print('Pipeline completed')
    print(f"output.xml: {result['output_xml']}")
    print(f"output.xrio: {result['output_xrio']}")
    if result['report_txt']:
        print(f"report.txt: {result['report_txt']}")
    print(f"summary: {result['summary']}")


if __name__ == '__main__':
    main()
