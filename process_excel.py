from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
EXCEL_PATH = BASE_DIR / 'data' / 'REGISTRO_FINAL_CURSOS.xlsx'
OUTPUT_PATH = BASE_DIR / 'data' / 'dashboard_data.js'


def to_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def main() -> None:
    wb = load_workbook(EXCEL_PATH, data_only=True)

    summary_rows: list[dict[str, Any]] = []
    summary_map: dict[str, dict[str, Any]] = {}

    if 'RESUMEN' in wb.sheetnames:
        ws = wb['RESUMEN']
        rows = list(ws.iter_rows(values_only=True))
        headers = [str(h).strip() if h is not None else '' for h in rows[0]]

        for row in rows[1:]:
            if not any(cell is not None for cell in row):
                continue
            item = {headers[i]: row[i] for i in range(min(len(headers), len(row)))}
            curso = str(item.get('CURSO', '')).strip()
            total = to_int(item.get('TOTAL_ARCHIVOS'))
            unicos_raw = item.get('UNICOS')
            unicos = None if unicos_raw in (None, '') else to_int(unicos_raw)
            parsed = {'curso': curso, 'total': total, 'unicos': unicos}
            summary_rows.append(parsed)
            summary_map[curso] = parsed

    course_totals: list[dict[str, Any]] = []
    suffix_yes = 0
    suffix_no = 0

    for sheet_name in wb.sheetnames:
        if sheet_name == 'RESUMEN':
            continue

        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            continue

        headers = [str(h).strip() if h is not None else '' for h in rows[0]]
        data_rows = [
            row for row in rows[1:]
            if any(cell is not None and str(cell).strip() != '' for cell in row)
        ]

        if sheet_name in {'AYB', 'EPP', 'EXT', 'OPR', 'PA'}:
            course_totals.append({
                'curso': sheet_name,
                'total': summary_map.get(sheet_name, {}).get('total', len(data_rows))
            })

        if 'TIENE_SUFIJO' in headers:
            idx = headers.index('TIENE_SUFIJO')
            for row in data_rows:
                if len(row) <= idx or row[idx] is None:
                    continue
                value = str(row[idx]).strip().upper()
                if value in {'SI', 'SÍ', 'YES', 'TRUE', '1'}:
                    suffix_yes += 1
                else:
                    suffix_no += 1

    total_archivos = summary_map.get('TOTAL', {}).get('total', sum(item['total'] for item in course_totals))
    documentos_unicos = sum(item['unicos'] for item in summary_rows if isinstance(item.get('unicos'), int))
    duplicados = summary_map.get('DUPLICADOS', {}).get('total', 0)
    no_legibles = summary_map.get('NO_LEGIBLE', {}).get('total', 0)

    curso_top = max(course_totals, key=lambda item: item['total']) if course_totals else {'curso': '-', 'total': 0}
    payload = {
        'generatedAt': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'kpis': {
            'totalArchivos': total_archivos,
            'documentosUnicos': documentos_unicos,
            'duplicados': duplicados,
            'noLegibles': no_legibles,
            'conSufijo': suffix_yes,
            'sinSufijo': suffix_no,
        },
        'courseTotals': course_totals,
        'statusBreakdown': [
            {'label': 'Únicos', 'value': documentos_unicos, 'color': '#2d7ff9'},
            {'label': 'Duplicados', 'value': duplicados, 'color': '#f4c430'},
            {'label': 'No legibles', 'value': no_legibles, 'color': '#f53b4d'},
        ],
        'summaryRows': summary_rows,
        'insights': [
            {'title': 'Curso con mayor volumen', 'detail': f"{curso_top['curso']} concentra {curso_top['total']} archivos."},
            {'title': 'Calidad del registro', 'detail': f'Se identificaron {duplicados} duplicados y {no_legibles} documentos no legibles.'},
            {'title': 'Normalización de nombres', 'detail': f'{suffix_yes} archivos presentan sufijo especial y {suffix_no} quedaron con formato estándar.'}
        ],
    }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        'window.DASHBOARD_DATA = ' + json.dumps(payload, ensure_ascii=False, indent=2) + ';',
        encoding='utf-8',
    )
    print(f'Dashboard data generated at: {OUTPUT_PATH}')


if __name__ == '__main__':
    main()
