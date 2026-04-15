from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_PATH = DATA_DIR / 'dashboard_data.js'
PREFERRED_EXCEL_NAME = 'REGISTRO_FINAL_CURSOS.xlsx'


def to_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def find_excel_path() -> Path:
    preferred = DATA_DIR / PREFERRED_EXCEL_NAME
    if preferred.exists():
        return preferred

    candidates = sorted(DATA_DIR.glob('*.xlsx'), key=lambda item: item.stat().st_mtime, reverse=True)
    if candidates:
        return candidates[0]

    raise FileNotFoundError('No se encontró ningún archivo Excel en la carpeta data.')


def get_cell_value(row: tuple[Any, ...], headers: list[str], column_name: str, default: str = '') -> str:
    if column_name not in headers:
        return default
    idx = headers.index(column_name)
    if len(row) <= idx or row[idx] is None:
        return default
    return str(row[idx]).strip()


def get_status(sheet_name: str) -> str:
    if sheet_name == 'DUPLICADOS':
        return 'Duplicado'
    if sheet_name == 'NO_LEGIBLE':
        return 'No legible'
    return 'Vigente'


def main() -> None:
    excel_path = find_excel_path()
    wb = load_workbook(excel_path, data_only=True)

    summary_rows: list[dict[str, Any]] = []
    summary_map: dict[str, dict[str, Any]] = {}
    records: list[dict[str, str]] = []

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

        for row in data_rows:
            suffix_value = get_cell_value(row, headers, 'TIENE_SUFIJO', '')
            normalized_suffix = suffix_value.upper()
            if normalized_suffix in {'SI', 'SÍ', 'YES', 'TRUE', '1'}:
                suffix_yes += 1
            elif suffix_value:
                suffix_no += 1

            nombre = get_cell_value(row, headers, 'NOMBRE_PERSONA', 'Sin nombre')
            archivo = get_cell_value(row, headers, 'NOMBRE_ARCHIVO', 'Sin archivo')
            records.append({
                'nombre': nombre,
                'curso': sheet_name,
                'estado': get_status(sheet_name),
                'archivo': archivo,
            })

    total_archivos = summary_map.get('TOTAL', {}).get('total', sum(item['total'] for item in course_totals))
    documentos_unicos = sum(item['unicos'] for item in summary_rows if isinstance(item.get('unicos'), int))
    duplicados = summary_map.get('DUPLICADOS', {}).get('total', 0)
    no_legibles = summary_map.get('NO_LEGIBLE', {}).get('total', 0)

    curso_top = max(course_totals, key=lambda item: item['total']) if course_totals else {'curso': '-', 'total': 0}
    payload = {
        'generatedAt': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'sourceFile': excel_path.name,
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
        'records': records,
        'insights': [
            {'title': 'Curso con mayor volumen', 'detail': f"{curso_top['curso']} concentra {curso_top['total']} archivos."},
            {'title': 'Calidad del registro', 'detail': f'Se identificaron {duplicados} duplicados y {no_legibles} documentos no legibles.'},
            {'title': 'Normalización de nombres', 'detail': f'{suffix_yes} archivos presentan sufijo especial y {suffix_no} quedaron con formato estándar.'}
        ],
        'updateGuide': [
            'Copia tu archivo Excel más reciente dentro de la carpeta data del proyecto.',
            'Si deseas, mantén el nombre REGISTRO_FINAL_CURSOS.xlsx para reemplazar el anterior.',
            'Ejecuta el archivo actualizar_dashboard.bat para regenerar el panel local.',
            'Si también quieres actualizar el enlace público, ejecuta publicar_actualizacion.bat.'
        ],
    }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(
        'window.DASHBOARD_DATA = ' + json.dumps(payload, ensure_ascii=False, indent=2) + ';',
        encoding='utf-8',
    )
    print(f'Dashboard data generated at: {OUTPUT_PATH} using {excel_path.name}')


if __name__ == '__main__':
    main()
