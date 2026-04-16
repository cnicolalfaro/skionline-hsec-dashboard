from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / 'data'
OUTPUT_PATH = DATA_DIR / 'dashboard_data.js'
PREFERRED_EXCEL_NAME = 'REGISTRO_FINAL_CURSOS.xlsx'
SPECIAL_SHEETS = {'RESUMEN', 'TARJA', 'NO_LEGIBLE', 'DUPLICADOS'}


def to_int(value: Any) -> int:
    try:
        return int(value)
    except (TypeError, ValueError):
        return 0


def find_excel_path() -> Path:
    search_paths = [
        DATA_DIR / PREFERRED_EXCEL_NAME,
        BASE_DIR / PREFERRED_EXCEL_NAME,
        BASE_DIR.parent / PREFERRED_EXCEL_NAME,
    ]

    for path in search_paths:
        if path.exists():
            return path

    candidates: list[Path] = []
    for folder in [DATA_DIR, BASE_DIR, BASE_DIR.parent]:
        candidates.extend(folder.glob('*.xlsx'))

    candidates = sorted(set(candidates), key=lambda item: item.stat().st_mtime, reverse=True)
    if candidates:
        return candidates[0]

    raise FileNotFoundError('No se encontró ningún archivo Excel para procesar.')


def normalize_text(value: Any) -> str:
    text = '' if value is None else str(value)
    text = unicodedata.normalize('NFD', text)
    text = ''.join(char for char in text if unicodedata.category(char) != 'Mn')
    text = re.sub(r'\bforms\b', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'\birl\s*general\b', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'[^A-Za-z ]+', ' ', text).lower()
    return ' '.join(text.split())


def fingerprint(*values: Any) -> str:
    words: list[str] = []
    for value in values:
        words.extend(normalize_text(value).split())
    return ' '.join(sorted(words))


def get_cell_value(row: tuple[Any, ...], headers: list[str], column_name: str, default: str = '') -> str:
    if column_name not in headers:
        return default
    idx = headers.index(column_name)
    if len(row) <= idx or row[idx] is None:
        return default
    return str(row[idx]).strip()


def is_evidence_sheet(sheet_name: str) -> bool:
    return sheet_name not in SPECIAL_SHEETS


def format_course_name(sheet_name: str) -> str:
    return sheet_name.replace('_', ' ').strip()


def format_rut(value: Any) -> str:
    text = '' if value is None else str(value).strip()
    cleaned = re.sub(r'[^0-9kK]', '', text)
    if len(cleaned) < 2:
        return text
    return f'{cleaned[:-1]}-{cleaned[-1].upper()}'


def get_first_available_cell(row: tuple[Any, ...], headers: list[str], column_names: list[str], default: str = '') -> str:
    normalized_headers = [header.upper().replace('.', '').strip() for header in headers]
    for column_name in column_names:
        normalized_name = column_name.upper().replace('.', '').strip()
        if normalized_name not in normalized_headers:
            continue
        idx = normalized_headers.index(normalized_name)
        if len(row) <= idx or row[idx] is None:
            continue
        value = str(row[idx]).strip()
        if value:
            return value
    return default


def load_external_rut_lookup() -> dict[str, str]:
    lookup: dict[str, str] = {}
    search_folders = [BASE_DIR, BASE_DIR.parent, DATA_DIR]

    for folder in search_folders:
        for path in sorted(folder.glob('TARJA*.xlsx')):
            try:
                wb = load_workbook(path, data_only=True)
            except Exception:
                continue

            if 'TARJA' not in wb.sheetnames:
                continue

            ws = wb['TARJA']
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = [str(h).strip() if h is not None else '' for h in rows[0]]
            if 'RUT' not in [header.upper().strip() for header in headers]:
                continue

            for row in rows[1:]:
                if not row or all(value is None or str(value).strip() == '' for value in row):
                    continue

                nombre = get_first_available_cell(row, headers, ['NOMBRE'])
                paterno = get_first_available_cell(row, headers, ['A. PATERNO'])
                materno = get_first_available_cell(row, headers, ['A. MATERNO'])
                rut = format_rut(get_first_available_cell(row, headers, ['RUT']))
                person_key = fingerprint(nombre, paterno, materno)

                if person_key and rut:
                    lookup[person_key] = rut

            if lookup:
                return lookup

    return lookup


def build_tarja_records(wb, evidence_map: dict[str, dict[str, Any]], external_rut_lookup: dict[str, str]) -> tuple[list[dict[str, Any]], int]:
    if 'TARJA' not in wb.sheetnames:
        return [], 0

    ws = wb['TARJA']
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], 0

    headers = [str(h).strip() if h is not None else '' for h in rows[0]]
    records: list[dict[str, Any]] = []
    sin_registros = 0

    for row in rows[1:]:
        if not row or all(value is None or str(value).strip() == '' for value in row):
            continue

        nombre = get_cell_value(row, headers, 'NOMBRE')
        paterno = get_cell_value(row, headers, 'A. PATERNO')
        materno = get_cell_value(row, headers, 'A. MATERNO')
        especialidad = get_cell_value(row, headers, 'ESPECIALIDAD', 'Sin especialidad informada')
        estado_tarja = get_cell_value(row, headers, 'ESTADO', 'Sin estado')

        nombre_mostrado = ' '.join(part for part in [paterno, materno, nombre] if part).strip() or 'Sin nombre'
        person_key = fingerprint(nombre, paterno, materno)
        evidence = evidence_map.get(person_key, {'courses': set(), 'flags': set(), 'notes': set()})
        courses = sorted(evidence.get('courses', set()))
        flags = evidence.get('flags', set())
        notes = evidence.get('notes', set())
        rut_tarja = format_rut(get_first_available_cell(row, headers, ['RUT', 'R.U.T', 'R U T'], ''))
        rut = rut_tarja or external_rut_lookup.get(person_key, '')

        if not courses and not flags:
            estado = 'No hay registros cargados, confirmar en portales'
            status_key = 'sin-registros'
            detalle = f'{especialidad} · Estado TARJA: {estado_tarja}'
            sin_registros += 1
        else:
            if courses:
                estado = 'Con registros'
                status_key = 'con-registros'
            elif 'no-legible' in flags:
                estado = 'No legible'
                status_key = 'no-legible'
            else:
                estado = 'Duplicado'
                status_key = 'duplicado'

            observaciones = []
            if 'duplicado' in flags:
                observaciones.append('incluye duplicados')
            if 'no-legible' in flags:
                observaciones.append('incluye no legibles')
            if 'irl-forms' in notes:
                observaciones.append('respaldo vía Forms; se recomienda validar el archivo original en la carpeta documental')
            extra = f" · {', '.join(observaciones)}" if observaciones else ''
            detalle = f'{especialidad} · Estado TARJA: {estado_tarja}{extra}'

        records.append({
            'nombre': nombre_mostrado,
            'rut': rut,
            'cursos': ', '.join(courses) if courses else '-',
            'courseList': courses,
            'estado': estado,
            'statusKey': status_key,
            'detalle': detalle,
        })

    records.sort(key=lambda item: normalize_text(item['nombre']))
    return records, sin_registros


def main() -> None:
    excel_path = find_excel_path()
    wb = load_workbook(excel_path, data_only=True)

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
    evidence_map: dict[str, dict[str, Any]] = {}

    for sheet_name in wb.sheetnames:
        if sheet_name in {'RESUMEN', 'TARJA'}:
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

        if is_evidence_sheet(sheet_name):
            course_totals.append({
                'curso': format_course_name(sheet_name),
                'total': summary_map.get(sheet_name, {}).get('total', len(data_rows))
            })

        if 'NOMBRE_PERSONA' not in headers:
            continue

        for row in data_rows:
            nombre_persona = get_cell_value(row, headers, 'NOMBRE_PERSONA')
            person_key = fingerprint(nombre_persona)
            if not person_key:
                continue

            current = evidence_map.setdefault(person_key, {'courses': set(), 'flags': set(), 'notes': set()})

            if is_evidence_sheet(sheet_name):
                current['courses'].add(format_course_name(sheet_name))
                if format_course_name(sheet_name) == 'IRL GENERAL FORMS':
                    current['notes'].add('irl-forms')
            elif sheet_name == 'NO_LEGIBLE':
                current['flags'].add('no-legible')
            elif sheet_name == 'DUPLICADOS':
                current['flags'].add('duplicado')

    external_rut_lookup = load_external_rut_lookup()
    records, sin_registros = build_tarja_records(wb, evidence_map, external_rut_lookup)

    total_archivos = summary_map.get('TOTAL', {}).get('total', sum(item['total'] for item in course_totals))
    documentos_unicos = sum(item['unicos'] for item in summary_rows if isinstance(item.get('unicos'), int))
    duplicados = summary_map.get('DUPLICADOS', {}).get('total', 0)
    no_legibles = summary_map.get('NO_LEGIBLE', {}).get('total', 0)
    trabajadores_tarja = len(records)
    con_registros = max(trabajadores_tarja - sin_registros, 0)

    curso_top = max(course_totals, key=lambda item: item['total']) if course_totals else {'curso': '-', 'total': 0}
    payload = {
        'generatedAt': datetime.now().strftime('%d/%m/%Y %H:%M'),
        'sourceFile': excel_path.name,
        'kpis': {
            'totalArchivos': total_archivos,
            'documentosUnicos': documentos_unicos,
            'duplicados': duplicados,
            'noLegibles': no_legibles,
            'trabajadoresTarja': trabajadores_tarja,
            'conRegistros': con_registros,
            'sinRegistros': sin_registros,
        },
        'courseTotals': course_totals,
        'statusBreakdown': [
            {'label': 'Con registros', 'value': con_registros, 'color': '#51b847'},
            {'label': 'Sin registros', 'value': sin_registros, 'color': '#ff7a59'},
            {'label': 'Duplicados', 'value': duplicados, 'color': '#f4c430'},
            {'label': 'No legibles', 'value': no_legibles, 'color': '#f53b4d'},
        ],
        'summaryRows': summary_rows,
        'records': records,
        'irlFormsNote': 'Los registros asociados a IRL GENERAL FORMS corresponden a respaldos cargados por los trabajadores mediante Forms. Para una validación formal, se recomienda revisar directamente el archivo original en la carpeta documental, a fin de confirmar su legibilidad, integridad y correcta carga. Los accesos disponibles a continuación funcionan solo para personal previamente autorizado en SharePoint.',
        'accessLinks': [
            {'label': 'Evidencias de certificaciones', 'url': 'https://empresassk.sharepoint.com/:f:/r/sites/ICSK-HSEC/Documentos%20compartidos/05%20-%20Respaldo%20HSEC%20faenas/250%20-%20Mantenimiento%20M2%20y%20M3/Contrato%20250%20Dch/Sistema%20de%20Gesti%C3%B3n%20n%20contrato%204600030982/3-%20Registros%20capacitaciones%20-%20difusiones/3-%20Evidencias%20de%20Certificaciones?csf=1&web=1&e=ovhAFT'},
            {'label': 'Procedimientos contrato 982', 'url': 'https://empresassk.sharepoint.com/:f:/r/sites/ICSK-HSEC/Documentos%20compartidos/05%20-%20Respaldo%20HSEC%20faenas/250%20-%20Mantenimiento%20M2%20y%20M3/Contrato%20250%20Dch/Sistema%20de%20Gesti%C3%B3n%20n%20contrato%204600030982/2-%20Procedimientos?csf=1&web=1&e=IsTxdJ'},
            {'label': 'Procedimientos contrato 984', 'url': 'https://empresassk.sharepoint.com/:f:/r/sites/ICSK-HSEC/Documentos%20compartidos/05%20-%20Respaldo%20HSEC%20faenas/250%20-%20Mantenimiento%20M2%20y%20M3/Contrato%20250%20Dch/Sistema%20de%20Gesti%C3%B3n%20n%20contrato%204600030984/2-%20Procedimientos?csf=1&web=1&e=RJlP81'},
            {'label': 'Check list actualizados', 'url': 'https://empresassk.sharepoint.com/:f:/r/sites/ICSK-HSEC/Documentos%20compartidos/05%20-%20Respaldo%20HSEC%20faenas/250%20-%20Mantenimiento%20M2%20y%20M3/Contrato%20250%20Dch/Sistema%20de%20Gesti%C3%B3n%20n%20contrato%204600030982/4-%20Formatos%20terreno/01-%20Check%20list/Check%20List%20Actualizados?csf=1&web=1&e=5arIgO'}
        ],
        'insights': [
            {'title': 'Curso con mayor volumen', 'detail': f"{curso_top['curso']} concentra {curso_top['total']} archivos."},
            {'title': 'Cruce con TARJA', 'detail': f'Se detectaron {sin_registros} trabajadores sin evidencias cargadas en el sistema.'},
            {'title': 'Calidad del registro', 'detail': f'Se identificaron {duplicados} duplicados y {no_legibles} documentos no legibles.'}
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
