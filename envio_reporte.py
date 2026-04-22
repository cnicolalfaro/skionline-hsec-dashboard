"""Genera un PDF ejecutivo del dashboard HSEC y lo envía por correo.

Lee los datos desde data/dashboard_data.js (generado por process_excel.py).
Credenciales de correo en config_correo.json (NO subir a git).
"""
from __future__ import annotations

import json
import re
import smtplib
import sys
from datetime import datetime
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / 'data' / 'dashboard_data.js'
CONFIG_FILE = BASE_DIR / 'config_correo.json'
OUTPUT_DIR = BASE_DIR / 'reportes_ejecutivos'
LOGO_FILE = BASE_DIR / 'assets' / 'logo.png'

COLORS = {
    'bg': '#0b1d33',
    'panel': '#0f2744',
    'primary': '#2d7ff9',
    'accent': '#8c63ff',
    'green': '#51b847',
    'orange': '#ff7a59',
    'yellow': '#ffcc66',
    'red': '#e74c3c',
    'text': '#eef4ff',
    'muted': '#9aacca',
}


def load_dashboard_data() -> dict:
    text = DATA_FILE.read_text(encoding='utf-8')
    match = re.search(r'window\.DASHBOARD_DATA\s*=\s*(\{.*\});', text, re.S)
    if not match:
        raise RuntimeError('No se pudo leer dashboard_data.js')
    return json.loads(match.group(1))


def load_config() -> dict:
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f'Falta {CONFIG_FILE.name}. Crealo con SMTP y destinatarios.'
        )
    return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))


def _fmt_number(value: int | float | None) -> str:
    if value is None:
        return '0'
    return f'{int(value):,}'.replace(',', '.')


def chart_distribucion(course_totals: list[dict], output: Path) -> None:
    cursos = [c['curso'] for c in course_totals]
    totales = [c['total'] for c in course_totals]

    fig, ax = plt.subplots(figsize=(7.5, 4.2), facecolor=COLORS['panel'])
    ax.set_facecolor(COLORS['panel'])
    y_pos = range(len(cursos))
    bars = ax.barh(y_pos, totales, color=COLORS['primary'], edgecolor=COLORS['accent'])
    ax.set_yticks(list(y_pos))
    ax.set_yticklabels(cursos, color=COLORS['text'], fontsize=10)
    ax.invert_yaxis()
    ax.tick_params(axis='x', colors=COLORS['muted'])
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color(COLORS['muted'])
    ax.spines['bottom'].set_color(COLORS['muted'])
    for bar, total in zip(bars, totales):
        ax.text(
            bar.get_width() + max(totales) * 0.01,
            bar.get_y() + bar.get_height() / 2,
            _fmt_number(total),
            va='center', color=COLORS['text'], fontsize=9, fontweight='bold',
        )
    ax.set_xlim(0, max(totales) * 1.15 if totales else 1)
    ax.set_title('Distribución por curso', color=COLORS['text'], fontsize=12, pad=12, loc='left')
    fig.tight_layout()
    fig.savefig(output, dpi=160, facecolor=COLORS['panel'])
    plt.close(fig)


def chart_donut(kpis: dict, output: Path) -> None:
    con = kpis.get('conRegistros', 0)
    sin = kpis.get('sinRegistros', 0)
    dup = kpis.get('duplicados', 0)
    noleg = kpis.get('noLegibles', 0)
    values = [con, sin, dup, noleg]
    labels = ['Con registros', 'Sin registros', 'Duplicados', 'No legibles']
    palette = [COLORS['green'], COLORS['orange'], COLORS['yellow'], COLORS['red']]

    fig, ax = plt.subplots(figsize=(6.4, 4.2), facecolor=COLORS['panel'])
    ax.set_facecolor(COLORS['panel'])
    wedges, _ = ax.pie(
        values, colors=palette, startangle=90,
        wedgeprops=dict(width=0.32, edgecolor=COLORS['panel']),
        radius=0.85,
    )
    total = sum(values) or 1
    ax.text(0, 0, f"{total}\narchivos", ha='center', va='center',
            color=COLORS['text'], fontsize=13, fontweight='bold')
    legend_labels = [f"{label}  {_fmt_number(val)}" for label, val in zip(labels, values)]
    ax.legend(
        wedges, legend_labels, loc='center left', bbox_to_anchor=(1.0, 0.5),
        frameon=False, labelcolor=COLORS['text'], fontsize=9,
    )
    ax.set_title('Estado del registro', color=COLORS['text'], fontsize=12, pad=12, loc='left')
    fig.subplots_adjust(left=0.02, right=0.72, top=0.88, bottom=0.05)
    fig.savefig(output, dpi=160, facecolor=COLORS['panel'])
    plt.close(fig)


COURSE_COLUMNS_PDF = [
    'EVALUACIONES IRL', 'IRL ESPECIFICA', 'IRL GENERAL',
    'AYB', 'EPP', 'EXT', 'OPR', 'PA',
]


def chart_acr_compliance(records: list[dict], output: Path) -> dict:
    total_cursos = len(COURSE_COLUMNS_PDF)
    acreditados = [
        r for r in records
        if (r.get('acrSucal') or '').strip().lower() == 'acreditado'
    ]
    buckets = [
        ('100% completo', 'COMPLETO', 100, 100, '#51b847'),
        ('75% - 99%',     'AVANZADO', 75,  99,  '#7dd87a'),
        ('50% - 74%',     'MEDIO',    50,  74,  '#f4c430'),
        ('25% - 49%',     'BAJO',     25,  49,  '#ff7a59'),
        ('0% - 24%',      'CRÍTICO',  0,   24,  '#f53b4d'),
    ]
    counts = [0] * len(buckets)
    sum_pct = 0
    for r in acreditados:
        course_set = set(r.get('courseList') or [])
        found = sum(1 for c in COURSE_COLUMNS_PDF if c in course_set)
        pct = round((found / total_cursos) * 100) if total_cursos else 0
        sum_pct += pct
        for i, (_, _s, lo, hi, _c) in enumerate(buckets):
            if lo <= pct <= hi:
                counts[i] += 1
                break

    total_acr = len(acreditados)
    promedio = round(sum_pct / total_acr) if total_acr else 0
    completos = counts[0]
    criticos = counts[3] + counts[4]

    # --- Figura ---
    fig = plt.figure(figsize=(11.0, 4.4), facecolor=COLORS['panel'])
    gs = fig.add_gridspec(
        2, 5,
        height_ratios=[1.1, 2.4],
        hspace=0.35, wspace=0.18,
        left=0.03, right=0.985, top=0.86, bottom=0.06,
    )

    # Título general
    fig.text(
        0.03, 0.94,
        'Cumplimiento de cursos — trabajadores acreditados',
        color=COLORS['text'], fontsize=13, fontweight='bold', ha='left',
    )
    fig.text(
        0.03, 0.895,
        f'Distribución de cumplimiento entre {total_acr} trabajadores acreditados '
        f'(de {len(records)} totales).',
        color=COLORS['muted'], fontsize=9, ha='left',
    )

    # --- Fila superior: 3 stats (PROMEDIO / COMPLETOS / CRÍTICOS) ---
    stats = [
        ('PROMEDIO',      f'{promedio}%',        COLORS['text']),
        ('COMPLETOS',     f'{completos}',        '#7dd87a'),
        ('CRÍTICOS <50%', f'{criticos}',         '#ff7a59'),
    ]
    # Dejamos los primeros 3 espacios de la fila superior para stats, los otros 2 vacíos
    for i, (lbl, val, col) in enumerate(stats):
        ax = fig.add_subplot(gs[0, i + 2])
        ax.set_facecolor('#14324a')
        for spine in ax.spines.values():
            spine.set_color('#1e4360')
            spine.set_linewidth(0.8)
        ax.set_xticks([]); ax.set_yticks([])
        ax.text(0.5, 0.72, lbl, ha='center', va='center',
                color=COLORS['muted'], fontsize=8, fontweight='bold',
                transform=ax.transAxes)
        ax.text(0.5, 0.32, val, ha='center', va='center',
                color=col, fontsize=16, fontweight='bold',
                transform=ax.transAxes)

    # --- Fila inferior: 5 "tarjetas" por tramo ---
    icons = ['✓', '◐', '◑', '◔', '!']
    for i, (label, short, _lo, _hi, color_hex) in enumerate(buckets):
        ax = fig.add_subplot(gs[1, i])
        ax.set_facecolor('#14324a')
        for spine in ax.spines.values():
            spine.set_visible(False)
        ax.set_xticks([]); ax.set_yticks([])
        ax.set_xlim(0, 1); ax.set_ylim(0, 1)

        # Borde superior de color (acento)
        ax.plot([0.02, 0.98], [0.985, 0.985], color=color_hex, linewidth=3, solid_capstyle='butt')

        # Ícono + etiquetas
        ax.text(0.07, 0.84, icons[i], color=color_hex, fontsize=14, fontweight='bold',
                ha='left', va='center')
        ax.text(0.22, 0.88, label, color=COLORS['text'], fontsize=9, fontweight='bold',
                ha='left', va='center')
        ax.text(0.22, 0.73, short, color=COLORS['muted'], fontsize=7.5,
                ha='left', va='center')

        # Valor grande
        ax.text(0.06, 0.42, f'{counts[i]}', color=COLORS['text'],
                fontsize=26, fontweight='bold', ha='left', va='center')

        # Barra inferior
        pct_acr = round((counts[i] / total_acr) * 100) if total_acr else 0
        bar_y = 0.18
        bar_h = 0.04
        ax.add_patch(plt.Rectangle((0.06, bar_y), 0.88, bar_h,
                                   facecolor='#1e3a52', edgecolor='none'))
        ax.add_patch(plt.Rectangle((0.06, bar_y), 0.88 * (pct_acr / 100.0), bar_h,
                                   facecolor=color_hex, edgecolor='none'))
        ax.text(0.06, 0.08, f'{pct_acr}% de acreditados',
                color=COLORS['muted'], fontsize=7.5, ha='left', va='center')

    fig.savefig(output, dpi=170, facecolor=COLORS['panel'])
    plt.close(fig)
    return {
        'acreditados': total_acr,
        'promedio': promedio,
        'completos': completos,
        'criticos': criticos,
        'counts': counts,
    }


def build_pdf(data: dict, pdf_path: Path) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    chart1 = OUTPUT_DIR / '_chart_dist.png'
    chart2 = OUTPUT_DIR / '_chart_donut.png'
    chart3 = OUTPUT_DIR / '_chart_acr.png'
    chart_distribucion(data.get('courseTotals', []), chart1)
    chart_donut(data.get('kpis', {}), chart2)
    chart_acr_compliance(data.get('records', []), chart3)

    doc = SimpleDocTemplate(
        str(pdf_path), pagesize=A4,
        leftMargin=14 * mm, rightMargin=14 * mm, topMargin=14 * mm, bottomMargin=14 * mm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='TitleWhite', parent=styles['Title'], textColor=colors.HexColor(COLORS['text']), fontSize=20, leading=24))
    styles.add(ParagraphStyle(name='Sub', parent=styles['Normal'], textColor=colors.HexColor(COLORS['muted']), fontSize=10))
    styles.add(ParagraphStyle(name='KpiCard', parent=styles['Normal'], textColor=colors.HexColor(COLORS['text']), fontSize=9, alignment=1, leading=12))

    story: list = []

    # Encabezado
    header_cells = [[
        Image(str(LOGO_FILE), width=40 * mm, height=18 * mm) if LOGO_FILE.exists() else '',
        [
            Paragraph('<b>SK INDUSTRIAL S.A. · HSEC</b>', styles['Sub']),
            Paragraph('Dashboard de Registro Final de Cursos', styles['TitleWhite']),
            Paragraph('Vista ejecutiva del consolidado de diplomas y archivos procesados.', styles['Sub']),
        ],
        Paragraph(f"<b>Actualizado:</b><br/>{data.get('generatedAt', '')}", styles['Sub']),
    ]]
    header = Table(header_cells, colWidths=[45 * mm, None, 40 * mm])
    header.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(COLORS['bg'])),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
        ('ROUNDEDCORNERS', (0, 0), (-1, -1), [8, 8, 8, 8]),
    ]))
    story.append(header)
    story.append(Spacer(1, 6 * mm))

    # Tarjetas KPI
    kpis = data.get('kpis', {})
    kpi_defs = [
        ('Total de archivos', kpis.get('totalArchivos'), COLORS['primary'], 'Evidencias y documentos procesados.'),
        ('Total trabajadores en TARJA', kpis.get('trabajadoresTarja'), COLORS['accent'], 'Dotación total considerada en el cruce.'),
        ('Trabajadores con registros', kpis.get('conRegistros'), COLORS['green'], 'Con al menos una evidencia asociada.'),
        ('Trabajadores sin registro', kpis.get('sinRegistros'), COLORS['orange'], 'Sin evidencia; confirmar en portales.'),
    ]

    def _card(label, value, color_hex, helptxt):
        muted = COLORS['muted']
        text = COLORS['text']
        html = (
            f'<para align="center" leading="14">'
            f'<font size="9" color="{muted}">{label}</font><br/>'
            f'<br/>'
            f'<font size="24" color="{text}"><b>{_fmt_number(value)}</b></font>'
            f'<br/><br/>'
            f'<font size="7" color="{muted}">{helptxt}</font>'
            f'</para>'
        )
        inner = Table(
            [[Paragraph(html, styles['KpiCard'])]],
            colWidths=[42 * mm],
            rowHeights=[38 * mm],
        )
        inner.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(COLORS['panel'])),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor(color_hex)),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor(color_hex)),
        ]))
        return inner

    cards_row = [[_card(*d) for d in kpi_defs]]
    kpi_table = Table(cards_row, colWidths=[45 * mm] * 4)
    kpi_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('LEFTPADDING', (0, 0), (-1, -1), 2),
        ('RIGHTPADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 6 * mm))

    # Gráficos
    charts_row = Table(
        [[Image(str(chart1), width=95 * mm, height=62 * mm),
          Image(str(chart2), width=85 * mm, height=62 * mm)]],
        colWidths=[95 * mm, 85* mm],
    )
    charts_row.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(COLORS['panel'])),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(charts_row)
    story.append(Spacer(1, 6 * mm))

    # Tercer gráfico: cumplimiento de acreditados (ancho completo)
    acr_panel = Table(
        [[Image(str(chart3), width=180 * mm, height=72 * mm)]],
        colWidths=[180 * mm],
    )
    acr_panel.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(COLORS['panel'])),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(acr_panel)
    story.append(Spacer(1, 8 * mm))
    story.append(Paragraph(
        f"Reporte generado automáticamente por el sistema SK Industrial HSEC el {datetime.now().strftime('%d/%m/%Y %H:%M')}. "
        f"Dashboard en vivo: https://cnicolalfaro.github.io/skionline-hsec-dashboard/",
        styles['Sub'],
    ))

    def _page_bg(canv, doc_):
        canv.saveState()
        canv.setFillColor(colors.HexColor(COLORS['bg']))
        canv.rect(0, 0, A4[0], A4[1], stroke=0, fill=1)
        canv.restoreState()

    doc.build(story, onFirstPage=_page_bg, onLaterPages=_page_bg)

    # limpiar charts auxiliares
    for p in (chart1, chart2, chart3):
        try:
            p.unlink()
        except OSError:
            pass

    return pdf_path


def send_mail(pdf_path: Path, data: dict, config: dict) -> None:
    kpis = data.get('kpis', {})
    asunto = f"Reporte Ejecutivo HSEC · {data.get('generatedAt', datetime.now().strftime('%d/%m/%Y %H:%M'))}"
    cuerpo_html = f"""
    <p>Estimados,</p>
    <p>Adjunto el reporte ejecutivo del <b>Dashboard HSEC SK Industrial</b>, actualizado automáticamente.</p>
    <ul>
      <li>Total archivos: <b>{_fmt_number(kpis.get('totalArchivos'))}</b></li>
      <li>Trabajadores en TARJA: <b>{_fmt_number(kpis.get('trabajadoresTarja'))}</b></li>
      <li>Con registros: <b>{_fmt_number(kpis.get('conRegistros'))}</b></li>
      <li>Sin registro: <b>{_fmt_number(kpis.get('sinRegistros'))}</b></li>
    </ul>
    <p>Dashboard en vivo:
      <a href="https://cnicolalfaro.github.io/skionline-hsec-dashboard/">cnicolalfaro.github.io/skionline-hsec-dashboard</a>
    </p>
    <p style="color:#888; font-size:11px;">Correo automático, no responder.</p>
    """

    msg = MIMEMultipart('mixed')
    msg['Subject'] = asunto
    from_addr = f"{config.get('remitente_nombre', 'Dashboard HSEC')} <{config['smtp_user']}>"
    msg['From'] = from_addr
    destinatarios = config.get('destinatarios', [])
    msg['To'] = ', '.join(destinatarios)

    alt = MIMEMultipart('alternative')
    alt.attach(MIMEText(cuerpo_html, 'html', 'utf-8'))
    msg.attach(alt)

    with open(pdf_path, 'rb') as f:
        part = MIMEBase('application', 'pdf')
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{pdf_path.name}"')
    msg.attach(part)

    host = config['smtp_host']
    port = int(config['smtp_port'])
    with smtplib.SMTP(host, port, timeout=60) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(config['smtp_user'], config['smtp_password'])
        smtp.sendmail(config['smtp_user'], destinatarios, msg.as_string())


def main() -> int:
    data = load_dashboard_data()
    config = load_config()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime('%Y%m%d_%H%M')
    pdf_path = OUTPUT_DIR / f'reporte_ejecutivo_hsec_{stamp}.pdf'
    build_pdf(data, pdf_path)
    print(f'PDF generado: {pdf_path}')

    try:
        send_mail(pdf_path, data, config)
        print(f"Correo enviado a: {', '.join(config.get('destinatarios', []))}")
    except Exception as exc:  # noqa: BLE001
        print(f'ERROR al enviar correo: {exc}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
