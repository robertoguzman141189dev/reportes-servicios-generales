#!/usr/bin/env python3
"""
Procesa el archivo Excel y genera el reporte HTML ejecutivo
Compatible con el formato UPGCH 2026 (con y sin acentos en columnas)
"""

import pandas as pd
import json
import os
import unicodedata
from datetime import datetime
from pathlib import Path

INPUT_FILE  = 'data/input.xlsx'
OUTPUT_FILE = 'docs/index.html'
DATA_FILE   = 'docs/data.json'
TEMPLATE_FILE = 'scripts/template.html'

COORD_CFG = {
    'COORDINACION DE MATTO TECNICO Y': {'name': 'Mantenimiento',    'color': '#3b82f6', 'icon': '🔧'},
    'COORDINACION DE SEGURIDAD Y PC':  {'name': 'Seguridad y PC',   'color': '#ef4444', 'icon': '🛡️'},
    'COORDINACION DE GESTION DE ESPA': {'name': 'Gestión Espacios', 'color': '#10b981', 'icon': '🏛️'},
    'INFRAESTRUCTURA':                 {'name': 'Infraestructura',  'color': '#f59e0b', 'icon': '🏗️'},
}

# Estados válidos de "completado"
ESTADOS_COMPLETADO = {'COMPLETADO'}
ESTADOS_EN_PROCESO = {'EN PROCESO', 'INICIADOS', 'INICIADO'}
ESTADOS_REVISION   = {'EN REVISION/ESPERANDO VALIDACION', 'EN REVISION', 'EN REVISIÓN'}


def norm(text):
    """Quita acentos y normaliza texto para comparaciones."""
    if not isinstance(text, str):
        return ''
    t = unicodedata.normalize('NFD', text.upper().strip())
    return ''.join(c for c in t if unicodedata.category(c) != 'Mn')


def find_col(columns, *keywords):
    """Encuentra una columna que contenga todas las keywords (sin acentos)."""
    for col in columns:
        col_n = norm(col)
        if all(kw in col_n for kw in keywords):
            return col
    return None


def classify_estado(val):
    """Clasifica un estado normalizado."""
    v = norm(str(val))
    if v in {norm(e) for e in ESTADOS_COMPLETADO}:
        return 'COMPLETADO'
    if v in {norm(e) for e in ESTADOS_EN_PROCESO}:
        return 'EN PROCESO'
    if v in {norm(e) for e in ESTADOS_REVISION}:
        return 'EN REVISION'
    return 'PENDIENTE'


def process_sheet(df):
    """Extrae todas las stats y actividades de un DataFrame."""
    cols = list(df.columns)

    # Detectar columnas (ignorando acentos)
    col_estado    = find_col(cols, 'ESTADO')
    col_prio      = find_col(cols, 'PRIORIDAD')
    col_area      = find_col(cols, 'AREA', 'ASIGNADA')
    col_resp      = find_col(cols, 'RESPONSABLE')
    col_solic     = find_col(cols, 'SOLICITANTE')
    col_actividad = find_col(cols, 'ACTIVIDAD')

    print(f"    estado={col_estado} | prio={col_prio} | area={col_area} | resp={col_resp} | solic={col_solic}")

    # Normalizar estados
    estados_raw = df[col_estado].apply(classify_estado) if col_estado else pd.Series(['PENDIENTE'] * len(df))

    completadas = int((estados_raw == 'COMPLETADO').sum())
    en_proceso  = int((estados_raw == 'EN PROCESO').sum())
    en_revision = int((estados_raw == 'EN REVISION').sum())
    pendientes  = int((estados_raw == 'PENDIENTE').sum())
    total       = len(df)
    cumplimiento = round(completadas / total * 100, 1) if total > 0 else 0

    # Prioridades
    prios = {'alta': 0, 'media': 0, 'baja': 0}
    if col_prio:
        for val, cnt in df[col_prio].value_counts().items():
            v = norm(str(val))
            if v == 'ALTA':   prios['alta']  += int(cnt)
            elif v == 'MEDIA': prios['media'] += int(cnt)
            elif v == 'BAJA':  prios['baja']  += int(cnt)

    # Áreas, responsables, solicitantes
    areas        = []
    responsables = []
    solicitantes = []

    if col_area:
        areas = [{'name': str(k), 'value': int(v)}
                 for k, v in df[col_area].value_counts().head(8).items()]
    if col_resp:
        responsables = [{'name': str(k).strip(), 'value': int(v)}
                        for k, v in df[col_resp].value_counts().head(10).items()]
    if col_solic:
        solicitantes = [{'name': str(k)[:40], 'value': int(v)}
                        for k, v in df[col_solic].value_counts().head(8).items()]

    # Actividades individuales (para el tab de detalle)
    actividades = []
    if col_actividad:
        for _, row in df.iterrows():
            act = str(row.get(col_actividad, '')).strip()
            if not act or act == 'nan':
                continue
            actividades.append({
                'nombre':      act[:100],
                'estado':      str(row.get(col_estado, '')).strip() if col_estado else '',
                'prioridad':   str(row.get(col_prio,   '')).strip() if col_prio   else '',
                'responsable': str(row.get(col_resp,   '')).strip() if col_resp   else '',
            })

    return {
        'total':       total,
        'completadas': completadas,
        'enProceso':   en_proceso,
        'enRevision':  en_revision,
        'noIniciadas': pendientes,
        'cumplimiento': cumplimiento,
        'prioridades': prios,
        'areas':       areas,
        'responsables': responsables,
        'solicitantes': solicitantes,
        'actividades': actividades,
    }


def main():
    print("=" * 50)
    print("📊 GENERADOR DE REPORTES EJECUTIVOS UPGCH")
    print("=" * 50)

    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"No se encontró: {INPUT_FILE}")

    xlsx = pd.ExcelFile(INPUT_FILE)
    print(f"📋 Hojas: {xlsx.sheet_names}")

    chart_data = {
        'coordinaciones': [],
        'totales': {'actividades': 0, 'completadas': 0, 'enProceso': 0, 'pendientes': 0},
        'prioridades': {'ALTA': 0, 'MEDIA': 0, 'BAJA': 0},
        'generado': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    }

    for sheet in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet, header=0)
        df.columns = df.columns.str.strip()
        df = df.dropna(how='all')

        cfg = COORD_CFG.get(sheet, {'name': sheet, 'color': '#6b7280', 'icon': '📋'})
        print(f"\n  📂 {sheet} ({len(df)} filas) → {cfg['name']}")

        stats = process_sheet(df)

        chart_data['coordinaciones'].append({
            'id':          sheet,
            'name':        cfg['name'],
            'icon':        cfg['icon'],
            'color':       cfg['color'],
            **stats,
        })

        chart_data['totales']['actividades'] += stats['total']
        chart_data['totales']['completadas'] += stats['completadas']
        chart_data['totales']['enProceso']   += stats['enProceso'] + stats['enRevision']
        chart_data['totales']['pendientes']  += stats['noIniciadas']

        chart_data['prioridades']['ALTA']  += stats['prioridades']['alta']
        chart_data['prioridades']['MEDIA'] += stats['prioridades']['media']
        chart_data['prioridades']['BAJA']  += stats['prioridades']['baja']

        print(f"    ✅ {stats['completadas']}/{stats['total']} completadas | {stats['cumplimiento']}%")

    total = chart_data['totales']['actividades']
    comp  = chart_data['totales']['completadas']
    chart_data['totales']['tasaCumplimiento'] = round(comp / total * 100, 1) if total > 0 else 0

    print(f"\n📊 TOTALES: {total} actividades | {comp} completadas | {chart_data['totales']['tasaCumplimiento']}%")

    # Generar HTML desde template
    template_path = Path(TEMPLATE_FILE)
    if template_path.exists():
        html = template_path.read_text(encoding='utf-8')
    else:
        raise FileNotFoundError(f"Template no encontrado: {TEMPLATE_FILE}")

    data_json = json.dumps(chart_data, ensure_ascii=False)
    html = html.replace('{{DATA_JSON}}',          data_json)
    html = html.replace('{{FECHA_GENERACION}}',   chart_data['generado'])
    html = html.replace('{{TOTAL_ACTIVIDADES}}',  str(chart_data['totales']['actividades']))
    html = html.replace('{{TOTAL_COMPLETADAS}}',  str(chart_data['totales']['completadas']))
    html = html.replace('{{TASA_CUMPLIMIENTO}}',  str(chart_data['totales']['tasaCumplimiento']))
    html = html.replace('{{TOTAL_EN_PROCESO}}',   str(chart_data['totales']['enProceso']))
    html = html.replace('{{TOTAL_PENDIENTES}}',   str(chart_data['totales']['pendientes']))

    os.makedirs('docs', exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(chart_data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Reporte: {OUTPUT_FILE} ({os.path.getsize(OUTPUT_FILE)//1024} KB)")
    print(f"✅ Datos:   {DATA_FILE}")
    print("=" * 50)
    print("✅ PROCESO COMPLETADO")
    print("=" * 50)


if __name__ == '__main__':
    main()
