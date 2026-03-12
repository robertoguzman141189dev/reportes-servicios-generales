#!/usr/bin/env python3
"""
Procesa el archivo Excel y genera el reporte HTML ejecutivo
"""

import pandas as pd
import json
import os
from datetime import datetime
from pathlib import Path

# Paths
INPUT_FILE = 'data/input.xlsx'
OUTPUT_FILE = 'docs/index.html'
TEMPLATE_FILE = 'scripts/template.html'


def load_excel_data(filepath):
    """Carga y procesa todas las hojas del Excel"""
    
    print(f"📂 Cargando archivo: {filepath}")
    xlsx = pd.ExcelFile(filepath)
    
    print(f"📋 Hojas encontradas: {xlsx.sheet_names}")
    
    data = {}
    
    for sheet_name in xlsx.sheet_names:
        df = pd.read_excel(xlsx, sheet_name=sheet_name, header=0)
        df.columns = df.columns.str.strip().str.upper()
        
        # Limpiar datos
        df = df.dropna(how='all')
        
        data[sheet_name] = {
            'df': df,
            'stats': extract_stats(df, sheet_name)
        }
        
        print(f"  ✅ {sheet_name}: {len(df)} registros")
    
    return data


def extract_stats(df, sheet_name):
    """Extrae estadísticas de un DataFrame"""
    
    stats = {
        'total': len(df),
        'estados': {},
        'prioridades': {},
        'areas': {},
        'responsables': {},
        'solicitantes': {}
    }
    
    # Buscar columnas relevantes
    for col in df.columns:
        col_upper = col.upper()
        
        if 'ESTADO' in col_upper:
            stats['estados'] = df[col].value_counts().to_dict()
        
        if 'PRIORIDAD' in col_upper:
            stats['prioridades'] = df[col].value_counts().to_dict()
        
        if 'AREA' in col_upper and 'ASIGNADA' in col_upper:
            stats['areas'] = df[col].value_counts().head(10).to_dict()
        
        if 'RESPONSABLE' in col_upper:
            stats['responsables'] = df[col].value_counts().head(10).to_dict()
        
        if 'SOLICITANTE' in col_upper:
            stats['solicitantes'] = df[col].value_counts().head(10).to_dict()
    
    # Calcular cumplimiento
    completadas = stats['estados'].get('COMPLETADO', 0)
    stats['cumplimiento'] = round((completadas / stats['total'] * 100), 1) if stats['total'] > 0 else 0
    
    return stats


def generate_chart_data(data):
    """Genera los datos para las gráficas en formato JSON"""
    
    # Mapeo de nombres de hojas
    coord_names = {
        'COORDINACION DE MATTO TECNICO Y': {'name': 'Mantenimiento', 'color': '#3b82f6', 'icon': '🔧'},
        'COORDINACION DE SEGURIDAD Y PC': {'name': 'Seguridad y PC', 'color': '#ef4444', 'icon': '🛡️'},
        'COORDINACION DE GESTION DE ESPA': {'name': 'Gestión Espacios', 'color': '#10b981', 'icon': '🏛️'},
        'INFRAESTRUCTURA': {'name': 'Infraestructura', 'color': '#f59e0b', 'icon': '🏗️'}
    }
    
    chart_data = {
        'coordinaciones': [],
        'totales': {
            'actividades': 0,
            'completadas': 0,
            'enProceso': 0,
            'pendientes': 0
        },
        'prioridades': {'ALTA': 0, 'MEDIA': 0, 'BAJA': 0},
        'generado': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    }
    
    for sheet_name, sheet_data in data.items():
        stats = sheet_data['stats']
        config = coord_names.get(sheet_name, {'name': sheet_name, 'color': '#6b7280', 'icon': '📋'})
        
        completadas = stats['estados'].get('COMPLETADO', 0)
        en_proceso = stats['estados'].get('EN PROCESO', 0)
        no_iniciadas = stats['estados'].get('NO INICIADO', 0)
        en_revision = stats['estados'].get('EN REVISION/ESPERANDO VALIDACION', 0) + stats['estados'].get('EN REVISIÓN', 0)
        
        coord_info = {
            'id': sheet_name,
            'name': config['name'],
            'icon': config['icon'],
            'color': config['color'],
            'total': stats['total'],
            'completadas': completadas,
            'enProceso': en_proceso,
            'noIniciadas': no_iniciadas,
            'enRevision': en_revision,
            'cumplimiento': stats['cumplimiento'],
            'prioridades': {
                'alta': stats['prioridades'].get('ALTA', 0),
                'media': stats['prioridades'].get('MEDIA', 0),
                'baja': stats['prioridades'].get('BAJA', 0)
            },
            'areas': [{'name': k, 'value': v} for k, v in list(stats['areas'].items())[:6]],
            'responsables': [{'name': k, 'value': v} for k, v in list(stats['responsables'].items())[:5]],
            'solicitantes': [{'name': k, 'value': v} for k, v in list(stats['solicitantes'].items())[:5]]
        }
        
        chart_data['coordinaciones'].append(coord_info)
        
        # Acumular totales
        chart_data['totales']['actividades'] += stats['total']
        chart_data['totales']['completadas'] += completadas
        chart_data['totales']['enProceso'] += en_proceso
        chart_data['totales']['pendientes'] += no_iniciadas + en_revision
        
        # Acumular prioridades
        chart_data['prioridades']['ALTA'] += stats['prioridades'].get('ALTA', 0)
        chart_data['prioridades']['MEDIA'] += stats['prioridades'].get('MEDIA', 0)
        chart_data['prioridades']['BAJA'] += stats['prioridades'].get('BAJA', 0)
    
    # Calcular tasa de cumplimiento global
    total = chart_data['totales']['actividades']
    completadas = chart_data['totales']['completadas']
    chart_data['totales']['tasaCumplimiento'] = round((completadas / total * 100), 1) if total > 0 else 0
    
    return chart_data


def generate_html(chart_data):
    """Genera el HTML del reporte usando el template"""
    
    # Leer template
    template_path = Path(TEMPLATE_FILE)
    if template_path.exists():
        with open(template_path, 'r', encoding='utf-8') as f:
            html = f.read()
    else:
        print("⚠️ Template no encontrado, usando template embebido")
        html = get_embedded_template()
    
    # Insertar datos
    html = html.replace('{{DATA_JSON}}', json.dumps(chart_data, ensure_ascii=False, indent=2))
    html = html.replace('{{FECHA_GENERACION}}', chart_data['generado'])
    html = html.replace('{{TOTAL_ACTIVIDADES}}', str(chart_data['totales']['actividades']))
    html = html.replace('{{TOTAL_COMPLETADAS}}', str(chart_data['totales']['completadas']))
    html = html.replace('{{TASA_CUMPLIMIENTO}}', str(chart_data['totales']['tasaCumplimiento']))
    html = html.replace('{{TOTAL_EN_PROCESO}}', str(chart_data['totales']['enProceso']))
    html = html.replace('{{TOTAL_PENDIENTES}}', str(chart_data['totales']['pendientes']))
    
    return html


def get_embedded_template():
    """Retorna un template HTML básico embebido"""
    return '''<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Reporte Ejecutivo - Servicios Generales</title>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { font-family: system-ui, sans-serif; background: #f1f5f9; padding: 20px; }
        .container { max-width: 1400px; margin: 0 auto; }
        .header { background: linear-gradient(135deg, #2563eb, #7c3aed); border-radius: 16px; padding: 30px; color: white; margin-bottom: 20px; }
        .header h1 { font-size: 2rem; }
        .header p { opacity: 0.9; margin-top: 5px; }
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 20px; margin-bottom: 20px; }
        .kpi-card { background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .kpi-card .value { font-size: 2.5rem; font-weight: bold; }
        .kpi-card .label { color: #64748b; margin-top: 5px; }
        .chart-card { background: white; border-radius: 12px; padding: 20px; margin-bottom: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .chart-card h3 { margin-bottom: 15px; color: #1e293b; }
        .chart-container { height: 300px; }
        .footer { text-align: center; padding: 20px; color: #64748b; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📊 Reporte Ejecutivo</h1>
            <p>Servicios Generales - Generado: {{FECHA_GENERACION}}</p>
        </div>
        
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="value" style="color: #3b82f6;">{{TOTAL_ACTIVIDADES}}</div>
                <div class="label">Total Actividades</div>
            </div>
            <div class="kpi-card">
                <div class="value" style="color: #10b981;">{{TOTAL_COMPLETADAS}}</div>
                <div class="label">Completadas</div>
            </div>
            <div class="kpi-card">
                <div class="value" style="color: #10b981;">{{TASA_CUMPLIMIENTO}}%</div>
                <div class="label">Cumplimiento</div>
            </div>
            <div class="kpi-card">
                <div class="value" style="color: #f59e0b;">{{TOTAL_EN_PROCESO}}</div>
                <div class="label">En Proceso</div>
            </div>
        </div>
        
        <div class="chart-card">
            <h3>Estado por Coordinación</h3>
            <div class="chart-container">
                <canvas id="mainChart"></canvas>
            </div>
        </div>
        
        <div class="footer">
            <p>Reporte generado automáticamente</p>
        </div>
    </div>
    
    <script>
        const reportData = {{DATA_JSON}};
        
        new Chart(document.getElementById('mainChart'), {
            type: 'bar',
            data: {
                labels: reportData.coordinaciones.map(c => c.name),
                datasets: [
                    { label: 'Completadas', data: reportData.coordinaciones.map(c => c.completadas), backgroundColor: '#10b981' },
                    { label: 'En Proceso', data: reportData.coordinaciones.map(c => c.enProceso), backgroundColor: '#f59e0b' },
                    { label: 'Pendientes', data: reportData.coordinaciones.map(c => c.noIniciadas + c.enRevision), backgroundColor: '#6b7280' }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: { x: { stacked: true }, y: { stacked: true } }
            }
        });
    </script>
</body>
</html>'''


def main():
    print("=" * 50)
    print("📊 GENERADOR DE REPORTES EJECUTIVOS")
    print("=" * 50)
    
    # Verificar archivo de entrada
    if not os.path.exists(INPUT_FILE):
        raise FileNotFoundError(f"No se encontró el archivo: {INPUT_FILE}")
    
    # Cargar datos
    data = load_excel_data(INPUT_FILE)
    
    # Generar datos para gráficas
    print("\n📈 Procesando estadísticas...")
    chart_data = generate_chart_data(data)
    
    print(f"\n📊 Resumen:")
    print(f"   Total actividades: {chart_data['totales']['actividades']}")
    print(f"   Completadas: {chart_data['totales']['completadas']}")
    print(f"   Tasa cumplimiento: {chart_data['totales']['tasaCumplimiento']}%")
    
    # Generar HTML
    print("\n🔧 Generando HTML...")
    html = generate_html(chart_data)
    
    # Guardar
    os.makedirs(os.path.dirname(OUTPUT_FILE), exist_ok=True)
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"\n✅ Reporte generado: {OUTPUT_FILE}")
    print(f"   Tamaño: {os.path.getsize(OUTPUT_FILE) / 1024:.1f} KB")
    
    # Guardar también los datos JSON
    with open('docs/data.json', 'w', encoding='utf-8') as f:
        json.dump(chart_data, f, ensure_ascii=False, indent=2)
    
    print("=" * 50)
    print("✅ PROCESO COMPLETADO")
    print("=" * 50)


if __name__ == '__main__':
    main()
