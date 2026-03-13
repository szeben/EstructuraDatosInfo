
import re
import os

def parse_ocr(text):
    tables = {}
    current_table = None
    
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        # Match table name AMCOR.NAME
        table_match = re.match(r'AMCOR\.([A-Z0-9_]+)', line)
        if table_match:
            current_table = table_match.group(1)
            tables[current_table] = {'columns': [], 'pks': [], 'fks': []}
            continue
            
        if not current_table:
            continue
            
        # Match columns (usually have types like VARCHAR2, NUMBER, DATE, etc.)
        # Example: PF* CIA_CODCIA VARCHAR2 (4 BYTE)
        col_match = re.search(r'([PF\* ]+)([A-Z0-9_]+)\s+([A-Z]+[0-9]*\s*\(.*?\)|[A-Z]+[0-9]*\b)', line)
        if col_match:
            flags = col_match.group(1).strip()
            name = col_match.group(2)
            dtype = col_match.group(3)
            tables[current_table]['columns'].append({
                'name': name,
                'type': dtype,
                'pk': 'P' in flags or 'PF' in flags,
                'fk': 'F' in flags or 'PF' in flags
            })
            continue
            
        # Match Primary Keys
        if '_PK (' in line or line.endswith('_PK'):
            tables[current_table]['pks'].append(line)
            continue
            
        # Match Foreign Keys
        if 'FK_' in line:
            tables[current_table]['fks'].append(line)
            continue
            
    return tables

def get_module(table_name):
    if table_name.startswith('NM'): return 'Nómina'
    if table_name.startswith('SPI') or table_name.startswith('SS'): return 'Seguridad y Sistema'
    if table_name.startswith('EO'): return 'Estructura Organizativa'
    if table_name.startswith('AR'): return 'Configuración y Archivos'
    if table_name.startswith('RH'): return 'Recursos Humanos'
    if table_name.startswith('TA'): return 'Tablas Maestras'
    if table_name.startswith('RS'): return 'Reclutamiento y Selección'
    return 'Otros'

def generate_html(tables, output_dir):
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    modules = {}
    for tname, info in tables.items():
        mod = get_module(tname)
        if mod not in modules:
            modules[mod] = []
        modules[mod].append((tname, info))
        
    # Styles
    css = """
    :root {
        --primary: #2563eb;
        --secondary: #64748b;
        --bg: #f8fafc;
        --card-bg: #ffffff;
        --text: #1e293b;
        --border: #e2e8f0;
    }
    body { font-family: 'Inter', sans-serif; background: var(--bg); color: var(--text); line-height: 1.6; margin: 0; padding: 20px; }
    header { background: linear-gradient(135deg, #1e293b, #334155); color: white; padding: 40px 20px; border-radius: 12px; margin-bottom: 30px; text-align: center; }
    nav { margin-bottom: 30px; display: flex; gap: 10px; flex-wrap: wrap; justify-content: center; }
    nav a { text-decoration: none; background: white; padding: 10px 20px; border-radius: 8px; border: 1px solid var(--border); color: var(--text); transition: all 0.2s; }
    nav a:hover { border-color: var(--primary); color: var(--primary); box-shadow: 0 4px 6px -1px rgb(0 0 0 / 0.1); }
    .module-card { background: var(--card-bg); padding: 20px; border-radius: 12px; border: 1px solid var(--border); margin-bottom: 20px; box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.1); }
    .table-detail { margin-top: 20px; border-collapse: collapse; width: 100%; border-radius: 8px; overflow: hidden; }
    .table-detail th, .table-detail td { padding: 12px; text-align: left; border-bottom: 1px solid var(--border); }
    .table-detail th { background: #f1f5f9; font-weight: 600; }
    .badge { padding: 2px 8px; border-radius: 4px; font-size: 0.75rem; font-weight: 600; }
    .badge-pk { background: #fee2e2; color: #991b1b; }
    .badge-fk { background: #dcfce7; color: #166534; }
    h2 { border-left: 4px solid var(--primary); padding-left: 15px; margin-top: 40px; }
    """
    
    # Index page
    index_content = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Documentación Modelo ER Nómina</title>
        <style>{css}</style>
    </head>
    <body>
        <header>
            <h1>Modelo de Entidad Relación - Sistema de Nómina</h1>
            <p>Diccionario de Datos y Documentación Técnica</p>
        </header>
        <nav>
            <a href="index.html">Inicio</a>
            {''.join([f'<a href="modulo_{m.lower().replace(" ", "_")}.html">{m}</a>' for m in modules.keys()])}
            <a href="diccionario_completo.html">Diccionario Completo</a>
        </nav>
        <div class="module-card">
            <h2>Resumen del Sistema</h2>
            <p>Este documento contiene la especificación detallada de las tablas que componen el sistema de gestión de Nómina. El modelo está normalizado y dividido en módulos funcionales para facilitar su comprensión.</p>
            <ul>
                {''.join([f'<li><strong>{m}:</strong> {len(items)} tablas</li>' for m, items in modules.items()])}
            </ul>
        </div>
    </body>
    </html>
    """
    with open(os.path.join(output_dir, 'index.html'), 'w', encoding='utf-8') as f:
        f.write(index_content)
        
    # Module pages
    for mod, items in modules.items():
        mod_filename = f"modulo_{mod.lower().replace(' ', '_')}.html"
        tables_html = ""
        for tname, info in items:
            cols_html = ""
            for col in info['columns']:
                pk_badge = '<span class="badge badge-pk">PK</span>' if col['pk'] else ''
                fk_badge = '<span class="badge badge-fk">FK</span>' if col['fk'] else ''
                cols_html += f"<tr><td>{col['name']}</td><td>{col['type']}</td><td>{pk_badge} {fk_badge}</td></tr>"
                
            tables_html += f"""
            <div id="{tname}" class="module-card">
                <h3>Tabla: {tname}</h3>
                <table class="table-detail">
                    <thead><tr><th>Campo</th><th>Tipo de Dato</th><th>Constraint</th></tr></thead>
                    <tbody>{cols_html}</tbody>
                </table>
                <p><strong>Relaciones:</strong> {', '.join(info['pks'] + info['fks']) if (info['pks'] or info['fks']) else 'Ninguna'}</p>
            </div>
            """
            
        mod_content = f"""
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>{mod} - Documentación</title>
            <style>{css}</style>
        </head>
        <body>
            <header>
                <h1>Módulo: {mod}</h1>
                <a href="index.html" style="color: white; text-decoration: none;">&larr; Volver al Inicio</a>
            </header>
            <nav>
                {''.join([f'<a href="modulo_{m.lower().replace(" ", "_")}.html">{m}</a>' for m in modules.keys()])}
            </nav>
            {tables_html}
        </body>
        </html>
        """
        with open(os.path.join(output_dir, mod_filename), 'w', encoding='utf-8') as f:
            f.write(mod_content)

    # Full dictionary
    dict_content = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <title>Diccionario de Datos Completo</title>
        <style>{css}</style>
    </head>
    <body>
        <header>
            <h1>Diccionario de Datos Completo</h1>
            <a href="index.html" style="color: white; text-decoration: none;">&larr; Volver al Inicio</a>
        </header>
        <table class="table-detail">
            <thead><tr><th>Tabla</th><th>Módulo</th><th>Campos</th></tr></thead>
            <tbody>
                {''.join([f'<tr><td><a href="modulo_{get_module(t).lower().replace(" ", "_")}.html#{t}">{t}</a></td><td>{get_module(t)}</td><td>{len(info["columns"])}</td></tr>' for t, info in tables.items()])}
            </tbody>
        </table>
    </body>
    </html>
    """
    with open(os.path.join(output_dir, 'diccionario_completo.html'), 'w', encoding='utf-8') as f:
        f.write(dict_content)

if __name__ == "__main__":
    from sys import stdin
    ocr_text = stdin.read()
    tables = parse_ocr(ocr_text)
    generate_html(tables, "documentacion_nomina")
