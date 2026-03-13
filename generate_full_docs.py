
import re
import os
import json

def parse_ocr_to_json(text):
    tables = {}
    current_table = None
    
    # Clean OCR artifacts (common misreadings)
    text = text.replace('PF* ', 'PF*').replace('P * ', 'P*').replace('F * ', 'F*')
    
    lines = text.split('\n')
    for line in lines:
        line = line.strip()
        if not line: continue
            
        # Match table name AMCOR.NAME
        table_match = re.match(r'AMCOR\.([A-Z0-9_]+)', line)
        if table_match:
            current_table = table_match.group(1)
            tables[current_table] = {'columns': [], 'pks': [], 'fks': []}
            continue
            
        if not current_table: continue
            
        # Match columns
        # Format: [Flags] NAME TYPE
        col_match = re.search(r'^([PF\*]+)?\s*([A-Z0-9_]+)\s+([A-Z]+[0-9]*\s*\(.*?\)|[A-Z]+[0-9]*\b)', line)
        if col_match:
            flags = col_match.group(1) or ""
            name = col_match.group(2)
            dtype = col_match.group(3)
            
            # Avoid matching constraint definitions as columns
            if name in ['PK', 'FK'] or dtype.startswith('('):
                continue

            tables[current_table]['columns'].append({
                'name': name,
                'type': dtype,
                'pk': 'P' in flags,
                'fk': 'F' in flags
            })
            continue
            
        # Capture Constraints (Simplified)
        if '_PK (' in line:
            tables[current_table]['pks'].append(line)
        elif 'FK_' in line:
            tables[current_table]['fks'].append(line)
            
    return tables

def get_module_data(tname):
    if tname.startswith('NM'): return 'Nómina', '📊'
    if tname.startswith('SPI') or tname.startswith('SS'): return 'Seguridad', '🛡️'
    if tname.startswith('EO'): return 'Estructura', '🏢'
    if tname.startswith('AR'): return 'Configuración', '⚙️'
    if tname.startswith('RH'): return 'RRHH', '📜'
    if tname.startswith('TA'): return 'Maestras', '📋'
    if tname.startswith('RS'): return 'Reclutamiento', '👥'
    return 'Otros', '📦'

def generate_full_dictionary(tables_data, output_file):
    html_template = """
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Diccionario Completo de Datos - Infocent</title>
    <link rel="stylesheet" href="style.css">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;800&display=swap" rel="stylesheet">
    <style>
        .search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 1rem 0; z-index: 100; border-bottom: 1px solid var(--border); }
        .table-list { margin-top: 2rem; }
        .table-card { margin-bottom: 4rem; scroll-margin-top: 100px; }
        .module-badge { font-size: 0.7rem; padding: 2px 8px; border-radius: 20px; background: var(--primary); margin-left:10px; }
    </style>
</head>
<body>
    <header><h1>Diccionario Universal</h1></header>
    <div class="container">
        <nav><a href="index.html">Inicio</a> <a href="diccionario.html">Volver</a></nav>
        
        <div class="search-container">
            <input type="text" id="masterSearch" class="search-box" placeholder="Filtrar entre {count} tablas..." onkeyup="filterAll()">
        </div>

        <div class="table-list" id="mainList">
            {table_blocks}
        </div>
    </div>
    <script>
        function filterAll() {{
            let q = document.getElementById('masterSearch').value.toUpperCase();
            let cards = document.getElementsByClassName('table-card');
            for(let card of cards) {{
                card.style.display = card.innerText.toUpperCase().includes(q) ? "" : "none";
            }}
        }}
    </script>
</body>
</html>
    """
    
    table_blocks = []
    for tname, info in sorted(tables_data.items()):
        mod, icon = get_module_data(tname)
        rows = []
        for col in info['columns']:
            pk_class = 'class="pk"' if col['pk'] else ''
            fk_class = 'class="fk"' if col['fk'] else ''
            rows.append(f"<tr><td><span {pk_class or fk_class}>{col['name']}</span></td><td>{col['type']}</td><td>{ 'PK' if col['pk'] else '' } { 'FK' if col['fk'] else '' }</td></tr>")
        
        block = f"""
        <div class="table-card table-item active" id="{tname}" data-mod="{mod}">
            <h3>{icon} {tname} <span class="module-badge">{mod}</span></h3>
            <table class="data-table">
                <thead><tr><th>Columna</th><th>Tipo</th><th>Atributos</th></tr></thead>
                <tbody>{''.join(rows)}</tbody>
            </table>
            <div style="font-size: 0.8rem; color: var(--text-muted); margin-top:10px;">
                <strong>Constraints:</strong> { ' | '.join(info['pks'] + info['fks']) or 'N/A' }
            </div>
        </div>
        """
        table_blocks.append(block)
        
    full_html = html_template.format(
        count=len(tables_data),
        table_blocks='\n'.join(table_blocks)
    )
    
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(full_html)

if __name__ == "__main__":
    import sys
    # Read the OCR text from a file if provided, or stdin
    content = sys.stdin.read()
    data = parse_ocr_to_json(content)
    generate_full_dictionary(data, "b:/Antigravity/BD_Infocent/docs/diccionario_total.html")
    print(f"Generadas {len(data)} tablas.")
