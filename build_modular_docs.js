
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// 1. CONFIGURATION & MAPPING
const outputDir = 'b:/Antigravity/BD_Infocent/docs';
const relationshipMap = {
    'ID_EMPRESA': 'EO_EMPRESA', 'CIA_CODCIA': 'EO_EMPRESA', 'CIA_CODACT': 'EO_EMPRESA', 'CIAANT': 'EO_EMPRESA',
    'ID_PERSONA': 'EO_PERSONA', 'ID_PERSONA_HR': 'EO_PERSONA',
    'FICHA': 'TA_RELACION_LABORAL', 'TRAB_FICTRA': 'TA_RELACION_LABORAL', 'FICANT': 'TA_RELACION_LABORAL', 'FICMBA': 'TA_RELACION_LABORAL',
    'ID_TIPO_IDEN': 'EO_TIPO_IDENTIFICACION', 'ID_PAIS': 'SPI_PAISES', 'CODPAI': 'SPI_PAISES',
    'ID_LOCALIDAD': 'EO_LOCALIDAD', 'CODLOC': 'EO_LOCALIDAD',
    'ID_CARGO': 'EO_CARGO', 'CODCAR': 'EO_CARGO', 'CGO_CODCAR': 'EO_CARGO',
    'ID_UNIDAD': 'EO_UNIDAD', 'CODDEP': 'EO_UNIDAD', 'DPTO_CODDEP': 'EO_UNIDAD', 'CODUNI': 'EO_UNIDAD',
    'ID_NOMINA': 'NMT003', 'TNOM_TIPNOM': 'NMT003', 'TIPNOM': 'NMT003',
    'CTO_CODCTO': 'NMT027', 'CODCTO': 'NMT027', 'CTOMBA': 'NMT027',
    'ID_PROCESO': 'NMT011', 'PROC_TIPPRO': 'NMT011', 'TIPPRO': 'NMT011',
    'ID_PERFIL': 'SS_PERFIL', 'ID_USUARIO': 'SPI_USUARIO', 'USERID': 'SPI_USUARIO', 'USR_USERID': 'SPI_USUARIO',
    'ID_BANDA': 'RS_BANDA', 'ID_COMPETENCIA': 'RS_COMPETENCIA', 'ID_CV': 'RS_CV',
    'ID_REQUISICION': 'RS_REQUISICION', 'ID_RESUMEN': 'RS_RESUMEN',
    'ID_PARENTESCO': 'TA_PARENTESCOS', 'ID_PARIENTE': 'EO_PARIENTE',
    'ID_BENEFICIARIO': 'TA_BENEFICIARIOS', 'CODBEN': 'TA_BENEFICIARIOS',
    'CODE_PLAN': 'TA_PRIMA_PLAN', 'CODE_POLIZA': 'TA_POLIZAS', 'CODE_FINANCIA': 'TA_FINANCIAMIENTO_SEGURO'
};

const moduleDefs = [
    { id: 'nm', name: 'Nómina', prefix: 'NM', icon: '📊', desc: 'Procesos de nómina, cálculos y movimientos.' },
    { id: 'eo', name: 'Estructura', prefix: 'EO', icon: '🏢', desc: 'Jerarquías, empresas, unidades y cargos.' },
    { id: 'spi', name: 'Seguridad', prefix: ['SPI', 'SS'], icon: '🛡️', desc: 'Usuarios, roles, perfiles y auditoría.' },
    { id: 'rs', name: 'Reclutamiento', prefix: 'RS', icon: '👥', desc: 'Candidatos, vacantes y evaluaciones.' },
    { id: 'rh', name: 'RRHH', prefix: 'RH', icon: '📜', desc: 'Históricos, capacitación y familiares.' },
    { id: 'ta', name: 'Maestras', prefix: 'TA', icon: '📋', desc: 'Seguros, guarderías y tablas maestras.' }
];

const sharedCss = `
:root {
    --primary: #6366f1; --bg-dark: #0f172a; --card-bg: #1e293b;
    --text-main: #f8fafc; --text-muted: #94a3b8; --accent: #38bdf8;
    --border: rgba(255, 255, 255, 0.1);
}
* { box-sizing: border-box; }
body { font-family: 'Inter', sans-serif; background: var(--bg-dark); color: var(--text-main); font-size: 0.75rem; margin: 0; scroll-behavior: smooth; line-height: 1.4; }
.container { max-width: 1400px; margin: 0 auto; padding: 1rem; }
header { padding: 2rem 1rem; border-bottom: 2px solid var(--primary); text-align: center; background: radial-gradient(circle at top, #1e293b 0%, #0f172a 100%); }
h1 { font-size: 1.6rem; color: var(--accent); margin:0; text-transform: uppercase; letter-spacing: 1px; }
nav { display: flex; gap: 0.8rem; justify-content: center; padding: 1rem; flex-wrap: wrap; }
nav a { color: var(--text-main); text-decoration: none; padding: 0.4rem 1rem; border-radius: 6px; background: rgba(255,255,255,0.05); border: 1px solid var(--border); transition: 0.2s; }
nav a:hover { background: var(--primary); border-color: var(--primary); transform: translateY(-2px); }

.module-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(350px, 1fr)); gap: 1.5rem; margin-top: 2rem; }
.module-card { background: var(--card-bg); border-radius: 12px; padding: 1.5rem; border: 1px solid var(--border); transition: 0.3s; cursor: pointer; text-decoration: none; color: inherit; display: block; }
.module-card:hover { transform: translateY(-5px); border-color: var(--primary); box-shadow: 0 10px 30px -10px var(--primary); }
.module-icon { font-size: 2rem; margin-bottom: 1rem; display: block; }
.module-card h3 { margin: 0.5rem 0; color: var(--accent); font-size: 1.2rem; }
.module-card p { color: var(--text-muted); font-size: 0.8rem; }
.table-count { display: inline-block; margin-top: 1rem; padding: 0.2rem 0.6rem; background: rgba(99,102,241,0.2); color: #a5b4fc; border-radius: 4px; font-weight: bold; }

.search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 1rem 0; z-index: 1000; border-bottom: 1px solid var(--border); }
.search-box { width: 100%; padding: 0.6rem; background: var(--card-bg); color: white; border: 1px solid var(--border); border-radius: 6px; font-size: 0.8rem; }

.table-card { background: var(--card-bg); border-radius: 8px; padding: 1rem; margin-bottom: 2rem; border: 1px solid var(--border); }
.table-card:target { border-color: var(--accent); box-shadow: 0 0 20px rgba(56, 189, 248, 0.3); }
.data-table { width: 100%; border-collapse: collapse; margin-top: 0.5rem; }
.data-table th { text-align: left; padding: 0.5rem; background: rgba(99, 102, 241, 0.1); color: var(--accent); font-weight: 800; border-bottom: 1px solid var(--border); text-transform: uppercase; font-size: 0.65rem; }
.data-table td { padding: 0.4rem; border-bottom: 1px solid var(--border); vertical-align: top; }
.col-name a { color: #facc15; text-decoration: none; border-bottom: 1px dotted #facc15; font-weight: 600; }
.fk-badge { font-size: 0.6rem; background: #10b981; color: white; padding: 1px 4px; border-radius: 3px; margin-left: 4px; font-weight: bold; }
.desc-cell { color: var(--text-muted); }
.ref-by { margin-top: 1rem; padding: 0.8rem; background: rgba(0,0,0,0.25); border-radius: 6px; border: 1px solid var(--border); }
.ref-links { display: flex; flex-wrap: wrap; gap: 0.5rem; margin-top: 0.5rem; }
.ref-links a { font-size: 0.65rem; color: var(--accent); text-decoration: none; padding: 3px 6px; border: 1px solid rgba(56, 189, 248, 0.2); border-radius: 4px; background: rgba(255,255,255,0.02); }
.ref-links a:hover { background: var(--accent); color: #000; }
@media (max-width: 768px) { .module-grid { grid-template-columns: 1fr; } }
`;

function getModuleId(tname) {
    for (const m of moduleDefs) {
        if (Array.isArray(m.prefix)) {
            if (m.prefix.some(p => tname.startsWith(p))) return m.id;
        } else {
            if (tname.startsWith(m.prefix)) return m.id;
        }
    }
    return 'others';
}

// 2. LOAD DATA
const workbook = XLSX.readFile('diccionario_datos_spi.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

const allTables = {};
let currentT = '';
for (let i = 2; i < rawData.length; i++) {
    const row = rawData[i];
    if (!row || row.length === 0) continue;
    const tname = row[0] ? String(row[0]).trim() : '';
    if (tname) {
        currentT = tname;
        allTables[currentT] = { columns: [], referencedBy: [], moduleId: getModuleId(currentT) };
    }
    const cname = row[2] ? String(row[2]).trim() : '';
    if (cname && currentT) {
        allTables[currentT].columns.push({
            name: cname, type: row[3] || '', len: row[4] || '', desc: row[6] || ''
        });
    }
}

// 3. CROSS REFS
for (const tname in allTables) {
    allTables[tname].columns.forEach(col => {
        const targetTable = relationshipMap[col.name];
        if (targetTable && allTables[targetTable] && targetTable !== tname) {
            if (!allTables[targetTable].referencedBy.includes(tname)) {
                allTables[targetTable].referencedBy.push(tname);
            }
        }
    });
}

// 4. GENERATION HELPERS
function generateTableHtml(tname, info, isSubpage = false) {
    const rows = info.columns.map(c => {
        const targetTable = relationshipMap[c.name];
        const isFk = targetTable && allTables[targetTable] && targetTable !== tname;
        const targetTableData = allTables[targetTable];
        if (!targetTableData) return `<tr><td class="col-name">${c.name}</td><td>${c.type}</td><td>${c.len}</td><td class="desc-cell">${c.desc}</td></tr>`;

        const targetPath = (isSubpage && targetTableData.moduleId !== info.moduleId) 
            ? `modulo_${targetTableData.moduleId}.html#${targetTable}` 
            : `#${targetTable}`;
        
        const colDisplay = isFk ? `<a href="${targetPath}">${c.name}</a>` : c.name;
        const badge = isFk ? `<span class="fk-badge">FK &rarr; ${targetTable}</span>` : "";
        return `<tr><td class="col-name">${colDisplay} ${badge}</td><td>${c.type}</td><td>${c.len}</td><td class="desc-cell">${c.desc}</td></tr>`;
    }).join("");

    const refs = info.referencedBy.length > 0 ? `
        <div class="ref-by"><strong>Referenciada por:</strong><div class="ref-links">${
            info.referencedBy.sort().map(r => {
                const refData = allTables[r];
                const rPath = (isSubpage && refData && refData.moduleId !== info.moduleId) 
                    ? `modulo_${refData.moduleId}.html#${r}` 
                    : `#${r}`;
                return `<a href="${rPath}">${r}</a>`;
            }).join("")
        }</div></div>` : "";


    return `<div class="table-card" id="${tname}" data-name="${tname}"><div style="display:flex; justify-content:space-between;"><h3>${tname}</h3><a href="#top" style="color:var(--accent); text-decoration:none; font-size:0.6rem;">&uarr; Volver</a></div><table class="data-table"><thead><tr><th>Campo</th><th>Tipo</th><th>Long</th><th>Descripción</th></tr></thead><tbody>${rows}</tbody></table>${refs}</div>`;
}

// 5. FILE WRITERS
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

// INDEX
const moduleStats = moduleDefs.map(m => {
    const count = Object.values(allTables).filter(t => t.moduleId === m.id).length;
    return `
    <a href="modulo_${m.id}.html" class="module-card">
        <span class="module-icon">${m.icon}</span>
        <h3>${m.name}</h3>
        <p>${m.desc}</p>
        <span class="table-count">${count} Tablas</span>
    </a>`;
}).join("");

const indexHtml = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Infocent Docs</title><style>${sharedCss}</style></head><body><header><h1>Sistema de Nómina Infocent</h1><p style="color:var(--text-muted);">Documentación Técnica Centralizada (Excel + Modelo ER)</p></header><div class="container"><nav><a href="index.html">Inicio</a><a href="diccionario_excel.html">Diccionario Global</a><a href="relaciones.html">Relaciones</a></nav><div class="module-grid">${moduleStats}</div></div></body></html>`;
fs.writeFileSync(path.join(outputDir, 'index.html'), indexHtml);

// MODULE PAGES
moduleDefs.forEach(m => {
    const tablesInMod = Object.keys(allTables).filter(tn => allTables[tn].moduleId === m.id).sort();
    const content = tablesInMod.map(tn => generateTableHtml(tn, allTables[tn], true)).join("");
    const modHtml = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>${m.name} - Infocent</title><style>${sharedCss}</style></head><body><div id="top"></div><header><h1>MÓDULO: ${m.name.toUpperCase()}</h1><p style="color:var(--text-muted);">${tablesInMod.length} Tablas Identificadas</p></header><div class="container"><nav><a href="index.html">&larr; Volver al Inicio</a></nav><div class="search-container"><input type="text" class="search-box" placeholder="Buscar en ${m.name}..." onkeyup="filter()"></div><div id="tableList">${content}</div></div><script>function filter(){let q=document.querySelector('.search-box').value.toUpperCase(); document.querySelectorAll('.table-card').forEach(c=>c.style.display=c.innerText.toUpperCase().includes(q)?'block':'none')}</script></body></html>`;
    fs.writeFileSync(path.join(outputDir, `modulo_${m.id}.html`), modHtml);
});

// GLOBAL DICCIONARY
const globalContent = Object.keys(allTables).sort().map(tn => generateTableHtml(tn, allTables[tn], false)).join("");
const globalHtml = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Diccionario Global</title><style>${sharedCss}</style></head><body><div id="top"></div><header><h1>DICCIONARIO TÉCNICO COMPLETO</h1></header><div class="container"><nav><a href="index.html">Inicio</a></nav><div class="search-container"><input type="text" class="search-box" placeholder="Filtrar entre ${Object.keys(allTables).length} tablas..." onkeyup="filter()"></div><div id="tableList">${globalContent}</div></div><script>function filter(){let q=document.querySelector('.search-box').value.toUpperCase(); document.querySelectorAll('.table-card').forEach(c=>c.style.display=c.innerText.toUpperCase().includes(q)?'block':'none')}</script></body></html>`;
fs.writeFileSync(path.join(outputDir, 'diccionario_excel.html'), globalHtml);

// RELACIONES (Updating to match Font Size)
const relHtml = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Relaciones - Infocent</title><style>${sharedCss}</style></head><body><header><h1>Relaciones y Reglas</h1></header><div class="container"><nav><a href="index.html">Inicio</a></nav><div class="module-card" style="cursor:default;"><h3>Análisis de Integridad</h3><p>El sistema se rige por un esquema de llaves compuestas. La mayoría de las tablas transaccionales (NMM) dependen de las tablas maestras (EO/TA) mediante vínculos FK dinámicos. El portal ahora permite saltar entre módulos de forma transparente mediante enlaces inteligentes.</p></div></div></body></html>`;
fs.writeFileSync(path.join(outputDir, 'relaciones.html'), relHtml);

console.log("Portal de documentación multi-página generado con éxito.");
