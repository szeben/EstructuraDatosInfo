
const fs = require('fs');
const XLSX = require('xlsx');

// 1. ADVANCED RELATIONSHIP MAPPING
// This map defines common ID patterns and their "Master" target table
const relationshipMap = {
    'ID_EMPRESA': 'EO_EMPRESA',
    'CIA_CODCIA': 'EO_EMPRESA',
    'ID_PERSONA': 'EO_PERSONA',
    'FICHA': 'TA_RELACION_LABORAL',
    'TRAB_FICTRA': 'TA_RELACION_LABORAL',
    'ID_TIPO_IDEN': 'EO_TIPO_IDENTIFICACION',
    'ID_PAIS': 'SPI_PAISES',
    'ID_LOCALIDAD': 'EO_LOCALIDAD',
    'ID_CARGO': 'EO_CARGO',
    'CODCAR': 'EO_CARGO',
    'ID_UNIDAD': 'EO_UNIDAD',
    'CODDEP': 'EO_UNIDAD',
    'DPTO_CODDEP': 'EO_UNIDAD',
    'ID_NOMINA': 'NMT003',
    'TNOM_TIPNOM': 'NMT003',
    'CTO_CODCTO': 'NMT027',
    'ID_PROCESO': 'NMT011',
    'PROC_TIPPRO': 'NMT011',
    'ID_PERFIL': 'SS_PERFIL',
    'ID_USUARIO': 'SPI_USUARIO',
    'ID_AUT_EMP': 'SS_AUTORIZA_EMPRESA',
    'ID_BANDA': 'RS_BANDA',
    'ID_COMPETENCIA': 'RS_COMPETENCIA',
    'ID_CV': 'RS_CV',
    'ID_REQUISICION': 'RS_REQUISICION',
    'ID_RESUMEN': 'RS_RESUMEN',
    'ID_PARENTESCO': 'TA_PARENTESCOS',
    'ID_PARIENTE': 'EO_PARIENTE',
    'ID_BENEFICIARIO': 'TA_BENEFICIARIOS',
    'CODE_PLAN': 'TA_PRIMA_PLAN',
    'CODE_POLIZA': 'TA_POLIZAS'
};

// 2. OCR-BASED CONSTRAINTS (Expanded)
const ocrContent = `
AMCOR.NMT037: TXTR_PK (CIA_CODCIA, TNOM_TIPNOM, PROC_TIPPRO, FPRO_SUBPRO, FPRO_ANOCAL, FPRO_NUMPER, SUBTIP, FICTRA)
AMCOR.SPI_USUARIO: PK_USU (ID) | FK_USUARIO_PAIS (ID_PAIS) | FK_USU_PERFIL (ID_PERFIL) | FK_USU_PERSONA (ID_PERSONA)
AMCOR.TA_RELACION_LABORAL: LABORAL_PK (ID_EMPRESA, FICHA) | LABORAL_PERSONA (ID_PERSONA) | TRAB_LOCA (ID_LOCALIDAD) | TRAB_SNDC (ID_SINDICATO)
AMCOR.EO_CARGO: PK_CGO (ID_EMPRESA, ID) | FK_CARGO_CLASI (ID_CLASIFICA) | FK_JERARQUIA_CARGO (ID_CARGO_SUP)
AMCOR.NMM019: MVAR_PK (CIA_CODCIA, TNOM_TIPNOM, PROC_TIPPRO, FPRO_SUBPRO, FPRO_ANOCAL, FPRO_NUMPER, USERMO, MODMOV, SECMOV, SUBSEC) | MVAR_CTO (CTO_CODCTO) | MVAR_TRAB_FICTRA (TRAB_FICTRA)
AMCOR.EO_PERSONA: PK_PERSON (ID) | FK_PERSONA_TIPO_ID (ID_TIPO_IDEN)
AMCOR.EO_EMPRESA: PK_EMPRE (ID)
AMCOR.SS_ROL_OPCION: PK_ROL_OPC (ID_ROL, ID_SISTEMA, ID_MODULO, ID_OPCION) | FK_ROL_OPC_OPC (ID_OPCION) | FK_ROL_OPC_ROL (ID_ROL)
AMCOR.NM_RELACION_PAGO: REL_PAGO_PK (ID_EMPRESA, FICHA) | LABORAL_NOMINA (ID_NOMINA) | TRAB_PROCESOS (ID_PROCESO)
AMCOR.NMT027: CTO_PK (CIA_CODCIA, TNOM_TIPNOM, CODCTO) | CTO_PROM (PROM_CODPRO)
AMCOR.EO_UNIDAD: PK_UNI (ID_EMPRESA, ID) | FK_JERARQUIA_UNIDAD (ID_UNIDAD_SUP)
`;

function getPKsForTable(tname) {
    if (tname.includes('PERSONA')) return ['ID'];
    if (tname.includes('EMPRESA')) return ['ID'];
    if (tname === 'NMT003') return ['TIPNOM'];
    if (tname === 'NMT027') return ['CODCTO'];
    // Default patterns from Infocent usually prefix or ID
    return [];
}

// 3. LOAD EXCEL & BUILD TABLES
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
        allTables[currentT] = { columns: [] };
    }
    const cname = row[2] ? String(row[2]).trim() : '';
    if (cname && currentT) {
        allTables[currentT].columns.push({
            name: cname,
            type: row[3] || '',
            len: row[4] || '',
            desc: row[6] || ''
        });
    }
}

// 4. GENERATE HTML WITH INTERLINKS
const names = Object.keys(allTables).sort();
let tableCards = "";

for (const tname of names) {
    const info = allTables[tname];
    const tablePKs = getPKsForTable(tname);
    
    let rows = info.columns.map(c => {
        const targetTable = relationshipMap[c.name] || null;
        const isPk = tablePKs.includes(c.name) || c.name === 'ID' && (tname.startsWith('EO_') || tname.startsWith('SPI_'));
        
        let colDisplay = `<span class="col-name">${c.name}</span>`;
        let constraintBadge = "";
        
        if (isPk) {
            constraintBadge = `<span class="pk-badge">PK</span>`;
        }
        
        if (targetTable && targetTable !== tname) {
            // Check if target table actually exists in our dictionary
            const finalTarget = allTables[targetTable] ? targetTable : null;
            if (finalTarget) {
                colDisplay = `<a href="#${finalTarget}" class="fk-link"><span class="col-name">${c.name}</span></a>`;
                constraintBadge += `<span class="fk-badge">FK &rarr; ${finalTarget}</span>`;
            }
        }

        return `
            <tr class="${isPk ? 'pk-row' : ''}">
                <td>${colDisplay} ${constraintBadge}</td>
                <td>${c.type}</td>
                <td>${c.len}</td>
                <td class="desc-cell">${c.desc}</td>
            </tr>
        `;
    }).join("");

    tableCards += `
    <div class="table-card" id="${tname}" data-name="${tname}">
        <div class="table-header-flex">
            <h3>AMCOR.${tname}</h3>
            <a href="#top" style="font-size:0.6rem; color: var(--accent); text-decoration:none;">&uarr; Ir arriba</a>
        </div>
        <table class="data-table">
            <thead><tr><th>Campo & Relación</th><th>Tipo</th><th>Long</th><th>Descripción</th></tr></thead>
            <tbody>${rows}</tbody>
        </table>
    </div>
    `;
}

const css = `
:root {
    --primary: #6366f1; --bg-dark: #0f172a; --card-bg: #1e293b;
    --text-main: #f8fafc; --text-muted: #94a3b8; --accent: #38bdf8;
    --border: rgba(255, 255, 255, 0.1);
}
body { font-family: 'Inter', sans-serif; background: var(--bg-dark); color: var(--text-main); font-size: 0.7rem; margin: 0; scroll-behavior: smooth; }
.container { max-width: 1400px; margin: 0 auto; padding: 1rem; }
header { padding: 1rem; border-bottom: 2px solid var(--primary); text-align: center; }
.fk-link { text-decoration: none; border-bottom: 1px dashed var(--accent); }
.fk-link:hover { background: rgba(56, 189, 248, 0.2); }
.search-container { position: sticky; top: 0; background: var(--bg-dark); padding: 0.5rem 0; z-index: 1000; border-bottom: 1px solid var(--border); }
.search-box { width: 100%; padding: 0.4rem; background: var(--card-bg); color: white; border: 1px solid var(--border); border-radius: 4px; }
.table-card { background: var(--card-bg); border-radius: 6px; padding: 0.8rem; margin-bottom: 2rem; border: 1px solid var(--border); }
.table-card:target { border-color: var(--accent); box-shadow: 0 0 15px var(--accent); }
.data-table { width: 100%; border-collapse: collapse; }
.data-table th { text-align: left; padding: 0.4rem; background: rgba(99, 102, 241, 0.15); color: var(--accent); font-weight: 800; border-bottom: 1px solid var(--border); }
.data-table td { padding: 0.3rem; border-bottom: 1px solid var(--border); }
.col-name { color: #facc15; font-weight: 600; }
.pk-row { background: rgba(248, 113, 113, 0.05); }
.pk-badge { font-size: 0.55rem; background: #ef4444; color: white; padding: 1px 3px; border-radius: 3px; margin-left: 4px; }
.fk-badge { font-size: 0.55rem; background: #10b981; color: white; padding: 1px 3px; border-radius: 3px; margin-left: 4px; }
.desc-cell { color: var(--text-muted); }
`;

const html = `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>GDD Navegable - Infocent</title>
    <style>${css}</style>
</head>
<body>
    <div id="top"></div>
    <header>
        <h1 style="margin:0;">DICCIONARIO NAVEGABLE (PK/FK LINKS)</h1>
        <p style="font-size:0.6rem; color:var(--text-muted);">Haz clic en los campos subrayados para saltar a la tabla Maestra</p>
    </header>
    <div class="container">
        <nav><a href="index.html">Inicio</a> <a href="relaciones.html">Volver</a></nav>
        <div class="search-container">
            <input type="text" id="masterSearch" class="search-box" placeholder="Buscar tabla o campo..." onkeyup="filter()">
        </div>
        <div id="tableList">${tableCards}</div>
    </div>
    <script>
        function filter() {
            let q = document.getElementById('masterSearch').value.toUpperCase();
            let cards = document.getElementsByClassName('table-card');
            for(let card of cards) {
                card.style.display = card.innerText.toUpperCase().includes(q) ? "block" : "none";
            }
        }
    </script>
</body>
</html>`;

fs.writeFileSync('b:/Antigravity/BD_Infocent/docs/diccionario_excel.html', html);
console.log("Diccionario enriquecido e interconectado generado.");
