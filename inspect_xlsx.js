
const XLSX = require('xlsx');

try {
    const workbook = XLSX.readFile('diccionario_datos_spi.xlsx');
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet);
    console.log(JSON.stringify(data.slice(0, 5), null, 2));
} catch (e) {
    console.error(e.message);
}
