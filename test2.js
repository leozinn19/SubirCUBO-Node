import XLSX from 'xlsx';

// Carrega arquivo de origem

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/Reckitt/Bases Originais/BISTEK/2022/Bistek setembro 22.xlsx'
);
const ws1 = wb1.Sheets['Base limpeza setembro 2022'];

function iterateFourthRow(sheet, callback) {
  const range = { s: { r: 3 }, e: { r: 3 } }; // lê somente a quarta linha
  const rows = XLSX.utils.sheet_to_json(sheet, { range });
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const keys = Object.keys(row);
    for (let j = 0; j < keys.length; j++) {
      const cellValue = row[keys[j]];
      callback(cellValue, i + 4, j + 1); // chama a função de callback para cada célula
    }
  }
}

iterateFourthRow(ws1, (value, row, col) => {
  console.log(`Valor na posição (${row}, ${col}): ${value}`);
});

// Carrega o arquivo de destino
const destino = 'C:/Users/leona/Desktop/TESTANTO_MODELO.xlsx';
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
XLSX.writeFile(wb2, destino);
