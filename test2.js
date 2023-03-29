import XLSX from 'xlsx';

// Carrega arquivo de origem

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/Reckitt/Bases Originais/BISTEK/2022/Bistek setembro 22.xlsx'
);
const ws1 = wb1.Sheets['Base limpeza setembro 2022'];
let headerN = 0;

const range = XLSX.utils.decode_range(ws1['!ref']);

for (let c = range.s.c; c <= range.e.c; c++) {
  const cellAddress = XLSX.utils.encode_cell({ r: 3, c });
  const cell = ws1[cellAddress];
  if (cell) {
    const cellValue = cell.v.toString().toLocaleUpperCase();
    const cellPosition = XLSX.utils.decode_cell(cellAddress);
    if (cellValue === 'SET') {
      headerN = cellPosition.c + 1;
    }
  }
}

function getMes(headerN) {
  console.log(headerN);
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerN - 1 })
    .map((row) => row[headerN - 1].toUpperCase());
  console.log(column);
  XLSX.utils.sheet_add_aoa(
    ws2,
    column.map((value) => [value]),
    {
      origin: 'M1',
    }
  );
}

// Carrega o arquivo de destino
const destino = 'C:/Users/leona/Desktop/TESTANTO_MODELO.xlsx';
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

getMes(headerN);
    


XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
XLSX.writeFile(wb2, destino);
