import XLSX from 'xlsx';

// Carrega arquivo de origem

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/Reckitt/Bases Originais/BISTEK/2022/Bistek setembro 22.xlsx'
);
const ws1 = wb1.Sheets['Base limpeza setembro 2022'];
let headerMes = 0;
let headerAno = 0;

const range = XLSX.utils.decode_range(ws1['!ref']);

for (let c = range.s.c; c <= range.e.c; c++) {
  const cellAddress = XLSX.utils.encode_cell({ r: 3, c });
  const cell = ws1[cellAddress];
  if (cell) {
    const cellValue = cell.v.toString().toLocaleUpperCase();
    const cellPosition = XLSX.utils.decode_cell(cellAddress);
    console.log(cellValue);
    const anos = ['2020', '2021', '2022', '2023', '2024'];
    const meses = [
      'JAN',
      'FEV',
      'MAR',
      'ABR',
      'MAI',
      'JUN',
      'JUL',
      'AGO',
      'SET',
      'OUT',
      'NOV',
      'DEZ',
    ];
    if (meses.includes(cellValue)) {
      headerMes = cellPosition.c;
    }
    if (anos.includes(cellValue)) {
      headerAno = cellPosition.c + 1;
    }
  }
}

function getMes(headerMes) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerMes })
    .map((row) => row[headerMes].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    column.map((value) => [value]),
    {
      origin: 'M1',
    }
  );
}
function getAno(headerAno) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerAno })
    .map((row) => row[headerAno - 1]);
  console.log(column);
  XLSX.utils.sheet_add_aoa(
    ws2,
    column.map((value) => [value]),
    { origin: 'P1' }
  );
}

// Carrega o arquivo de destino
const destino = 'C:/Users/leona/Desktop/TESTANTO_MODELO.xlsx';
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

getMes(headerMes);
getAno(headerAno);
    


XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
XLSX.writeFile(wb2, destino);
