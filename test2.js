import XLSX from 'xlsx';

// Carrega arquivo de origem

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/Reckitt/Bases Originais/BISTEK/2022/Bistek setembro 22.xlsx'
);
const ws1 = wb1.Sheets['Base limpeza setembro 2022'];
let headerMes = 0;
let headerAno = 0;

const range = XLSX.utils.decode_range(ws1['!ref']);

// BUSCA VELOR DE CÉLUNAS
for (let c = range.s.c; c <= range.e.c; c++) {
  const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
  const cell = ws1[cellAddress];
  if (cell) {
    const cellValue = cell.v.toString().toLocaleUpperCase();
    const cellPosition = XLSX.utils.decode_cell(cellAddress);
    console.log(cellValue);
    if (cellValue === 'MÊS' || cellValue === 'MES') {
      headerMes = cellPosition.c;
    }
    else if (cellValue === 'ANO') {
      headerAno = cellPosition.c + 1;
    }
    else if (cellValue === 'EAN') {
      headerAno = cellPosition.c + 1;
    }
  }
}
// REF VAREJISTA
function refVarejista(ref) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'A1' })
    .map((row) => row['A1']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['REF VAREJISTA']].concat(column.map((value) => [ref])),
    {
      origin: 'A1',
    }
  );
}
// DIA
function getDia(headerMes, headerAno) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerMes })
    .map((row) => row[headerMes].toUpperCase());
  const column2 = XLSX.utils
    .sheet_to_json(ws1, { header: headerAno })
    .map((row) => row[headerAno - 1]);

  const concatenated = column.map((value, index) => {
    '01/' + value + column2[index];
    if (value === 'JAN') return '01/01/' + column2[index];
    else if (value === 'FEV') return '01/02/' + column2[index];
    else if (value === 'MAR') return '01/03/' + column2[index];
    else if (value === 'ABR') return '01/04/' + column2[index];
    else if (value === 'MAI') return '01/05/' + column2[index];
    else if (value === 'JUN') return '01/06/' + column2[index];
    else if (value === 'JUL') return '01/07/' + column2[index];
    else if (value === 'AGO') return '01/08/' + column2[index];
    else if (value === 'SET') return '01/09/' + column2[index];
    else if (value === 'OUT') return '01/10/' + column2[index];
    else if (value === 'NOV') return '01/11/' + column2[index];
    else if (value === 'DEZ') return '01/12/' + column2[index];
  });
  XLSX.utils.sheet_add_aoa(
    ws2,
    (concatenated.map((value) => [value])),
    { origin: 'K1' }
  );
}
// MES
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
// DECR_MES
function descrMes(headerMes) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerMes })
    .map((row) => row[headerMes].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    (
      column.map((value) => {
        let result = [value].toString();
        if (result === 'JAN') return ['01-JAN'];
        else if (result === 'FEV') return ['02-FEV'];
        else if (result === 'MAR') return ['03-MAR'];
        else if (result === 'ABR') return ['04-ABR'];
        else if (result === 'MAI') return ['05-MAI'];
        else if (result === 'JUN') return ['06-JUN'];
        else if (result === 'JUL') return ['07-JUL'];
        else if (result === 'AGO') return ['08-AGO'];
        else if (result === 'SET') return ['09-SET'];
        else if (result === 'OUT') return ['10-OUT'];
        else if (result === 'NOV') return ['11-NOV'];
        else if (result === 'DEZ') return ['12-DEZ'];
      })
    ),
    { origin: 'N1' }
  );
    }
// ANO
function getAno(headerAno) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerAno })
    .map((row) => row[headerAno - 1]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    column.map((value) => [value]),
    { origin: 'P1' }
  );
}
// PERIODO
function getPeriodo(headerMes) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerMes })
    .map((row) => row[headerMes].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    (column.map((value) => {
        let result = [value].toString();
        if (result === 'JAN' || result === 'FEV') return ['1 BIM'];
        else if (result === 'MAR' || result === 'ABR') return ['2 BIM'];
        else if (result === 'MAI' || result === 'JUN') return ['3 BIM'];
        else if (result === 'JUL' || result === 'AGO') return ['4 BIM'];
        else if (result === 'SET' || result === 'OUT') return ['5 BIM'];
        else if (result === 'NOV' || result === 'DEZ') return ['6 BIM'];
      })
    ),
    { origin: 'Q1' }
  );
  }

// CONCAT
function concat(headerMes, headerAno) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerMes })
    .map((row) => row[headerMes].toUpperCase());
  const column2 = XLSX.utils
    .sheet_to_json(ws1, { header: headerAno})
    .map((row) => row[headerAno - 1]);

  const concatenated = column.map((value, index) => value + column2[index]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    concatenated.map((value) => [value]),
    { origin: 'W1' }
  );
}

// CARREGA ARQUIVO DESTINO 
const destino = 'C:/Users/leona/Desktop/TESTANTO_MODELO.xlsx';
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

getDia(headerMes, headerAno);
getMes(headerMes);
descrMes(headerMes);
getAno(headerAno);
concat(headerMes, headerAno);
getPeriodo(headerMes);

// FUNÇÔES QUE PRECISAM SER COMPLETAS DE ACORDO COM O CLIENTE:
refVarejista('BISTEK')

// SALVA ARQUIVO DESTINO
XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
XLSX.writeFile(wb2, destino);
