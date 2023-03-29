import XLSX from 'xlsx';

// Carrega arquivo de origem

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/Reckitt/Bases Originais/BISTEK/2022/Bistek setembro 22.xlsx'
);
const ws1 = wb1.Sheets['Base limpeza setembro 2022'];

// REF VAREJISTA
function refVarejista(ref) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'MES' })
    .map((row) => row['MES']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['REF VAREJISTA']].concat(column.map((value) => [ref])),
    {
      origin: 'A1',
    }
  );
}
// REF COD_varejista
function refCodVarejista() {
  let i = 4;
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'MES' })
    .map((row) => row['MES']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['REF COD_VAREJISTA']].concat(
      column.map((value) => {
        let result = '=PROCV($A' + i + ";'DE PARA GERAL'!$L:$M;2;0)";
        i++;
        return [result];
      })
    ),
    { origin: 'B1' }
  );
}
// REF ID_loja
function refIdLoja(headerC) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerC })
    .map((row) => row[headerC]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['REF ID_LOJA']].concat(column.map((value) => [value])),
    { origin: 'C1' }
  );
}
// CHECK EAN
function checkEAN() {
  let i = 4;
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'MES' })
    .map((row) => row['MES']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['CHECK EAN']].concat(
      column.map((value) => {
        let result = '=PROCV($I' + i + ";'DE PARA GERAL'!$AF:$AG;2;0)";
        i++;
        return [result];
      })
    ),
    { origin: 'D1' }
  );
}
// CHECK LOJA
function checkLoja() {
  let i = 4;
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'MES' })
    .map((row) => row['MES']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['CHECK LOJA']].concat(
      column.map((value) => {
        let result = '=PROCV($H' + i + ";'DE PARA GERAL'!$AP:$AR;3;0)";
        i++;
        return [result];
      })
    ),
    { origin: 'E1' }
  );
}
// COD_LOJA
function codLoja(headerH, ref) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerH })
    .map((row) => row[headerH]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['cod_loja']].concat(
      column.map((value) => {
        let result = ref + '10' + value;
        return [result * 1];
      })
    ),
    { origin: 'H1' }
  );
}
// COD_PRODUTO
function codProduto(headerI) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerI })
    .map((row) => row[headerI]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['cod_produto']].concat(
      column.map((value) => {
        let result = [value].toString().slice(0);
        return [result * 1];
      })
    ),
    { origin: 'I1' }
  );
}
// REF_VAREJISTA
function refVarejista2(ref) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: 'MES' })
    .map((row) => row['MES']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['ref_varejista']].concat(column.map((value) => [ref])),
    { origin: 'J1' }
  );
}
// DIA
function getDia(headerK, headerK2) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerK })
    .map((row) => row[headerK]);
  const column2 = XLSX.utils
    .sheet_to_json(ws1, { header: headerK2 })
    .map((row) => row[headerK2]);

  const concatenated = column.map((value, index) => {
    '01/' + value + column2[index];
    if (value === 'jan') return '01/01/' + column2[index];
    else if (value === 'fev') return '01/02/' + column2[index];
    else if (value === 'mar') return '01/03/' + column2[index];
    else if (value === 'abr') return '01/04/' + column2[index];
    else if (value === 'mai') return '01/05/' + column2[index];
    else if (value === 'jun') return '01/06/' + column2[index];
    else if (value === 'jul') return '01/07/' + column2[index];
    else if (value === 'ago') return '01/08/' + column2[index];
    else if (value === 'set') return '01/09/' + column2[index];
    else if (value === 'out') return '01/10/' + column2[index];
    else if (value === 'nov') return '01/11/' + column2[index];
    else if (value === 'dez') return '01/12/' + column2[index];
  });
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['dia']].concat(concatenated.map((value) => [value])),
    { origin: 'K1' }
  );
}
// MES
function getMes(headerN) {
  console.log(headerN);
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerN })
    .map((row) => row[headerN].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['mes']].concat(column.map((value) => [value])),
    { origin: 'M1' }
  );
}
// DECR_MES
function descrMes(headerO) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerO })
    .map((row) => row['Coluna2'].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['decr_mes']].concat(
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
function getAno(headerQ) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerQ })
    .map((row) => row[headerQ]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['ano']].concat(column.map((value) => [value])),
    { origin: 'P1' }
  );
}
// PERIODO
function getPeriodo(headerR) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerR })
    .map((row) => row[headerR].toUpperCase());
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['periodo']].concat(
      column.map((value) => {
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
// VALOR
function valor(headerT) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerT })
    .map((row) => row[headerT]);

  // Somar os valores
  const soma = column.reduce((acc, curr) => acc + curr, 0);

  XLSX.utils.sheet_add_aoa(
    ws2,
    [['valor']].concat(column.map((value) => [value])),
    { origin: 'S1' }
  );
  console.log('Somatória dos VALORES: ', soma.toFixed(2));
}
// QUANTIDADE
function quantidade(headerU) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerU })
    .map((row) => row[headerU]);

  // Somar as quantidades
  const soma = column.reduce((acc, curr) => acc + curr, 0);

  XLSX.utils.sheet_add_aoa(
    ws2,
    [['quantidade']].concat(column.map((value) => [value])),
    { origin: 'T1' }
  );
  console.log('Somatória das QUANTIDADES: ', soma);
}
// CONCAT
function concat(headerX, headerX2) {
  const column = XLSX.utils
    .sheet_to_json(ws1, { header: headerX })
    .map((row) => row[headerX].toUpperCase());
  const column2 = XLSX.utils
    .sheet_to_json(ws1, { header: headerX2 })
    .map((row) => row[headerX2]);

  const concatenated = column.map((value, index) => value + column2[index]);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [['CONCAT']].concat(concatenated.map((value) => [value])),
    { origin: 'W1' }
  );
}

// Carrega o arquivo de destino
const destino = 'C:/Users/leona/Desktop/TESTANTO_MODELO.xlsx';
const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

refVarejista('BISTEK');
refCodVarejista();
refIdLoja('Loja');
checkEAN();
checkLoja();
codLoja('Loja', '7');
codProduto('EAN');
refVarejista2('BISTEK');
getDia('Mês', 'Ano');
getMes('Mês' && 'Mes');
descrMes('Mês');
getAno('Ano');
getPeriodo('Mês');
valor('Venda valor');
quantidade('Venda quantidade');
concat('Mês', 'Ano');
  

XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
XLSX.writeFile(wb2, destino);
