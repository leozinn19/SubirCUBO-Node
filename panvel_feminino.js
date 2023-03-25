import XLSX from 'xlsx';

const wb1 = XLSX.readFile(
  'C:/Users/leona/Desktop/Clientes - Desktop/KC/Bases/Panvel - Cuidado Feminino/Sell Out H.I - 2021 (2).xlsx'
);
const ws1 = wb1.Sheets['Sheet1'];

function refVarejista(ref, ref2) {
  const headerA = ['REF VAREJISTA'];
  const columnA = XLSX.utils
    .sheet_to_json(ws1, { header: 'Filial' })
    .map((row) => row['Filial']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerA].concat(columnA.map((value) => [ref])),
    { origin: ref2 }
  );
}

function refCodVarejista(headerB, ref) {
  let i = 4;
  const columnB = XLSX.utils
    .sheet_to_json(ws1, { header: 'Filial' })
    .map((row) => row['Filial']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerB].concat(
      columnB.map((value) => {
        let result = '=PROCV($A' + i + ";'DE PARA GERAL'!$K:$L;2;0)";
        i++;
        return [result];
      })
    ),
    { origin: 'B1' }
  );
}

function refIdLoja(headerC) {
  const columnB = XLSX.utils
    .sheet_to_json(ws1, { header: 'Filial' })
    .map((row) => row['Filial']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerC].concat(columnB.map((value) => [value])),
    { origin: 'C1' }
  );
}

function codLoja(headerH) {
  const columnC = XLSX.utils
    .sheet_to_json(ws1, { header: 'Filial' })
    .map((row) => row['Filial']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerH].concat(
      columnC.map((value) => {
        let result = '410' + value;
        return [result];
      })
    ),
    { origin: 'H1' }
  );
}

function codProduto(headerI) {
  const columnC = XLSX.utils
    .sheet_to_json(ws1, { header: 'Item - EAN' })
    .map((row) => row['item - EAN']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerI].concat(columnC.map((value) => [value])),
    { origin: 'I1' }
  );
}

function getDia(headerDia) {
  const columnData = XLSX.utils
    .sheet_to_json(ws1, { header: 'Tempo - Ano Mes' })
    .map((row) => row['Tempo - Ano Mes']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerDia].concat(
      columnData.map((value) => {
        let result =
          '01/' +
          value.toString().slice(4, 6) +
          '/' +
          value.toString().slice(0, 4);
        return [result];
      })
    ),
    { origin: 'K1' }
  );
}

function valor(headerT) {
  const columnD = XLSX.utils
    .sheet_to_json(ws1, { header: 'V.Efet.HB+OTC' })
    .map((row) => row['V.Efet.HB+OTC']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerT].concat(columnD.map((value) => [value])),
    { origin: 'T1' }
  );
}

function quantidade(headerU) {
  const columnE = XLSX.utils
    .sheet_to_json(ws1, { header: 'Qtd Vda HB+OTC' })
    .map((row) => row['Qtd Vda HB+OTC']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerU].concat(columnE.map((value) => [value])),
    { origin: 'U1' }
  );
}

function getAno(headerAno) {
  const columnData = XLSX.utils
    .sheet_to_json(ws1, { header: 'Tempo - Ano Mes' })
    .map((row) => row['Tempo - Ano Mes']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerAno].concat(
      columnData.map((value) => {
        return [[value].toString().slice(0, 4)];
      })
    ),
    { origin: 'Q1' }
  );
}

function getPeriodo(headerBim) {
  const columnMes = XLSX.utils
    .sheet_to_json(ws1, { header: 'Tempo - Ano Mes' })
    .map((row) => row['Tempo Ano Mes']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerBim].concat(
      columnMes.map((value) => {
        let result = [value].toString().slice(4, 6);
        if (result === '01' || result === '02') return ['1 BIM'];
        else if (result === '03' || result === '04') return ['2 BIM'];
        else if (result === '05' || result === '06') return ['3 BIM'];
        else if (result === '07' || result === '08') return ['4 BIM'];
        else if (result === '09' || result === '10') return ['5 BIM'];
        else if (result === '11' || result === '12') return ['6 BIM'];
      })
    ),
    { origin: 'R1' }
  );
}

function getMes(headerMes) {
  const columnMes = XLSX.utils
    .sheet_to_json(ws1, { header: 'Tempo - Ano Mes' })
    .map((row) => row['Tempo - Ano Mes']);
  XLSX.utils.sheet_add_aoa(
    ws2,
    [headerMes].concat(
      columnMes.map((value) => {
        let result = [value].toString().slice(4, 6);
        if (result === '01') return ['JAN'];
        else if (result === '02') return ['FEV'];
        else if (result === '03') return ['MAR'];
        else if (result === '04') return ['ABR'];
        else if (result === '05') return ['MAI'];
        else if (result === '06') return ['JUN'];
        else if (result === '07') return ['JUL'];
        else if (result === '08') return ['AGO'];
        else if (result === '09') return ['SET'];
        else if (result === '10') return ['OUT'];
        else if (result === '11') return ['NOV'];
        else if (result === '12') return ['DEZ'];
      })
    ),
    { origin: 'N1' }
  );
}

const wb2 = XLSX.utils.book_new();
const ws2 = XLSX.utils.aoa_to_sheet([[]]);

refVarejista('PANVEL LOJA', 'A1');
refCodVarejista(['REF COD_VAREJISTA'], ['2']);
refIdLoja(['RED ID_LOJA']);
codLoja(['cod_loja']);
codProduto(['cod_pruduto']);
valor(['valor']);
quantidade(['quantidade']);
getPeriodo(['periodo']);
getAno(['ano']);
getMes(['mes']);
getDia(['dia']);
refVarejista('PANVEL LOJA', 'J1');

XLSX.utils.book_append_sheet(wb2, ws2, 'Planilha1');
XLSX.writeFile(wb2, 'C:/Users/leona/Desktop/PANVEL-LOJA_BASE_FORMAT.xlsx');
