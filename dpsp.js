import readline from 'readline';
import XLSX from 'xlsx';

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout,
});

// Carrega arquivo de origem
rl.question('Informe a BASE a ser lida: ', (filePath) => {
  const wb1 = XLSX.readFile(filePath);
  rl.question('Informe o nome da planilha: ', (sheetPath) => {
    const ws1 = wb1.Sheets[sheetPath];

    // Obter o intervalo de células da planilha de origem que deseja copiar
    const range = XLSX.utils.decode_range(ws1['!ref']);

    function refVarejista(ref) {
      const columnA = XLSX.utils
        .sheet_to_json(ws1, { header: 'ANOMES' })
        .map((row) => row['ANOMES']);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['REF VAREJISTA']].concat(columnA.map((value) => [ref])),
        {
          origin: 'A1',
        }
      );
    }

    function refCodVarejista() {
      let i = 4;
      const columnB = XLSX.utils
        .sheet_to_json(ws1, { header: 'ANOMES' })
        .map((row) => row['ANOMES']);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['REF COD_VAREJISTA']].concat(
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
        .sheet_to_json(ws1, { header: headerC })
        .map((row) => row[headerC]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['REF ID_LOJA']].concat(columnB.map((value) => [value])),
        { origin: 'C1' }
      );
    }

    function checkEAN() {
      let i = 4;
      const columnB = XLSX.utils
        .sheet_to_json(ws1, { header: 'ANOMES' })
        .map((row) => row['ANOMES']);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['CHECK EAN']].concat(
          columnB.map((value) => {
            let result = '=PROCV($I' + i + ";'DE PARA GERAL'!$AD:$AF;3;0)";
            i++;
            return [result];
          })
        ),
        { origin: 'D1' }
      );
    }

    function checkLoja() {
      let i = 4;
      const columnB = XLSX.utils
        .sheet_to_json(ws1, { header: 'ANOMES' })
        .map((row) => row['ANOMES']);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['CHECK LOJA']].concat(
          columnB.map((value) => {
            let result = '=PROCV($H' + i + ";'DE PARA GERAL'!$AP:$AR;3;0)";
            i++;
            return [result];
          })
        ),
        { origin: 'E1' }
      );
    }

    function codLoja(headerH, ref) {
      const columnC = XLSX.utils
        .sheet_to_json(ws1, { header: headerH })
        .map((row) => row[headerH]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['cod_loja']].concat(
          columnC.map((value) => {
            let result = ref + '10' + value;
            return [result * 1];
          })
        ),
        { origin: 'H1' }
      );
    }

    function codProduto(headerI) {
      const columnC = XLSX.utils
        .sheet_to_json(ws1, { header: headerI })
        .map((row) => row[headerI]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['cod_produto']].concat(columnC.map((value) => [value * 1])),
        { origin: 'I1' }
      );
    }

    function refVarejista2(ref) {
      const columnA = XLSX.utils
        .sheet_to_json(ws1, { header: 'ANOMES' })
        .map((row) => row['ANOMES']);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['ref_varejista']].concat(columnA.map((value) => [ref])),
        { origin: 'J1' }
      );
    }

    function getDia(headerDia) {
      const columnData = XLSX.utils
        .sheet_to_json(ws1, { header: headerDia })
        .map((row) => row[headerDia]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['dia']].concat(
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

    function getMes(headerMes) {
      const columnMes = XLSX.utils
        .sheet_to_json(ws1, { header: headerMes })
        .map((row) => row[headerMes]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['mes']].concat(
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

    function descrMes(headerMes) {
      const columnMes = XLSX.utils
        .sheet_to_json(ws1, { header: headerMes })
        .map((row) => row[headerMes]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['decr_mes']].concat(
          columnMes.map((value) => {
            let result = [value].toString().slice(4, 6);
            if (result === '01') return ['01-JAN'];
            else if (result === '02') return ['02-FEV'];
            else if (result === '03') return ['03-MAR'];
            else if (result === '04') return ['04-ABR'];
            else if (result === '05') return ['05-MAI'];
            else if (result === '06') return ['06-JUN'];
            else if (result === '07') return ['07-JUL'];
            else if (result === '08') return ['08-AGO'];
            else if (result === '09') return ['09-SET'];
            else if (result === '10') return ['10-OUT'];
            else if (result === '11') return ['11-NOV'];
            else if (result === '12') return ['12-DEZ'];
          })
        ),
        { origin: 'O1' }
      );
    }

    function getAno(headerAno) {
      const columnData = XLSX.utils
        .sheet_to_json(ws1, { header: headerAno })
        .map((row) => row[headerAno]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['ano']].concat(
          columnData.map((value) => {
            return [[value].toString().slice(0, 4)];
          })
        ),
        { origin: 'Q1' }
      );
    }

    function getPeriodo(headerBim) {
      const columnMes = XLSX.utils
        .sheet_to_json(ws1, { header: headerBim })
        .map((row) => row[headerBim]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['periodo']].concat(
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

    function valor(headerT) {
      const columnD = XLSX.utils
        .sheet_to_json(ws1, { header: headerT })
        .map((row) => row[headerT]);

      // Somar os valores
      const soma = columnD.reduce((acc, curr) => acc + curr, 0);

      XLSX.utils.sheet_add_aoa(
        ws2,
        [['valor']].concat(columnD.map((value) => [value])),
        { origin: 'T1' }
      );
      console.log('Somatória dos VALORES: ', soma.toFixed(2));
    }

    function quantidade(headerU) {
      const columnE = XLSX.utils
        .sheet_to_json(ws1, { header: headerU })
        .map((row) => row[headerU]);

      // Somar as quantidades
      const soma = columnE.reduce((acc, curr) => acc + curr, 0);

      XLSX.utils.sheet_add_aoa(
        ws2,
        [['quantidade']].concat(columnE.map((value) => [value])),
        { origin: 'U1' }
      );
      console.log('Somatória das QUANTIDADES: ', soma);
    }

    function concat(headerX) {
      const columnMes = XLSX.utils
        .sheet_to_json(ws1, { header: headerX })
        .map((row) => row[headerX]);
      XLSX.utils.sheet_add_aoa(
        ws2,
        [['CONCAT']].concat(
          columnMes.map((value) => {
            let result = [value].toString().slice(4, 6);
            let ano = [value].toString().slice(0, 4);
            if (result === '01') return ['JAN' + ano];
            else if (result === '02') return ['FEV' + ano];
            else if (result === '03') return ['MAR' + ano];
            else if (result === '04') return ['ABR' + ano];
            else if (result === '05') return ['MAI' + ano];
            else if (result === '06') return ['JUN' + ano];
            else if (result === '07') return ['JUL' + ano];
            else if (result === '08') return ['AGO' + ano];
            else if (result === '09') return ['SET' + ano];
            else if (result === '10') return ['OUT' + ano];
            else if (result === '11') return ['NOV' + ano];
            else if (result === '12') return ['DEZ' + ano];
          })
        ),
        { origin: 'X1' }
      );
    }

    // Carrega o arquivo de destino
    const destino = 'C:/Users/leona/Desktop/DPSP_MODELO.xlsx';
    const wb2 = XLSX.utils.book_new();
    const ws2 = XLSX.utils.aoa_to_sheet([[]]);

    refVarejista('DPSP');
    refCodVarejista();
    refIdLoja('COD_LOJA');
    checkEAN();
    checkLoja();
    codLoja('COD_LOJA', '2');
    codProduto('EAN');
    refVarejista2('DPSP');
    getDia('ANOMES');
    getMes('ANOMES');
    descrMes('ANOMES');
    getAno('ANOMES');
    getPeriodo('ANOMES');
    valor('VENDA');
    quantidade('QTD');
    concat('ANOMES');

    XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
    XLSX.writeFile(wb2, destino);

    rl.close();
  });
});
