import XLSX from 'xlsx';
import chokidar from 'chokidar';
import fs from 'fs';
import path from 'path';

const folderPath = 'C:/Users/leona/Desktop/Recebe Bases/';

fs.watch(folderPath, (eventType, fileName) => {
  if (eventType === 'rename') {
    const fileExt = path.extname(fileName);
    if (fileExt === '.xlsx' || fileExt === '.xls') {
      const filePath = path.join(folderPath, fileName);
      const wb1 = XLSX.readFile(filePath);
      const sheetNames = wb1.SheetNames;
      sheetNames.forEach((sheetName) => {
        const ws1 = wb1.Sheets[sheetName];
        console.log('Novo arquivo adicionado: ', fileName);
        //HEADERS
        let headerMes = 0;
        let headerAno = 0;
        let headerEAN = 0;
        let headerLoja = 0;
        let headerVenda = 0;
        let headerQtd = 0;

        const range = XLSX.utils.decode_range(ws1['!ref']);

        // BUSCA VELOR DE CÉLUNAS
        for (let c = range.s.c; c <= range.e.c; c++) {
          const cellAddress = XLSX.utils.encode_cell({ r: 0, c });
          const cell = ws1[cellAddress];
          if (cell) {
            const cellValue = cell.v.toString().toLocaleUpperCase();
            const cellPosition = XLSX.utils.decode_cell(cellAddress);
            if (cellValue === 'MÊS' || cellValue === 'MES') {
              headerMes = cellPosition.c;
            } else if (cellValue === 'ANO') {
              headerAno = cellPosition.c;
            } else if (cellValue === 'EAN') {
              headerEAN = cellPosition.c;
            } else if (cellValue === 'LOJA' || cellValue === 'SITE') {
              headerLoja = cellPosition.c;
            } else if (
              cellValue === 'VENDA VALOR' ||
              cellValue === 'VENDA' ||
              cellValue === 'VALOR'
            ) {
              headerVenda = cellPosition.c;
            } else if (
              cellValue === 'VENDA QUANTIDADE' ||
              cellValue === 'VENDA QTDE' ||
              cellValue === 'QUANTIDADE'
            ) {
              headerQtd = cellPosition.c;
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
            [['REF Varejista']].concat(column.map((value) => [ref])),
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
        // RED ID_LOJA
        function refIdLoja(headerLoja) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerLoja]);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => [value]),
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
        function codLoja(headerLoja, ref) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerLoja]);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => {
              let result = ref + '10' + value;
              return [result * 1];
            }),
            { origin: 'H1' }
          );
        }
        // COD_PRODUTO
        function codProduto(headerEAN) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerEAN]);
          //console.log(column);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => {
              let result = [value].toString().slice(0);
              return [result * 1];
            }),
            { origin: 'I1' }
          );
        }
        // REF VAREJISTA 2
        function refVarejista2(ref) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 'A1' })
            .map((row) => row['A1']);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => [ref]),
            {
              origin: 'J1',
            }
          );
        }
        // DIA
        function getDia(headerMes, headerAno) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerMes].toUpperCase());
          const column2 = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerAno]);

          const concatenated = column.map((value, index) => {
            let result = [value].toString();
            if (result === 'JAN') return '01/01/' + column2[index];
            else if (result === 'FEV') return '01/02/' + column2[index];
            else if (result === 'MAR') return '01/03/' + column2[index];
            else if (result === 'ABR') return '01/04/' + column2[index];
            else if (result === 'MAI') return '01/05/' + column2[index];
            else if (result === 'JUN') return '01/06/' + column2[index];
            else if (result === 'JUL') return '01/07/' + column2[index];
            else if (result === 'AGO') return '01/08/' + column2[index];
            else if (result === 'SET') return '01/09/' + column2[index];
            else if (result === 'OUT') return '01/10/' + column2[index];
            else if (result === 'NOV') return '01/11/' + column2[index];
            else if (result === 'DEZ') return '01/12/' + column2[index];
          });
          XLSX.utils.sheet_add_aoa(
            ws2,
            concatenated.map((value) => [value]),
            { origin: 'K1' }
          );
        }
        // MES
        function getMes(headerMes) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
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
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerMes].toUpperCase());
          XLSX.utils.sheet_add_aoa(
            ws2,
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
            }),
            { origin: 'N1' }
          );
        }
        // ANO
        function getAno(headerAno) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerAno]);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => [value]),
            { origin: 'P1' }
          );
        }
        // PERIODO
        function getPeriodo(headerMes) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerMes].toUpperCase());
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => {
              let result = [value].toString();
              if (result === 'JAN' || result === 'FEV') return ['1 BIM'];
              else if (result === 'MAR' || result === 'ABR') return ['2 BIM'];
              else if (result === 'MAI' || result === 'JUN') return ['3 BIM'];
              else if (result === 'JUL' || result === 'AGO') return ['4 BIM'];
              else if (result === 'SET' || result === 'OUT') return ['5 BIM'];
              else if (result === 'NOV' || result === 'DEZ') return ['6 BIM'];
            }),
            { origin: 'Q1' }
          );
        }
        // VALOR
        function valor(headerVenda) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerVenda]);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => [value]),
            { origin: 'S1' }
          );
        }
        // QUANTIDADE
        function quantidade(headerQtd) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerQtd]);
          XLSX.utils.sheet_add_aoa(
            ws2,
            column.map((value) => [value]),
            { origin: 'T1' }
          );
        }
        // CONCAT
        function concat(headerMes, headerAno) {
          const column = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerMes].toUpperCase());
          const column2 = XLSX.utils
            .sheet_to_json(ws1, { header: 1 })
            .map((row) => row[headerAno]);

          const concatenated = column.map(
            (value, index) => value + column2[index]
          );
          XLSX.utils.sheet_add_aoa(
            ws2,
            concatenated.map((value) => [value]),
            { origin: 'W1' }
          );
        }

        // CARREGA ARQUIVO DESTINO
        const destino = 'C:/Users/leona/Desktop/MODELO - ' + fileName;
        const wb2 = XLSX.utils.book_new();
        const ws2 = XLSX.utils.aoa_to_sheet([[]]);

        refCodVarejista();
        refIdLoja(headerLoja);
        checkLoja();
        checkEAN();
        codProduto(headerEAN);
        getDia(headerMes, headerAno);
        getMes(headerMes);
        descrMes(headerMes);
        getAno(headerAno);
        getPeriodo(headerMes);
        valor(headerVenda);
        quantidade(headerQtd);
        concat(headerMes, headerAno);

        // FUNÇÔES QUE PRECISAM SER COMPLETAS DE ACORDO COM O CLIENTE:
        refVarejista('BISTEK');
        refVarejista2('BISTEK');
        codLoja(headerLoja, '7');

        // SALVA ARQUIVO DESTINO
        XLSX.utils.book_append_sheet(wb2, ws2, 'MODELO');
        XLSX.writeFile(wb2, destino);
      });
    }
  }
});
