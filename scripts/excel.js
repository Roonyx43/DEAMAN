document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('form-excel').addEventListener('submit', function(event) {
        event.preventDefault();

        const getValueOrEmpty = (id) => {
            const element = document.getElementById(id);
            return element ? element.value || "" : ""; // Verifica se o elemento existe e retorna o valor ou vazio
        };

        const data = getValueOrEmpty('data');
        const eqp1 = getValueOrEmpty('eqp1');
        const eqp2 = getValueOrEmpty('eqp2');
        const eqp3 = getValueOrEmpty('eqp3');
        const eqp4 = getValueOrEmpty('eqp4');
        const eqp5 = getValueOrEmpty('eqp5');
        const eqp6 = getValueOrEmpty('eqp6');
        const eqp7 = getValueOrEmpty('eqp7');
        const eqp8 = getValueOrEmpty('eqp8');
        const eqp9 = getValueOrEmpty('eqp9');
        const eqp10 = getValueOrEmpty('eqp10');
        const novo1 = getValueOrEmpty('novo1');
        const novo2 = getValueOrEmpty('novo2');
        const novo3 = getValueOrEmpty('novo3');
        const novo4 = getValueOrEmpty('novo4');
        const novo5 = getValueOrEmpty('novo5');
        const novo6 = getValueOrEmpty('novo6');
        const novo7 = getValueOrEmpty('novo7');
        const novo8 = getValueOrEmpty('novo8');
        const novo9 = getValueOrEmpty('novo9');
        const novo10 = getValueOrEmpty('novo10');
        const revi1 = getValueOrEmpty('revi1');
        const revi2 = getValueOrEmpty('revi2');
        const revi3 = getValueOrEmpty('revi3');
        const revi4 = getValueOrEmpty('revi4');
        const revi5 = getValueOrEmpty('revi5');
        const revi6 = getValueOrEmpty('revi6');
        const revi7 = getValueOrEmpty('revi7');
        const revi8 = getValueOrEmpty('revi8');
        const revi9 = getValueOrEmpty('revi9');
        const revi10 = getValueOrEmpty('revi10');
        const desc1 = getValueOrEmpty('desc1');
        const desc2 = getValueOrEmpty('desc2');
        const desc3 = getValueOrEmpty('desc3');
        const desc4 = getValueOrEmpty('desc4');
        const desc5 = getValueOrEmpty('desc5');
        const desc6 = getValueOrEmpty('desc6');
        const desc7 = getValueOrEmpty('desc7');
        const desc8 = getValueOrEmpty('desc8');
        const desc9 = getValueOrEmpty('desc9');
        const desc10 = getValueOrEmpty('desc10');

        fetch('../assets/modelo.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = new ExcelJS.Workbook();
            return workbook.xlsx.load(data);
        })
        .then(workbook => {
            const worksheet = workbook.getWorksheet('Setembro - 09 & 10'); // Obtém a primeira planilha
            if (!worksheet) {
                throw new Error('A planilha não foi encontrada no arquivo Excel.');
            }

            worksheet.getCell('E1').value = data;
            worksheet.getCell('B3').value = eqp1;
            worksheet.getCell('B4').value = eqp2;
            worksheet.getCell('B5').value = eqp3;
            worksheet.getCell('B6').value = eqp4;
            worksheet.getCell('B7').value = eqp5;
            worksheet.getCell('B8').value = eqp6;
            worksheet.getCell('B9').value = eqp7;
            worksheet.getCell('B10').value = eqp8;
            worksheet.getCell('B11').value = eqp9;
            worksheet.getCell('B12').value = eqp10;

            worksheet.getCell('C3').value = novo1;
            worksheet.getCell('C4').value = novo2;
            worksheet.getCell('C5').value = novo3;
            worksheet.getCell('C6').value = novo4;
            worksheet.getCell('C7').value = novo5;
            worksheet.getCell('C8').value = novo6;
            worksheet.getCell('C9').value = novo7;
            worksheet.getCell('C10').value = novo8;
            worksheet.getCell('C11').value = novo9;
            worksheet.getCell('C12').value = novo10;

            worksheet.getCell('D3').value = revi1;
            worksheet.getCell('D4').value = revi2;
            worksheet.getCell('D5').value = revi3;
            worksheet.getCell('D6').value = revi4;
            worksheet.getCell('D7').value = revi5;
            worksheet.getCell('D8').value = revi6;
            worksheet.getCell('D9').value = revi7;
            worksheet.getCell('D10').value = revi8;
            worksheet.getCell('D11').value = revi9;
            worksheet.getCell('D12').value = revi10;

            worksheet.getCell('E3').value = desc1;
            worksheet.getCell('E4').value = desc2;
            worksheet.getCell('E5').value = desc3;
            worksheet.getCell('E6').value = desc4;
            worksheet.getCell('E7').value = desc5;
            worksheet.getCell('E8').value = desc6;
            worksheet.getCell('E9').value = desc7;
            worksheet.getCell('E10').value = desc8;
            worksheet.getCell('E11').value = desc9;
            worksheet.getCell('E12').value = desc10;

            // Gera o arquivo e faz o download
            return workbook.xlsx.writeBuffer();
        })
        .then(buffer => {
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'arquivo_editado.xlsx'; // Nome do arquivo baixado
            link.click();
        })
        .catch(err => console.error("Erro ao gerar o arquivo Excel: ", err));
    });
});
