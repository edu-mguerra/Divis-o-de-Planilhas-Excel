const XlsxPopulate = require('xlsx-populate');
const path = require('path');

async function dividePlanilhas(filename) {
    try {
        console.log('Lendo arquivo Excel...');
        const workbook = await XlsxPopulate.fromFileAsync(filename); //ler o excell do filename e guarda no variavel

        console.log('Dividindo planilhas...');
        workbook.sheets().forEach(async (sheet) => { // função para obter todas as planilhar do arquivo
            console.log(`Dividindo planilha: ${sheet.name()}`);
            
            // Criando uma nova pasta de trabalho e uma nova planilha
            const newWorkbook = await XlsxPopulate.fromBlankAsync();
            const newSheet = newWorkbook.sheet(0);

            // Acessando o intervalo de células usado na planilha original e copiando seus valores para a nova planilha
            const usedRange = sheet.usedRange();
            newSheet.cell("A1").value(usedRange.value());

            // path to the work area
            const desktopPath = path.join(require('os').homedir(), 'Desktop');
            await newWorkbook.toFileAsync(path.join(desktopPath, `${sheet.name()}.xlsx`));
            console.log(`Planilha ${sheet.name()} dividida com sucesso.`);
        });

        console.log('Todas as planilhas foram divididas com sucesso.');
    } catch (error) {
        console.error('Erro ao dividir planilhas:', error);
    }
}

// Nome do arquivo Excel na mesma pasta que o código
const filename = 'FINAL - RELAÇÃO DE VEÍCULOS - OFICIALL.xlsx';
dividePlanilhas(filename);
