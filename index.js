const { google } = require('googleapis');

// Função para autenticar a API do Google Sheets
async function authenticateGoogleSheets() {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: 'api/credentials.json',
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });

        const client = await auth.getClient();

        const sheets = google.sheets({
            version: 'v4',
            auth: client
        });

        return { auth, sheets };
    } catch (error) {
        console.error('Erro na autenticação do Google Sheets:', error.message);
        throw error;
    }
}


// Função para escrever no Google Sheets
async function writeToGoogleSheets(data) {
    try {
        const { auth, sheets } = await authenticateGoogleSheets();

        // ID do documento do Google Sheets
        const spreadSheetId = '1v43MQYzn7uJeMZ67emVy5qjf7uzOOCytpLQ_MbI4JqU';

        // Intervalo ajustado para atualização de dados
        const updateRange = 'engenharia_de_software!G4:H';

        // Atualizando o Google Sheets com os novos dados
        await sheets.spreadsheets.values.update({
            spreadsheetId: spreadSheetId,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: data,
            },
        });

        console.log('Dados dos alunos atualizados com sucesso.');
    } catch (error) {
        console.error('Erro ao atualizar o Google Sheets:', error.message);
    }
}

// Função para calcular a situação de cada aluno e atualizar o Google Sheets
async function calculateStudentSituation() {
    try {
        const { sheets } = await authenticateGoogleSheets();

        // ID do documento do Google Sheets
        const spreadSheetId = '1v43MQYzn7uJeMZ67emVy5qjf7uzOOCytpLQ_MbI4JqU';

        // Intervalo ajustado com base na célula inicial
        const range = 'engenharia_de_software!B4:F';

        // Obtendo os valores do documento do Google Sheets
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadSheetId,
            range: range,
        });

        // Obtendo as linhas da resposta
        const rows = response.data.values;

        if (rows && rows.length > 0) {
            const updatedData = rows.map(row => {
                const [name, Faltas, P1, P2, P3] = row;

                // Convertendo valores de string para números
                const P1Value = parseInt(P1) / 10;
                const P2Value = parseInt(P2) / 10;
                const P3Value = parseInt(P3) / 10;
                const FaltasValue = parseInt(Faltas);

                // Calculando a média
                const average = Math.round((P1Value + P2Value + P3Value) / 3);

                // Calculando a situação
                let situation = '';
                let naf = 0;

                if (FaltasValue > 0.25 * 60) {
                    situation = 'Reprovado por Falta';
                } else if (average < 5) {
                    situation = 'Reprovado por Nota';
                } else if (average >= 5 && average < 7) {
                    situation = 'Exame Final';

                    // Calculando o NAF (Nota para Aprovação Final)
                    naf = Math.ceil((10 - average));
                } else {
                    situation = 'Aprovado';
                }

                console.log(`
                    Nome: ${name}
                    P1: ${P1Value}
                    P2: ${P2Value}
                    P3: ${P3Value}
                    Faltas: ${FaltasValue}
                    Média: ${average}
                    Situação: ${situation}
                    NAF: ${situation === 'Exame Final' ? naf : 0}
                    ------------------------------------
                `);

                return [
                    situation,
                    situation === 'Exame Final' ? naf : 0,
                ];
            });

            // Chama a função para escrever na tabela
            await writeToGoogleSheets(updatedData);
        } else {
            console.log('Nenhum dado encontrado');
        }
    } catch (error) {
        console.error('Erro ao calcular a situação dos alunos:', error.message);
    }
}

// Chama a função para calcular a situação de cada aluno e atualizar o Google Sheets
calculateStudentSituation();
