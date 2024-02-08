const { google } = require('googleapis');

// Function to authenticate the Google Sheets API
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
        console.error('Error authenticating Google Sheets:', error.message);
        throw error;
    }
}

// Function to write to Google Sheets
async function writeToGoogleSheets(data) {
    try {
        const { auth, sheets } = await authenticateGoogleSheets();

        // ID of the Google Sheets document
        const spreadSheetId = '1v43MQYzn7uJeMZ67emVy5qjf7uzOOCytpLQ_MbI4JqU';

        // Adjusted range for data update
        const updateRange = 'engenharia_de_software!G4:H';

        // Updating Google Sheets with new data
        await sheets.spreadsheets.values.update({
            spreadsheetId: spreadSheetId,
            range: updateRange,
            valueInputOption: 'USER_ENTERED',
            resource: {
                values: data,
            },
        });

        console.log('Student data updated successfully.');
    } catch (error) {
        console.error('Error updating Google Sheets:', error.message);
    }
}

// Function to calculate the situation for each student and update Google Sheets
async function calculateStudentSituation() {
    try {
        const { sheets } = await authenticateGoogleSheets();

        // ID of the Google Sheets document
        const spreadSheetId = '1v43MQYzn7uJeMZ67emVy5qjf7uzOOCytpLQ_MbI4JqU';

        // Adjusted range based on the starting cell
        const range = 'engenharia_de_software!B4:F';

        // Getting values from the Google Sheets document
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadSheetId,
            range: range,
        });

        // Getting rows from the response
        const rows = response.data.values;

        if (rows && rows.length > 0) {
            const updatedData = rows.map(row => {
                const [name, Faltas, P1, P2, P3] = row;

                // Converting string values to numbers
                const P1Value = parseInt(P1) / 10;
                const P2Value = parseInt(P2) / 10;
                const P3Value = parseInt(P3) / 10;
                const FaltasValue = parseInt(Faltas);

                // Calculating average
                const average = Math.round((P1Value + P2Value + P3Value) / 3);

                // Calculating the situation
                let situation = '';
                let naf = 0;

                if (FaltasValue > 0.25 * 60) {
                    situation = 'Reprovado por Falta';
                } else if (average < 5) {
                    situation = 'Reprovado por Nota';
                } else if (average >= 5 && average < 7) {
                    situation = 'Exame Final';

                    // Calculating NAF (Grade needed to pass)
                    naf = Math.ceil((10 - average));
                } else {
                    situation = 'Aprovado';
                }

                console.log(`
                    Name: ${name}
                    P1: ${P1Value}
                    P2: ${P2Value}
                    P3: ${P3Value}
                    Faltas: ${FaltasValue}
                    Average: ${average}
                    Situation: ${situation}
                    NAF: ${situation === 'Exame Final' ? naf : 0}
                    ------------------------------------
                `);

                return [
                    situation,
                    situation === 'Exame Final' ? naf : 0,
                ];
            });

            // Call the function to write to the table
            await writeToGoogleSheets(updatedData);
        } else {
            console.log('No data found');
        }
    } catch (error) {
        console.error('Error calculating student situation:', error.message);
    }
}

// Call the function to calculate the situation for each student and update Google Sheets
calculateStudentSituation();
