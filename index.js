const { google } = require('googleapis');


// Function to authenticate the Google Sheets API
async function authenticateGoogleSheets() {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: 'credentials.json',
            scopes: ['https://www.googleapis.com/auth/spreadsheets'],
        });
        // We need to authenticate the client before to use the Google Sheets API

        const client = await auth.getClient();

        // Instance of Google Sheets API

        const sheets = google.sheets({
            version: 'v4',
            auth: client
        });

        // Returning the authenticated client and the Google Sheets API instance

        return { auth, sheets };
    } catch (error) {
        console.error('Error authenticating Google Sheets:', error.message);
        throw error;
    }
}

// Function to calculate the situation for each student
async function calculateStudentSituation() {
    try {
        const { auth, sheets } = await authenticateGoogleSheets();

        // ID of the Google Sheets document
        const spreadSheetId = '1v43MQYzn7uJeMZ67emVy5qjf7uzOOCytpLQ_MbI4JqU';
        


        // Adjusted range based on the starting cell
        const range = 'engenharia_de_software!B4:F'; 
        

        
        // Getting the values from the Google Sheets document
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: spreadSheetId,
            range: range,
        });


        // Getting the rows from the response
        const rows = response.data.values;
        

        if (rows && rows.length > 0) {
            rows.forEach(row => {

                // Destructuring the values from the row
                const [name, Faltas, P1, P2, P3, ] = row;
                

                // Converting string values to numbers
                const P1Value = parseInt(P1)/10;
                const P2Value = parseInt(P2)/10;
                const P3Value = parseInt(P3)/10;
                const FaltasValue = parseInt(Faltas);

                // Calculating average
                const average = Math.round((P1Value + P2Value + P3Value) / 3);
                


                // Calculating the situation
                let situation = '';
                let naf = 0;


                if (FaltasValue > 0.25 * 100) {
                    situation = 'Reprovado por Falta';
                } else if (average < 5) {
                    situation = 'Reprovado por Nota';
                } else if (average >= 5 && average < 7) {
                    situation = 'Exame Final';

                    // Calculating naf the grade needed to pass
                    naf = Math.ceil((10 - average));
                } else {
                    situation = 'Aprovado';
                }
                
                // Printing the student situation
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
            });
        } else {
            console.log('No data found');
        }
    } catch (error) {
        console.error('Error calculating student situation:', error.message);
    }
}

// Call the function to calculate the situation for each student
calculateStudentSituation();
