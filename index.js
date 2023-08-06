// right-1
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage }).single('excelFile');

const convertExcelToJson = async (excelBuffer, sheetIdentifier) => {
    const workbook = new ExcelJS.Workbook();

    try {
        await workbook.xlsx.load(excelBuffer);

        let worksheet;

        if (typeof sheetIdentifier === 'number') {
            worksheet = workbook.worksheets[sheetIdentifier];
        } else {
            worksheet = workbook.getWorksheet(sheetIdentifier);
        }

        if (!worksheet) {
            console.log(`Worksheet '${sheetIdentifier}' not found.`);
            return []; // Return an empty array or handle the error as needed
        }

        const json = [];

        worksheet.eachRow(row => {
            const rowObject = {};

            row.eachCell((cell, colNumber) => {
                rowObject[`col${colNumber}`] = cell.value;
            });

            json.push(rowObject);
        });

        return json;
    } catch (error) {
        console.error('Error:', error);
        throw error;
    }
};

app.post('/convert', upload, async (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).send('No file uploaded.');
        }

        const excelBuffer = req.file.buffer;
        const sheetIdentifier = req.body.sheetIdentifier; // Access the sheet identifier from req.body
        console.log(sheetIdentifier)
        if (!sheetIdentifier) {
            return res.status(400).send('Sheet identifier is missing.');
        }

        const jsonData = await convertExcelToJson(excelBuffer, sheetIdentifier);

        return res.status(200).json(jsonData);
    } catch (error) {
        console.error(error);
        return res.status(500).send('An error occurred during conversion.');
    }
});





app.get('/', (req, res) => {
    res.send('Hello World');
});

const port = process.env.PORT || 5000;
app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});
