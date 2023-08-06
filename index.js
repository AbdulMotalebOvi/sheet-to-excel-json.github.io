const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();


app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage }).single('excelFile');

const convertExcelToJson = async (excelBuffer, sheetName) => {
    const workbook = new ExcelJS.Workbook();

    try {
        await workbook.xlsx.load(excelBuffer);
        const worksheet = workbook.getWorksheet(sheetName);

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

        const jsonData = await convertExcelToJson(excelBuffer, 'Sheet1');

        return res.status(200).json(jsonData);
    } catch (error) {
        console.error(error);
        return res.status(500).send('An error occurred during conversion.');
    }
});

app.get('/', (req, res) => {
    res.send('Hello World');
});


app.listen(5000, () => console.log('listening on port 5000...'));