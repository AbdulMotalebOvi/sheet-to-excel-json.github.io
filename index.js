// const express = require('express');
// const multer = require('multer');
// const ExcelJS = require('exceljs');
// const jsonfile = require('jsonfile');
// const path = require('path');

// const app = express();
// const port = 5000;

// app.use(express.static('public'));
// app.use(express.urlencoded({ extended: true }));

// const storage = multer.memoryStorage();
// const upload = multer({ storage: storage }).single('excelFile');

// const convertExcelToJson = async (excelBuffer, outputFile, sheetName) => {
//     const workbook = new ExcelJS.Workbook();

//     try {
//         await workbook.xlsx.load(excelBuffer);
//         const worksheet = workbook.getWorksheet(sheetName);

//         const json = [];

//         worksheet.eachRow(row => {
//             const rowObject = {};

//             row.eachCell((cell, colNumber) => {
//                 rowObject[`col${colNumber}`] = cell.value;
//             });

//             json.push(rowObject);
//         });

//         await jsonfile.writeFile(outputFile, json, { spaces: 4 });
//         console.log(`JSON file saved to ${outputFile}`);
//     } catch (error) {
//         console.error('Error:', error);
//         throw error;
//     }
// };


// // // Serve static files from the "public" directory
// // app.use(express.static(path.join(__dirname, 'views')));
// // app.get('/', (req, res) => {
// //     res.sendFile(path.join(__dirname, 'views', 'index.html'));
// // });

// app.get('/download', (req, res) => {
//     const jsonFilePath = path.join(__dirname, 'public', 'output.json');
//     res.download(jsonFilePath, 'output.json'); // Serve the JSON file for download
// })


// app.post('/convert', upload, async (req, res) => {
//     try {
//         if (!req.file) {
//             return res.status(400).send('No file uploaded.');
//         }

//         const excelBuffer = req.file.buffer;
//         const outputFile = __dirname + '/public/output.json';

//         await convertExcelToJson(excelBuffer, outputFile, 'Sheet1');

//         return res.status(200).sendFile(outputFile);
//     } catch (error) {
//         console.error(error);
//         return res.status(500).send('An error occurred during conversion.');
//     }
// });

// app.listen(port, () => {
//     console.log(`Server is listening on port ${port}`);
// });



const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const jsonfile = require('jsonfile');
const path = require('path');
const cors = require('cors');

const app = express();
const port = 5000;

app.use(cors());
app.use(express.json());

const storage = multer.memoryStorage();
const upload = multer({ storage: storage }).single('excelFile');

const convertExcelToJson = async (excelBuffer, outputFile, sheetName) => {
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

        await jsonfile.writeFile(outputFile, json, { spaces: 4 });
        console.log(`JSON file saved to ${outputFile}`);
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
        const outputFile = path.join(__dirname, 'public', 'output.json');

        await convertExcelToJson(excelBuffer, outputFile, 'Sheet1');

        const jsonData = await jsonfile.readFile(outputFile);
        return res.status(200).json(jsonData);
    } catch (error) {
        console.error(error);
        return res.status(500).send('An error occurred during conversion.');
    }
});

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});

