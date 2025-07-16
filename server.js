const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(bodyParser.json());

const PORT = 3001;

app.post('/submit-form', async (req, res) => {
    const { name, dob, address, contact } = req.body;
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    try {
        await workbook.xlsx.readFile('data.xlsx');
        worksheet = workbook.getWorksheet('FormData');
    } catch (err) {
        worksheet = workbook.addWorksheet('FormData');
        worksheet.columns = [
            { header: 'Name', key: 'name' },
            { header: 'DOB', key: 'dob' },
            { header: 'Address', key: 'address' },
            { header: 'Contact', key: 'contact' },
            { header: 'Timestamp', key: 'timestamp' }
        ];
    }

    worksheet.addRow({
        name,
        dob,
        address,
        contact,
        timestamp: new Date().toLocaleString()
    });

    await workbook.xlsx.writeFile('data.xlsx');
    res.json({ message: 'Data saved successfully!' });
});

app.listen(PORT, () => {
    console.log(`Server is running on http://localhost:${PORT}`);
});
