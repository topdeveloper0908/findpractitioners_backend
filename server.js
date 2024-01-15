// Import the 'http' module
const fs = require('fs');
const xlsx = require('xlsx');
const express = require('express');
const bodyParser = require('body-parser')

const app = express();

app.use(function(req, res, next) {
    res.header("Access-Control-Allow-Origin", "*"); // Replace with the actual origin of your frontend
    res.header("Access-Control-Allow-Headers", "Origin, X-Requested-With, Content-Type, Accept");
    next();
});

app.use(bodyParser.json())

app.get('/data', (req, res) => {
    // Read the xlsx file
    const workbook = xlsx.readFile('db.xlsx'); // Replace with the actual path to your xlsx file
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);
  
    // Send the data as a JSON response
    res.json(data);
});

app.post('/saveData', (req, res) => {
    const newData = req.body;
  
    // Read existing data from the Excel file
    let existingData = [];
    if (fs.existsSync('db.xlsx')) {
        const workbook = xlsx.readFile('db.xlsx');
        const sheetName = workbook.SheetNames[0];
        existingData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
    }

    // Append the new data to the existing data
    const updatedData = [...existingData, ...newData];

    // Create a new workbook and add the updated data to a new sheet
    const updatedWorkbook = xlsx.utils.book_new();
    const updatedWorksheet = xlsx.utils.json_to_sheet(updatedData);
    xlsx.utils.book_append_sheet(updatedWorkbook, updatedWorksheet, 'Sheet1');

    // Write the updated workbook to the Excel file
    xlsx.writeFile(updatedWorkbook, 'db.xlsx');

    res.json({ message: 'New row added to Excel file' });
});
  

app.listen(3000, () => {
    console.log('Server running at http://localhost:3000/');
});