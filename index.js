const xlsx = require('xlsx');
const fs = require('fs');

// Load the Excel file
const workbook = xlsx.readFile('test-book.xlsx');

// Prepare an object to hold all the data
const jsonData = {};

// Iterate over each sheet in the workbook
workbook.SheetNames.forEach((sheetName) => {
  // Get the data from the current sheet
  const worksheet = workbook.Sheets[sheetName];
  
  // Convert the sheet data to JSON
  const sheetData = xlsx.utils.sheet_to_json(worksheet, { header: 1 }); // Header: 1 for raw arrays
  
  // Add it to the main object with the sheet name as the key
  jsonData[sheetName] = sheetData;
});

// Save the JSON to a file
fs.writeFileSync('output.json', JSON.stringify(jsonData, null, 2));

console.log('Excel data converted to JSON successfully!');
