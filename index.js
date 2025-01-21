const xlsx = require('xlsx');
const fs = require('fs');

// Load the Excel file
const workbook = xlsx.readFile('test-book.xlsx');

// Prepare an object to hold all the data
const jsonData = {};

// Helper function to format headers
const formatHeader = (header) =>
  header.toLowerCase().replace(/\s+/g, '_');

// Iterate over each sheet in the workbook
workbook.SheetNames.forEach((sheetName) => {
  // Get the data from the current sheet
  const worksheet = workbook.Sheets[sheetName];
  
  // Convert the sheet data to JSON, processing headers
  const sheetData = xlsx.utils.sheet_to_json(worksheet, {
    header: 1 // Extract as raw arrays first
  });

  // Ensure sheetData has rows to process
  if (sheetData.length > 0) {
    const headers = sheetData[0].map(formatHeader); // Format the headers
    const dataRows = sheetData.slice(1); // Skip the header row

    // Convert rows into objects using the formatted headers
    const formattedData = dataRows.map((row) => {
      const rowObject = {};
      headers.forEach((header, index) => {
        rowObject[header] = row[index] || null; // Assign data or null if missing
      });
      return rowObject;
    });

    // Add formatted data to the JSON object
    jsonData[sheetName.toLowerCase().replace(/\s+/g, '_')] = formattedData;
  }
});

// Save the JSON to a file
fs.writeFileSync('output.json', JSON.stringify(jsonData, null, 2));

console.log('Excel data converted to JSON with formatted property names!');