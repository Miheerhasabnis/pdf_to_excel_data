const XLSX = require('xlsx');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const axios = require('axios');


async function processExcel() {
  // Load the Excel file
  const workbook = XLSX.readFile('./ass.xlsx');
  let sheetName = workbook.SheetNames[0]; // Assuming the data is in the first sheet
  const worksheet = workbook.Sheets[sheetName];

  // Convert worksheet to JSON
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  console.log("before json data ");
  console.log(jsonData);

  // Define column names that will be appended
  const orderNumberColumn = 'Order Number';
  const invoiceNumberColumn = 'Invoice Number';
  const buyername = 'Buyer name';
  const buyeraddress='Buyer address'
  const taxablevalue='Taxable value';
  let orderNumber;
  let invoiceNumber;
  let billTo;
  let buyeraddrs;
  let taxableval;
  // Add the column names to the first row of the JSON data ie headers
  jsonData[0].push(orderNumberColumn, invoiceNumberColumn, buyername,buyeraddress,taxablevalue);

  // Process each row except headers row that is 1st row
  for (let rowIndex = 1; rowIndex < jsonData.length; rowIndex++) {
    const row = jsonData[rowIndex];

    // Get the PDF link from the second column
    const pdfLink = row[1];
    

    // Extract the file name from the PDF link
    const fileName = pdfLink.split('/').pop();

    // Download the PDF file
    const response = await axios({
      url: pdfLink,
      method: 'GET',
      responseType: 'stream',
    });

    // Create a writable stream and pipe the response data into it
    const fileStream = fs.createWriteStream(fileName);
    response.data.pipe(fileStream);

    // Wait for the file to be downloaded
    await new Promise((resolve) => {
      fileStream.on('finish', resolve);
    });

    // Read the downloaded PDF file
    const pdfData = fs.readFileSync(fileName);
    let pdf;

    // Extract the required information from the PDF
    try {
      pdf = await pdfParse(pdfData);
      console.log(pdf.text)
      orderNumber = extractOrderNumber(pdf.text);
      invoiceNumber = extractInvoiceNumber(pdf.text);
      billTo = extractBillTo(pdf.text);
      buyeraddrs=extractaddress(pdf.text);
      taxableval=extracttaxval(pdf.text);
      // Update the JSON data with the extracted order number, invoice number, and bill to data
      row.push(orderNumber, invoiceNumber, billTo,buyeraddrs,taxableval);
    } catch (err) {
      console.log("error occurred bro");
      console.log(err);
    }
  }

  console.log("after json data ");
  console.log(jsonData);

  // Convert the updated jsonData back to worksheet format
  const updatedWorksheet = XLSX.utils.json_to_sheet(jsonData);

  // Assign the updated worksheet to the workbook
  workbook.Sheets[sheetName] = updatedWorksheet;

  // Save the updated workbook
  XLSX.writeFile(workbook, './updated_data.xlsx');
}

// Function to extract the order number from the PDF text
function extractOrderNumber(pdfText) {
  const regex = /Purchase Order Number\s+(\d+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return '';
}

// Function to extract the invoice number from the PDF text
function extractInvoiceNumber(pdfText) {
  const regex = /Invoice Number\s+(\w+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return '';
}

// Function to extract the data after "BILL TO:"
function extractBillTo(pdfText) {
  const regex = /BILL TO:\s+(\w+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1].trim();
  }
  return '';
}

function extractaddress(pdfText) {
    const regex = /SHIP TO:\n([\s\S]*?)\nInvoice Date/;
    const match = pdfText.match(regex);
    if (match && match[1]) {
      return match[1].trim();
    }
    return '';
  }
  function extracttaxval(pdfText) {
    const regex = /TotalRs\.([\d.]+)/;
    const match = pdfText.match(regex);
    if (match && match[1]) {
      return match[1].trim();
    }
    return '';
  }

// Call the async function to start the process
processExcel().catch((error) => {
  console.error(error);
});

async function processExcel() {
  // Load the Excel file
  const workbook = XLSX.readFile('./ass.xlsx');
  let sheetName = workbook.SheetNames[0]; // Assuming the data is in the first sheet
  const worksheet = workbook.Sheets[sheetName];

  // Convert worksheet to JSON
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
  console.log("before json data ");
  console.log(jsonData);

  // Define column names that will be appended
  const orderNumberColumn = 'Order Number';
  const invoiceNumberColumn = 'Invoice Number';
  const buyername = 'Buyer name';
  const buyeraddress='Buyer address'
  let orderNumber;
  let invoiceNumber;
  let billTo;
  let buyeraddrs;
  // Add the column names to the first row of the JSON data ie headers
  jsonData[0].push(orderNumberColumn, invoiceNumberColumn, buyername,buyeraddress);

  // Process each row except headers row that is 1st row
  for (let rowIndex = 1; rowIndex < jsonData.length; rowIndex++) {
    const row = jsonData[rowIndex];

    // Get the PDF link from the second column
    const pdfLink = row[1];
    

    // Extract the file name from the PDF link
    const fileName = pdfLink.split('/').pop();

    // Download the PDF file
    const response = await axios({
      url: pdfLink,
      method: 'GET',
      responseType: 'stream',
    });

    // Create a writable stream and pipe the response data into it
    const fileStream = fs.createWriteStream(fileName);
    response.data.pipe(fileStream);

    // Wait for the file to be downloaded
    await new Promise((resolve) => {
      fileStream.on('finish', resolve);
    });

    // Read the downloaded PDF file
    const pdfData = fs.readFileSync(fileName);
    let pdf;

    // Extract the required information from the PDF
    try {
      pdf = await pdfParse(pdfData);
      console.log(pdf.text)
      orderNumber = extractOrderNumber(pdf.text);
      invoiceNumber = extractInvoiceNumber(pdf.text);
      billTo = extractBillTo(pdf.text);
      buyeraddrs=extractaddress(pdf.text);
      // Update the JSON data with the extracted order number, invoice number, and bill to data
      row.push(orderNumber, invoiceNumber, billTo,buyeraddrs);
    } catch (err) {
      console.log("error occurred bro");
      console.log(err);
    }
  }

  console.log("after json data ");
  console.log(jsonData);

  // Convert the updated jsonData back to worksheet format
  const updatedWorksheet = XLSX.utils.json_to_sheet(jsonData);

  // Assign the updated worksheet to the workbook
  workbook.Sheets[sheetName] = updatedWorksheet;

  // Save the updated workbook
  XLSX.writeFile(workbook, './updated_data.xlsx');
}

// Function to extract the order number from the PDF text
function extractOrderNumber(pdfText) {
  const regex = /Purchase Order Number\s+(\d+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return '';
}

// Function to extract the invoice number from the PDF text
function extractInvoiceNumber(pdfText) {
  const regex = /Invoice Number\s+(\w+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1];
  }
  return '';
}

// Function to extract the data after "BILL TO:"
function extractBillTo(pdfText) {
  const regex = /BILL TO:\s+(\w+)/;
  const match = pdfText.match(regex);
  if (match && match[1]) {
    return match[1].trim();
  }
  return '';
}

function extractaddress(pdfText) {
    const regex = /SHIP TO:\n([\s\S]*?)\nInvoice Date/;
    const match = pdfText.match(regex);
    if (match && match[1]) {
      return match[1].trim();
    }
    return '';
  }

// Call the async function to start the process
processExcel().catch((error) => {
  console.error(error);
});
