const XLSX = require("xlsx");
const fs = require("fs").promises;
const path = require("path");
const crypto = require("crypto");

async function excelToHTML(filePath, sessionTempDir = null) {
  try {
    // Generate a unique ID for this conversion
    const uniqueId = crypto.randomBytes(16).toString('hex');
    const tempDir = sessionTempDir || path.join(process.cwd(), 'temp');
    
    // Ensure temp directory exists
    try {
      await fs.mkdir(tempDir, { recursive: true });
    } catch (error) {
      if (error.code !== 'EEXIST') throw error;
    }

    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    let data = [];
    
    // Read up to 100 rows or until the end of the data
    for (let row = range.s.r; row <= range.s.r + 100; row++) {
      data[row] = [];
      for (let col = range.s.c; col <= range.e.c; col++) {
        let cell = worksheet[XLSX.utils.encode_cell({ r: row, c: col })];
        data[row].push(cell ? cell.v : "");
      }
    }

    // Filter out empty rows
    data = data.filter((row) => row.some((cell) => cell !== ""));

    // Generate HTML table
    let htmlTable = `<table border="1" style="border-collapse: collapse; width: 100%;">`;
    
    // Add data rows
    for (let row of data) {
      htmlTable += "<tr>";
      for (let cell of row) {
        htmlTable += `<td style="padding: 8px; border: 1px solid #ddd;">${cell}</td>`;
      }
      htmlTable += "</tr>";
    }

    htmlTable += "</table>";

    // Write to a unique temporary file
    const tempFilePath = path.join(tempDir, `excel_${uniqueId}.html`);
    await fs.writeFile(tempFilePath, htmlTable, 'utf8');

    return {
      htmlContent: htmlTable,
      tempFilePath
    };
  } catch (error) {
    console.error('Error converting Excel to HTML:', error);
    throw error;
  }
}

module.exports = {
  excelToHTML,
};
