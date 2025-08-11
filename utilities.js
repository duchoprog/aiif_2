const fs = require("fs");
const fsPromises = fs.promises;
const path = require("path");
const { setTimeout } = require("timers");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");

let folderPath = "newproject";

// Function to save the file buffer to the uploads directory
async function saveFileToUploads(file, sessionID) {
  const uploadsDir = path.join(__dirname, "sessions", sessionID, "uploads");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`1 File saved to ${filePath}`);
}

async function saveFileToFiles(file, sessionID) {
  console.log("saving file");
  const uploadsDir = path.join(__dirname, "sessions", sessionID, "files");

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, file.originalname);

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`2 File saved to ${filePath}`);
}

async function savePrevFileToExcelBase(file, sessionID) {
  console.log("saving file");
  const files = await fsPromises.readdir(path.join(__dirname, "sessions", sessionID));
  console.log("files", files);

  const uploadsDir = path.join(__dirname, "sessions", sessionID, `excelBase`);
  console.log(uploadsDir);

  // Ensure the uploads directory exists
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  // Construct the full path for the file
  const filePath = path.join(uploadsDir, "addInfoToThis.xlsx");

  // Write the file buffer to the uploads directory
  await fs.writeFileSync(filePath, file.buffer);

  console.log(`3 File saved to ${filePath}`);
}
async function writeOutputToExcel(responseArray, res, projectName, sessionID) {
  console.log("Starting writeOutputToExcel");
  console.log("responseArray:", responseArray);
  console.log("projectName:", projectName);
  console.log("sessionID:", sessionID);

  // Process the data
  console.log("About to process data");
  const processedData = await processData(responseArray);
  console.log("Processed data:", processedData);

  // Select starting workbook
  console.log("Selecting workbook");
  const filePath = fs.existsSync(`./sessions/${sessionID}/excelBase/addInfoToThis.xlsx`)
    ? `./sessions/${sessionID}/excelBase/addInfoToThis.xlsx`
    : "INQUIRY 2024 TEMPLATE v4 pablo2.xlsx";
  console.log("Using file path:", filePath);

  const workbook = new ExcelJS.Workbook();
  console.log("Reading workbook");
  await workbook.xlsx.readFile(filePath);
  console.log("Workbook read successfully");

  // Get the first sheet
  const sheetName = workbook.worksheets[0].name;
  const worksheet = workbook.getWorksheet(sheetName);
  //add project id
  worksheet.getCell("H1").value = projectName;
  console.log("Added project name to cell H1");

  // Define the starting row for the new data */
  let startRow = 1;
  let isRowEmpty = false;
  console.log("Finding first empty row");

  while (!isRowEmpty) {
    isRowEmpty = true;
    const row = worksheet.getRow(startRow);

    for (let col = 1; col <= worksheet.columnCount; col++) {
      const cell = row.getCell(col);
      if (cell.value !== null && !cell.formula) {
        isRowEmpty = false;
        break;
      }
    }

    if (!isRowEmpty) {
      startRow++;
    }
  }
  console.log("First empty row found at:", startRow);

  // Define column mappings for each property
  const columnMappings = {
    'item': 3,
    'description': 4,
    'unit': 5,
    'quantity': 6,
    'unitPrice': 9,
    'totalPrice': 10,
    'currency': 11,
    'deliveryTime': 12,
    'paymentTerms': 13,
    'warranty': 14,
    'manufacturer': 15,
    'countryOfOrigin': 16,
    'modelNumber': 17,
    'serialNumber': 18,
    'specifications': 19,
    'dimensions': 20,
    'weight': 21,
    'material': 22,
    'color': 23,
    'brand': 24,
    'category': 25,
    'subcategory': 26,
    'condition': 27,
    'notes': 41,
    'supplier': 46,
    'supplierContact': 47,
    'supplierEmail': 48,
    'supplierPhone': 49,
    'supplierAddress': 50,
    'supplierWebsite': 51,
    'supplierFax': 52,
    'supplierTaxId': 53,
    'supplierBankDetails': 54,
    'supplierPaymentTerms': 55
  };

  console.log("Starting to fill data");
  processedData.forEach((rowData, index) => {
    console.log("Processing rowData:", rowData);
    const row = worksheet.getRow(startRow + index);

    // Map each property to its corresponding column
    Object.entries(rowData).forEach(([property, value]) => {
      const columnIndex = columnMappings[property];
      if (columnIndex) {
        const cell = row.getCell(columnIndex);
        // Check if the cell contains a formula
        if (!cell.formula) {
          cell.value = value;
        } else {
          // Handle shared formulas by copying the formula from the master cell
          const masterCell = worksheet.getCell(startRow + index - 1, columnIndex);
          if (masterCell.formula) {
            cell.formula = masterCell.formula;
          }
        }
      }
    });

    row.commit();
  });
  console.log("Data filling completed");

  // Preserve column widths
  const columnWidths = worksheet.columns.map((col) => col.width);
  worksheet.columns.forEach((col, index) => {
    col.width = columnWidths[index];
  });

  //hide columns
  let hiddenCols = [
    "AC", "AD", "AE", "AF", "AG", "AJ", "AK", "AL", "AM", "AN",
    "AP", "AQ", "AR", "AS",
  ];

  for (col of hiddenCols) {
    let colToHide = worksheet.getColumn(col);
    colToHide.hidden = true;
  }

  // Write the workbook back to the file
  var d = new Date();
  d = d.getTime().toString();
  const outputPath = `./sessions/${sessionID}/output/${projectName}-${d}.xlsx`;
  console.log("Writing workbook to:", outputPath);
  await workbook.xlsx.writeFile(outputPath);
  console.log("Workbook written successfully");

  setTimeout(() => {
    console.log("nice wait");
    res.json({
      success: true,
      redirectUrl: `/download?id=${sessionID}&projectName=${projectName}`,
      outputPath: outputPath
    });
  }, 5000);
  console.log("wating...");
}

//DELETE ALL FILES IN FOLDER
async function deleteAllFilesInDir(dirPath) {
  try {
    console.log("Delete all files deleting ", dirPath);
    fs.readdirSync(dirPath).forEach((file) => {
      console.log("deleting", `${dirPath}${file}`);
      fs.rmSync(path.join(dirPath, file));
    });
  } catch (error) {
    console.log(error);
  }
}
async function deleteOneFile(file) {
  if (fs.existsSync(file)) {
    try {
      fs.unlinkSync(file);
      console.log("File deleted successfully");
    } catch (error) {
      console.error("Error deleting the file:", error);
    }
  } else {
    console.log("File does not exist.");
  }
}
//CREATE FOLDER
async function createFolder(name) {
  var d = new Date();
  d = d.getTime().toString();
  folderPath = path.join(__dirname, `projects/${d}${name}`);

  try {
    fs.mkdirSync(folderPath);
    console.log("Folder ", folderPath, " created successfully!");
    return folderPath;
  } catch (err) {
    console.error("Error creating folder:", err);
  }
}
//DELETE FOLDER
async function deleteFolder(folder) {
  const directoryPath = path.resolve(__dirname, folder);

  // Check if the directory exists
  if (fs.existsSync(directoryPath)) {
    // Delete the directory if it exists
    try {
      fs.rmSync(directoryPath, { recursive: true });
      console.log(`${directoryPath} is deleted!`);
    } catch (err) {
      console.error(`Error while deleting ${directoryPath}.`, err);
    }
  } else {
    console.log(`${directoryPath} does not exist.`);
  }
}

// Function to process the data
async function processData(responseArray) {
  let allData = [];
  //console.log("responsearray,", responseArray);

  responseArray.forEach((item) => {
    // Remove any content after the closing bracket '}]' but keep the closing single quote
    const regex = /【[^【】]*】/g;
    if (item.openaiResponse) {
      console.log("item de responseArray", item.openaiResponse);

      let cleanedResponse = item.openaiResponse.replace(regex, "");
      cleanedResponse = cleanText(cleanedResponse);
      // Parse the JSON data
      let jsonData = JSON.parse(cleanedResponse);
      allData = allData.concat(jsonData);
    }
  });

  return allData;
}

async function manageFolders(sessionID) {
  const folders = [
    "images",
    "uploads",
    "output",
    "imageVault",
    "files",
    "excelBase",
  ];
  console.log("sessionID", sessionID);
  const folderPath = path.join(__dirname, "sessions", sessionID);
  console.log(folderPath);

  await fsPromises.mkdir(folderPath, { recursive: true });

  for (const folderName of folders) {
    const folderPath = path.resolve(__dirname, "sessions", sessionID, folderName);

    // Create the folder
    await fsPromises.mkdir(folderPath, { recursive: true });
  }
}

function cleanText(dirtyText) {
  //console.log("dirtyText", dirtyText);

  // Step 1: Extract substrings between a colon and a comma or a closing curly bracket
  const regex = /:\s*([^,}]*)[,\}]/g;
  let match;
  let cleanedText = dirtyText;

  while ((match = regex.exec(dirtyText)) !== null) {
    if ((match[1].match(/"/g) || []).length > 2) {
      console.log("mal! ", match[1]);
      const split = match[1].split('"');
      let singleQuote = `${split.join("'")}`;
      singleQuote = `"${singleQuote.slice(1, -1).replace(/\\/g, "")}"`;
      cleanedText = cleanedText.replace(match[1], singleQuote);
      //console.log("cleanedText", cleanedText);
    }
  }
  //console.log("cleanedText!", cleanedText);

  return cleanedText;
}

// Function to check if a folder is older than 2 days
function isOlderThanTwoDays(folderPath) {
  const currentTime = new Date();
  const stats = fs.statSync(folderPath);
  const folderTime = new Date(stats.mtime);
  const timeDifference = currentTime - folderTime;
  const twoDaysInMilliseconds = 2 * 24 * 60 * 60 * 1000;
  return timeDifference > twoDaysInMilliseconds;
}

// Function to delete a folder
function deleteFolder(folderPath) {
  if (fs.existsSync(folderPath)) {
    fs.rmdirSync(folderPath, { recursive: true });
    console.log(`Deleted folder: ${folderPath}`);
  }
}

// Function to check if a folder name matches the pattern
function isMatchingFolderName(folderName) {
  return /^17\d{11}$/.test(folderName);
}

// Main function to traverse the directory and delete old folders
function deleteOldFolders(dir) {
  console.log("deleteOldFolders");

  fs.readdirSync(dir).forEach((file) => {
    const filePath = path.join(dir, file);
    const stats = fs.statSync(filePath);

    if (stats.isDirectory()) {
      console.log("stats:", filePath);

      if (isMatchingFolderName(file) && isOlderThanTwoDays(filePath)) {
        deleteFolder(filePath);
      }
    }
  });
}

///function to log memory usage
function logMemoryUsage(stage) {
  const memoryUsage = process.memoryUsage();
  console.log(`\nMemory usage at ${stage}:`);
  console.log(`  RSS: ${(memoryUsage.rss / 1024 / 1024).toFixed(2)} MB`);
  console.log(
    `  Heap Total: ${(memoryUsage.heapTotal / 1024 / 1024).toFixed(2)} MB`
  );
  console.log(
    `  Heap Used: ${(memoryUsage.heapUsed / 1024 / 1024).toFixed(2)} MB`
  );
  console.log(
    `  External: ${(memoryUsage.external / 1024 / 1024).toFixed(2)} MB`
  );
  console.log(
    `  Array Buffers: ${(memoryUsage.arrayBuffers / 1024 / 1024).toFixed(2)} MB`
  );
}

async function groupAndWriteToExcel(jsonObjects, res, projectName, sessionID) {
  try {
    // Group all JSON objects into a single array
    const groupedData = jsonObjects.reduce((acc, obj) => {
      if (obj.openaiResponse) {
        // Clean and parse the response
        const cleanedResponse = cleanText(obj.openaiResponse);
        const parsedData = JSON.parse(cleanedResponse);
        return acc.concat(parsedData);
      }
      return acc;
    }, []);

    // Write the grouped data to Excel
    await writeOutputToExcel(groupedData, res, projectName, sessionID);
    
    return true;
  } catch (error) {
    console.error('Error in groupAndWriteToExcel:', error);
    throw error;
  }
}

module.exports = {
  deleteOldFolders,
  deleteAllFilesInDir,
  deleteFolder,
  saveFileToUploads,
  writeOutputToExcel,
  processData,
  manageFolders,
  saveFileToFiles,
  savePrevFileToExcelBase,
  deleteOneFile,
  createFolder,
  cleanText,
  logMemoryUsage,
  groupAndWriteToExcel,
};
