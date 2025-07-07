const ExcelJS = require("exceljs");
const fs = require("fs");
const fsp = require("fs").promises;
const path = require("path");
const { getLastModifiedFile } = require("./lastFile.js");
const { logMemoryUsage } = require("./utilities.js");

async function replaceImages(req) {
  // Define the directories
  const outputDir = path.join(__dirname, "sessions", req.body.sessionID, "output");
  const imagesDir = path.join(__dirname, "sessions", req.body.sessionID, "imageVault");

  /* let originalFile = await getLastModifiedFile(
    `./sessions/${req.body.sessionID}/output`
  ); */

  try {
    await fsp.access(outputDir, fs.constants.F_OK);
    console.log(`Directory exists: ${outputDir}`);
  } catch (error) {
    console.error(`Directory does not exist: ${outputDir}`);
    return;
  }

  // Read the directory
  let files;
  try {
    files = await fsp.readdir(outputDir);
    console.log(`Files in directory: ${files}`);
  } catch (error) {
    console.error(`Error reading directory: ${outputDir}`, error);
    return;
  }

  // Check if the directory is empty
  if (files.length === 0) {
    console.error(`Directory is empty: ${outputDir}`);
    return;
  }

  // Get the first file in the directory
  let originalFile = files[0];

  console.log("originalFile", originalFile, "dir:", outputDir);

  // Construct the full path to the file
  const filePath = path.join(outputDir, originalFile);

  // Read the spreadsheet
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(filePath);
    console.log(`Successfully read file: ${filePath}`);
    // Continue with your logic to replace images
  } catch (error) {
    console.error(`Error reading file: ${filePath}`, error);
  }
  const worksheet = workbook.getWorksheet(1);

  // Function to find and replace JPEG filenames with corresponding images
  async function replaceJpegFilenamesWithImages() {
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        if (
          cell.value &&
          typeof cell.value === "string" &&
          (cell.value.toLowerCase().endsWith(".jpeg") ||
            cell.value.toLowerCase().endsWith(".jpg") ||
            cell.value.toLowerCase().endsWith(".png") ||
            cell.value.toLowerCase().endsWith(".gif") ||
            cell.value.toLowerCase().endsWith(".tif") ||
            cell.value.toLowerCase().endsWith(".tiff") ||
            cell.value.toLowerCase().endsWith(".bmp")) //
        ) {
          console.log("tengo que reemplazar ", cell.value, "col:", colNumber, "row:", rowNumber);

          const imagePath = path.join(imagesDir, cell.value);
          console.log("de aca:", imagePath);

          if (fs.existsSync(imagePath)) {
            const imageBuffer = fs.readFileSync(imagePath);

            const imageId = workbook.addImage({
              buffer: imageBuffer,
              extension: "jpeg",
            });
            console.log("imageId", imageId);

            cell.value = "";
            worksheet.addImage(imageId, {
              tl: { col: colNumber - 1, row: rowNumber - 1 },
              ext: { width: 100, height: 100 },
            });
            row.height = 100;
            worksheet.getColumn(colNumber).width = 30;
            console.log("completado addImage");
          } else {
            console.log("no encontre la imagen", cell.value);
          }
        }
      });
    });
  }

  
  await replaceJpegFilenamesWithImages();
 

  // Write the updated workbook to a new file
  await workbook.xlsx.writeFile(
    filePath
    /* originalFile.replace(/\.xlsx$/, "_revisado.xlsx") */
  );

  console.log("Spreadsheet updated successfully.");
}

module.exports = { replaceImages };
