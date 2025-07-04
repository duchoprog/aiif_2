const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

async function writeToExcel(data, projectName, sessionOutputDir = null) {
    try {
        console.log('Starting Excel write process');
        console.log('Data received:', data);
        console.log('Data type:', typeof data);
        console.log('Data length:', Array.isArray(data) ? data.length : 'Not an array');
        console.log('Data structure:', JSON.stringify(data, null, 2));
        
        // Set default paths
        let templatePath;
        if (sessionOutputDir) {
            // First try session's excelBase
            const sessionTemplatePath = path.join(path.dirname(sessionOutputDir), 'excelBase', 'INQUIRY 2024 TEMPLATE v4 pablo2.xlsx');
            const rootTemplatePath = './excelBase/INQUIRY 2024 TEMPLATE v4 pablo2.xlsx';
            
            // Check if session template exists, otherwise use root template
            if (fs.existsSync(sessionTemplatePath)) {
                templatePath = sessionTemplatePath;
                console.log(`Using session template: ${templatePath}`);
            } else {
                templatePath = rootTemplatePath;
                console.log(`Session template not found, using root template: ${templatePath}`);
            }
        } else {
            templatePath = './excelBase/INQUIRY 2024 TEMPLATE v4 pablo2.xlsx';
        }
        
        const timestamp = new Date().getTime();
        const outputDir = sessionOutputDir || './output';
        const outputPath = path.join(outputDir, `${projectName}-${timestamp}.xlsx`);
        
        // Ensure output directory exists
        if (!fs.existsSync(outputDir)) {
            fs.mkdirSync(outputDir, { recursive: true });
            console.log(`Created output directory: ${outputDir}`);
        }
        
        // Load the template workbook
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.readFile(templatePath);
        
        // Get the first worksheet
        const worksheet = workbook.getWorksheet(1);
        
        // Add project name to cell H1
        worksheet.getCell('H1').value = projectName;
        
        // Define the columns to fill in order
        const columnsToFill = [3, 4, 5, 6, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24,
            25, 26, 27, 41, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55];
        
        // Find first empty row
        let startRow = 1;
        let isRowEmpty = false;
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
        
        console.log(`First empty row found at: ${startRow}`);
        //console.log("data", data);
        
        // Check if data is empty or invalid
        if (!data || !Array.isArray(data) || data.length === 0) {
            console.error('Data is empty or invalid. Cannot write to Excel.');
            throw new Error('No data provided to write to Excel');
        }
        
        // Flatten the data array - data is an array of arrays (one array per document)
        const flattenedData = data.flat();
        //console.log("flattenedData", flattenedData);
        console.log("flattenedData length:", flattenedData.length);
        
        if (flattenedData.length === 0) {
            console.error('Flattened data is empty. Cannot write to Excel.');
            throw new Error('No valid data after flattening');
        }
        
        // Write each JSON object to a new row
        flattenedData.forEach((jsonObj, index) => {
            console.log(`Processing row ${index + 1}:`, jsonObj);
            const row = worksheet.getRow(startRow + index);
            
            // First, write non-image properties
            const nonImageProps = Object.entries(jsonObj).filter(([key]) => !key.startsWith('IMAGE '));
            nonImageProps.forEach(([_, value], propIndex) => {
                if (propIndex < columnsToFill.length) {
                    const cell = row.getCell(columnsToFill[propIndex]);
                    if (!cell.formula) {
                        cell.value = value;
                    } else {
                        // Handle shared formulas
                        const masterCell = worksheet.getCell(startRow + index - 1, columnsToFill[propIndex]);
                        if (masterCell.formula) {
                            cell.formula = masterCell.formula;
                        }
                    }
                }
            });
            
            // Then handle image properties
            const imageProps = Object.entries(jsonObj).filter(([key]) => key.startsWith('IMAGE '));
            imageProps.forEach(([key, value]) => {
                if (value && typeof value === 'string') {
                    // Find the corresponding column for this image
                    const imageNumber = parseInt(key.split(' ')[1]);
                    const imageColumn = columnsToFill[imageNumber + 6]; // Adjust this offset based on your template
                    if (imageColumn) {
                        const cell = row.getCell(imageColumn);
                        cell.value = value; // Store the image filename
                    }
                }
            });
            
            row.commit();
        });
        
        // Hide specific columns
        const hiddenColumns = [
            'AC', 'AD', 'AE', 'AF', 'AG', 'AJ', 'AK', 'AL', 'AM', 'AN',
            'AP', 'AQ', 'AR', 'AS'
        ];
        
        hiddenColumns.forEach(col => {
            worksheet.getColumn(col).hidden = true;
        });
        
        // Save the workbook
        await workbook.xlsx.writeFile(outputPath);
        console.log(`Excel file written successfully to: ${outputPath}`);
        
        return outputPath;
    } catch (error) {
        console.error('Error writing to Excel:', error);
        throw error;
    }
}

module.exports = {
    writeToExcel
};