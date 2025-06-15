const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const logger = require('./logger');
const crypto = require('crypto');
const { PDFDocument } = require('pdf-lib');
const express = require('express');
const multer = require('multer');
const fsExtra = require('fs-extra');
const JSZip = require('jszip');

// Create images directory if it doesn't exist
const imagesDir = 'images';
try {
    fs.mkdirSync(imagesDir, { recursive: true });
    logger.info('Ensured images directory exists');
} catch (error) {
    if (error.code !== 'EEXIST') {
        logger.error('Error creating images directory:', error);
        throw error;
    }
}

/**
 * Generates a unique filename for an image
 * @param {string} baseFilename - Base filename from the original document
 * @param {string} sheetName - Sheet name (for Excel files)
 * @param {string} index - Index or key of the image
 * @param {string} ext - File extension
 * @returns {string} Unique filename
 */
function generateUniqueFilename(baseFilename, sheetName, index, ext) {
    const timestamp = Date.now();
    const random = crypto.randomBytes(4).toString('hex');
    const sheetPart = sheetName ? `_${sheetName}` : '';
    console.log("name ", `${baseFilename}${sheetPart}_${index}_${timestamp}_${random}.${ext}`)
    return `${baseFilename}${sheetPart}_${index}_${timestamp}_${random}.${ext}`;
}

/**
 * Safely writes a file using atomic operations
 * @param {string} filePath - Path where to write the file
 * @param {Buffer} data - File data to write
 */
function safeWriteFile(filePath, data) {
    logger.info(`Attempting to write file: ${filePath}`);
    logger.info(`Data size: ${data.length} bytes`);
    
    const tempPath = `${filePath}.${crypto.randomBytes(4).toString('hex')}.tmp`;
    try {
        logger.info(`Writing to temp file: ${tempPath}`);
        fs.writeFileSync(tempPath, data);
        logger.info(`Successfully wrote temp file`);
        
        logger.info(`Moving temp file to final location: ${filePath}`);
        fs.renameSync(tempPath, filePath);
        logger.info(`Successfully moved file to final location`);
    } catch (error) {
        logger.error(`Error writing file ${filePath}:`, error);
        // Clean up temp file if it exists
        try {
            if (fs.existsSync(tempPath)) {
                fs.unlinkSync(tempPath);
                logger.info(`Cleaned up temp file: ${tempPath}`);
            }
        } catch (cleanupError) {
            logger.error('Error cleaning up temp file:', cleanupError);
        }
        throw error;
    }
}

/**
 * Extracts images from various file types (XLSX, PDF, DOCX)
 * @param {string} filePath - Path to the source file
 * @param {string} originalFilename - Original filename of the source file
 * @returns {Promise<Array<{filename: string, path: string}>>} Array of extracted image information
 */
async function extractImages(filePath, originalFilename) {
    const fileExt = path.extname(filePath).toLowerCase();
    console.log("extracting ", fileExt);
    
    const baseFilename = path.basename(originalFilename, fileExt);
    const images = [];

    logger.info(`Starting image extraction for ${originalFilename} (${fileExt})`);
    logger.info(`File path: ${filePath}`);
    logger.info(`Base filename: ${baseFilename}`);

    try {
        if (fileExt === '.xlsx' || fileExt === '.xls') {
            logger.info('Processing Excel file...');
            const workbook = XLSX.readFile(filePath);
            logger.info(`Found ${workbook.SheetNames.length} sheets`);
            
            for (const sheetName of workbook.SheetNames) {
                const sheet = workbook.Sheets[sheetName];
                logger.info(`Processing sheet: ${sheetName}`);
                
                if (sheet['!images']) {
                    logger.info(`Found ${Object.keys(sheet['!images']).length} images in sheet: ${sheetName}`);
                    for (const [key, image] of Object.entries(sheet['!images'])) {
                        logger.info(`Processing image ${key} from sheet ${sheetName}`);
                        const imageExt = image.type.split('/')[1] || 'png';
                        const imageFilename = generateUniqueFilename(baseFilename, sheetName, key, imageExt);
                        const imagePath = path.join('images', imageFilename);
                        logger.info(`Generated image path: ${imagePath}`);
                        safeWriteFile(imagePath, image.data);
                        images.push({
                            filename: imageFilename,
                            path: imagePath
                        });
                        logger.info(`Successfully extracted image: ${imageFilename}`);
                    }
                } else {
                    logger.info(`No images found in sheet: ${sheetName}`);
                }
            }
        } else if (fileExt === '.pdf') {
            logger.info('Processing PDF file...');
            // Read the PDF file
            const pdfBytes = fs.readFileSync(filePath);
            logger.info(`Read PDF file, size: ${pdfBytes.length} bytes`);
            
            const pdfDoc = await PDFDocument.load(pdfBytes);
            logger.info('Successfully loaded PDF document');
            
            // Get all pages
            const pages = pdfDoc.getPages();
            logger.info(`Found ${pages.length} pages in PDF`);
            
            // Extract images from each page
            for (let i = 0; i < pages.length; i++) {
                const page = pages[i];
                logger.info(`Processing page ${i + 1}`);
                
                // Get all XObjects (which include images) from the page
                const xObjects = page.node.Resources().lookup(PDFName.of('XObject'), PDFDict);
                if (xObjects) {
                    logger.info(`Found XObjects on page ${i + 1}`);
                    for (const [name, xObject] of Object.entries(xObjects.dict)) {
                        if (xObject instanceof PDFStream) {
                            const subtype = xObject.dict.get(PDFName.of('Subtype'));
                            if (subtype && subtype.toString() === '/Image') {
                                try {
                                    logger.info(`Found image object: ${name}`);
                                    // Get image data
                                    const imageData = xObject.getContents();
                                    if (imageData) {
                                        logger.info(`Got image data, size: ${imageData.length} bytes`);
                                        // Determine image type
                                        const filter = xObject.dict.get(PDFName.of('Filter'));
                                        let imageExt = 'png';
                                        if (filter) {
                                            if (filter.toString().includes('/DCTDecode')) {
                                                imageExt = 'jpg';
                                            } else if (filter.toString().includes('/FlateDecode')) {
                                                imageExt = 'png';
                                            }
                                        }
                                        
                                        const imageFilename = generateUniqueFilename(baseFilename, `page${i + 1}`, name, imageExt);
                                        const imagePath = path.join('images', imageFilename);
                                        logger.info(`Generated image path: ${imagePath}`);
                                        safeWriteFile(imagePath, imageData);
                                        images.push({
                                            filename: imageFilename,
                                            path: imagePath
                                        });
                                        logger.info(`Successfully extracted image: ${imageFilename}`);
                                    } else {
                                        logger.warn(`No image data found for object: ${name}`);
                                    }
                                } catch (error) {
                                    logger.error(`Error extracting image from PDF page ${i + 1}:`, error);
                                }
                            }
                        }
                    }
                } else {
                    logger.info(`No XObjects found on page ${i + 1}`);
                }
            }
        } else if (fileExt === '.docx' || fileExt === '.doc') {
            logger.info('Processing Word document...');
            try {
                logger.info('Starting mammoth extraction...');
                const result = await mammoth.extractRawText({ path: filePath });
                
                logger.info('Mammoth extraction completed');
                
                if (result.images && Object.keys(result.images).length > 0) {
                    logger.info(`Found ${Object.keys(result.images).length} images in document`);
                    
                    // Save each image
                    for (const [imageId, image] of Object.entries(result.images)) {
                        logger.info(`Processing image ${imageId}`);
                        logger.info(`Image content type: ${image.contentType}`);
                        
                        const imageExt = image.contentType.split('/')[1] || 'png';
                        const imageFilename = generateUniqueFilename(baseFilename, '', imageId, imageExt);
                        const imagePath = path.join('images', imageFilename);
                        
                        logger.info(`Saving image to: ${imagePath}`);
                        safeWriteFile(imagePath, image.buffer);
                        
                        images.push({
                            filename: imageFilename,
                            path: imagePath
                        });
                        logger.info(`Successfully extracted image: ${imageFilename}`);
                    }
                } else {
                    logger.info('No images found in document');
                }
                
                if (result.messages && result.messages.length > 0) {
                    logger.info('Mammoth messages:', result.messages);
                }
            } catch (error) {
                logger.error(`Error processing Word document:`, error);
                throw error;
            }
        }

        logger.info(`Completed image extraction for ${originalFilename}. Found ${images.length} images.`);
        return images;
    } catch (error) {
        logger.error(`Error extracting images from ${originalFilename}:`, error);
        // Clean up any partially written images
        for (const image of images) {
            try {
                fs.unlinkSync(image.path);
            } catch (cleanupError) {
                logger.error(`Error cleaning up image ${image.path}:`, cleanupError);
            }
        }
        throw error;
    }
}

async function extractImagesFromZip(filePath, baseFilename, fileExt) {
    const images = [];
    try {
        logger.info(`Starting ZIP-based extraction for ${fileExt}...`);
        const data = await fsExtra.readFile(filePath);
        const zip = await JSZip.loadAsync(data);
        
        // Define image paths based on file type
        let imagePaths = [];
        if (fileExt === '.docx') {
            // DOCX images are in word/media/
            imagePaths = Object.keys(zip.files).filter(filename => 
                filename.startsWith('word/media/') && 
                !filename.endsWith('/')
            );
        } else if (fileExt === '.xlsx') {
            // XLSX images are in xl/media/
            imagePaths = Object.keys(zip.files).filter(filename => 
                filename.startsWith('xl/media/') && 
                !filename.endsWith('/')
            );
        }

        logger.info(`Found ${imagePaths.length} potential images in ${fileExt}`);
        
        // Process each image
        for (let i = 0; i < imagePaths.length; i++) {
            const imagePath = imagePaths[i];
            const imageFile = zip.files[imagePath];
            
            if (!imageFile.dir) {
                const imageExt = path.extname(imagePath).slice(1) || 'png';
                const imageFilename = generateUniqueFilename(baseFilename, '', i, imageExt);
                const outputPath = path.join('images', imageFilename);
                
                logger.info(`Processing image ${i + 1}/${imagePaths.length}: ${imagePath}`);
                logger.info(`Saving as: ${outputPath}`);
                
                const imageBuffer = await imageFile.async('nodebuffer');
                safeWriteFile(outputPath, imageBuffer);
                
                images.push({
                    filename: imageFilename,
                    path: outputPath
                });
                logger.info(`Successfully extracted image: ${imageFilename}`);
            }
        }
        
        logger.info(`Completed ZIP-based extraction. Found ${images.length} images.`);
        return images;
    } catch (error) {
        logger.error('Error in ZIP-based extraction:', error);
        throw error;
    }
}

async function extractImages(filePath, baseFilename, fileExt) {
    logger.info(`Starting image extraction for ${baseFilename}${fileExt}`);
    logger.info(`File path: ${filePath}`);
    logger.info(`Base filename: ${baseFilename}`);
    
    const images = [];
    
    try {
        if (fileExt === '.pdf') {
            // ... existing PDF code ...
        } else if (fileExt === '.docx' || fileExt === '.xlsx') {
            logger.info(`Processing ${fileExt.toUpperCase()} document...`);
            try {
                const data = await fsExtra.readFile(filePath);
                const zip = await JSZip.loadAsync(data);
                
                // Define image paths based on file type
                let imagePaths = [];
                if (fileExt === '.docx') {
                    // DOCX images are in word/media/
                    imagePaths = Object.keys(zip.files).filter(filename => 
                        filename.startsWith('word/media/') && 
                        !filename.endsWith('/')
                    );
                } else if (fileExt === '.xlsx') {
                    // XLSX images are in xl/media/
                    imagePaths = Object.keys(zip.files).filter(filename => 
                        filename.startsWith('xl/media/') && 
                        !filename.endsWith('/')
                    );
                }

                logger.info(`Found ${imagePaths.length} potential images in ${fileExt}`);
                
                // Process each image
                for (let i = 0; i < imagePaths.length; i++) {
                    const imagePath = imagePaths[i];
                    const imageFile = zip.files[imagePath];
                    
                    if (!imageFile.dir) {
                        const imageExt = path.extname(imagePath).slice(1) || 'png';
                        const imageFilename = generateUniqueFilename(baseFilename, '', i, imageExt);
                        const outputPath = path.join('images', imageFilename);
                        
                        logger.info(`Processing image ${i + 1}/${imagePaths.length}: ${imagePath}`);
                        logger.info(`Saving as: ${outputPath}`);
                        
                        const imageBuffer = await imageFile.async('nodebuffer');
                        safeWriteFile(outputPath, imageBuffer);
                        
                        images.push({
                            filename: imageFilename,
                            path: outputPath
                        });
                        logger.info(`Successfully extracted image: ${imageFilename}`);
                    }
                }
                
                if (images.length === 0) {
                    logger.info(`No images found in ${fileExt} document`);
                }
            } catch (error) {
                logger.error(`Error processing ${fileExt}:`, error);
                throw error;
            }
        } else {
            throw new Error(`Unsupported file type: ${fileExt}`);
        }
        
        logger.info(`Completed image extraction for ${baseFilename}${fileExt}. Found ${images.length} images.`);
        return images;
    } catch (error) {
        logger.error(`Error extracting images from ${baseFilename}${fileExt}:`, error);
        throw error;
    }
}

module.exports = {
    extractImages
}; 