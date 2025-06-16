const fs = require('fs-extra');
const path = require('path');
const { extractImagesFromPdf } = require('pdf-extract-image');
const crypto = require('crypto');

// Create a unique temporary directory for each extraction
async function createTempDir() {
    const tempDir = path.join('temp', `pdf_extract_${crypto.randomBytes(8).toString('hex')}`);
    await fs.ensureDir(tempDir);
    return tempDir;
}

// Clean up temporary directory
async function cleanupTempDir(tempDir) {
    try {
        await fs.remove(tempDir);
    } catch (error) {
        console.error('Error cleaning up temporary directory:', error);
    }
}

async function extractImagesFromPDF(pdfPath, baseFilename) {
    let tempDir = null;
    try {
        // Create temporary directory for this extraction
        tempDir = await createTempDir();
        console.log(`Created temporary directory: ${tempDir}`);

        // Extract images from PDF
        const images = await extractImagesFromPdf(pdfPath);
        const extractedImages = [];

        // Save each image with a unique name
        for (let i = 0; i < images.length; i++) {
            const imagePath = path.join(tempDir, `${baseFilename}_image_${i + 1}.png`);
            await fs.writeFile(imagePath, images[i]);
            
            const relativePath = path.relative(process.cwd(), imagePath);
            extractedImages.push({
                path: relativePath,
                type: 'png'
            });
        }

        console.log(`Successfully extracted ${extractedImages.length} images from PDF`);
        return extractedImages;
    } catch (error) {
        console.error('Error in PDF image extraction:', error);
        throw error;
    } finally {
        // Clean up temporary directory
        if (tempDir) {
            await cleanupTempDir(tempDir);
        }
    }
}

module.exports = {
    extractImagesFromPDF
};