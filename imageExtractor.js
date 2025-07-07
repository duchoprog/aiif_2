const path = require('path');
const fs = require('fs-extra');
const JSZip = require('jszip');



// Helper function to generate unique filenames
function generateUniqueFilename(baseFilename, prefix = '', index = 0, extension = 'png') {
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(2, 8);
    return `${baseFilename}-${prefix}${index}-${timestamp}-${random}.${extension}`;
}

// Helper function to safely write files
function safeWriteFile(filePath, data) {
    try {
        fs.writeFileSync(filePath, data);
    } catch (error) {
        console.error(`Error writing file ${filePath}:`, error);
        throw error;
    }
}

async function extractImages(filePath, baseFilename, fileExt, outputDir = 'images') {
    console.log(`Starting image extraction for ${baseFilename}${fileExt}`);
    console.log(`File path: ${filePath}`);
    console.log(`Base filename: ${baseFilename}`);
    
    const images = [];
    
    try {
        if (fileExt === '.pdf') {
            // ... existing PDF code ...
        } else if (fileExt === '.docx' || fileExt === '.xlsx') {
            console.log(`Processing ${fileExt.toUpperCase()} document...`);
            try {
                // Create a temporary directory for extraction
                const tempDir = path.join('temp', Date.now().toString());
                await fs.ensureDir(tempDir);
                console.log(`Created temporary directory: ${tempDir}`);

                // Read the file and extract it
                const data = await fs.readFile(filePath);
                const zip = await JSZip.loadAsync(data);
                
                // Extract all files
                console.log('Extracting files from archive...');
                await Promise.all(
                    Object.keys(zip.files).map(async (filename) => {
                        console.log("filename", filename, "baseFilename", baseFilename );
                        const file = zip.files[filename];
                        if (!file.dir) {
                            const content = await file.async('nodebuffer');
                            const fullPath = path.join(tempDir, filename);
                            await fs.ensureDir(path.dirname(fullPath));
                            await fs.writeFile(fullPath, content);
                        }
                    })
                );
                console.log('Files extracted successfully');

                // Define media directory based on file type
                const mediaDir = fileExt === '.docx' ? 'word/media' : 'xl/media';
                const fullMediaPath = path.join(tempDir, mediaDir);
                
                // Check if media directory exists
                if (await fs.pathExists(fullMediaPath)) {
                    console.log(`Found media directory: ${fullMediaPath}`);
                    
                    // Get all files in media directory
                    const files = await fs.readdir(fullMediaPath);
                    console.log(`Found ${files.length} files in media directory`);
                    
                    // Process each file
                    for (let i = 0; i < files.length; i++) {
                        const file = files[i];
                        const sourcePath = path.join(fullMediaPath, file);
                        const stats = await fs.stat(sourcePath);
                        
                        if (stats.isFile()) {
                            // Get the original file extension from the source file
                            const originalExt = path.extname(file).toLowerCase();
                            const imageExt = originalExt 
                            const imageFilename = generateUniqueFilename(baseFilename, '', i, imageExt);
                            const outputPath = path.join(outputDir, imageFilename);
                            
                            console.log(`Processing image ${i + 1}/${files.length}: ${file}`);
                            console.log(`Original extension: ${originalExt}`);
                            console.log(`Using extension: ${imageExt}`);
                            console.log(`Saving as: ${outputPath}`);
                            
                            await fs.copy(sourcePath, outputPath);
                            
                            images.push({
                                filename: imageFilename,
                                path: outputPath
                            });
                            console.log(`Successfully extracted image: ${imageFilename}`);
                        }
                    }
                } else {
                    console.log(`No media directory found at: ${fullMediaPath}`);
                }
                
                // Clean up temporary directory
                console.log('Cleaning up temporary directory...');
                await fs.remove(tempDir);
                console.log('Cleanup complete');
                
                if (images.length === 0) {
                    console.log(`No images found in ${fileExt} document`);
                }
            } catch (error) {
                console.error(`Error processing ${fileExt}:`, error);
                throw error;
            }
        } else {
            throw new Error(`Unsupported file type: ${fileExt}`);
        }
        
        console.log(`Completed image extraction for ${baseFilename}${fileExt}. Found ${images.length} images.`);
        return images;
    } catch (error) {
        console.error(`Error extracting images from ${baseFilename}${fileExt}:`, error);
        throw error;
    }
}

// Export the function
module.exports = { extractImages }; 