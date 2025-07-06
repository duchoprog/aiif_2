require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const OpenAI = require('openai');
const path = require('path');
const fs = require('fs-extra');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const compression = require('compression');
const expressQueue = require('express-queue');
const { extractImages } = require('./imageExtractor');
const pdfParse = require('pdf-parse');
const xlsx = require('xlsx');
const { extractImagesFromPDF } = require('./pdfImageExtractor');
const { excelToHTML } = require('./excelToHTML');
const { writeToExcel } = require('./excelWriter');
// Validate environment variables
const requiredEnvVars = ['OPENAI_API_KEY', 'OPENAI_ASSISTANT_ID'];
const missingEnvVars = requiredEnvVars.filter(envVar => !process.env[envVar]);
if (missingEnvVars.length > 0) {
    console.error(`Missing required environment variables: ${missingEnvVars.join(', ')}`);
    process.exit(1);
}

const app = express();
const port = process.env.PORT || 3000;

let uploadDir = ""
let imagesDir = ""
let tempDir = ""
let outputDir = ""
let excelBaseDir = ""

// Session configuration.
const SESSION_CLEANUP_MINUTES = parseInt(process.env.SESSION_CLEANUP_MINUTES) || 1; // 24 hourswould be 1440
const SESSION_CLEANUP_MS = SESSION_CLEANUP_MINUTES * 60 * 1000;

// Helper function to validate and sanitize session name
function validateSessionName(sessionName) {
    if (!sessionName) {
        return null;
    }
    // Sanitize session name to prevent directory traversal, but allow hyphens
    return sessionName.replace(/[^a-zA-Z0-9-_]/g, '_');
}

// Helper function to generate session name (fallback for backward compatibility)
function generateSessionName(projectName) {
    if (projectName && projectName.trim()) {
        // Remove unsafe characters and trim
        const safeName = projectName.trim().replace(/[^a-zA-Z0-9-_]/g, '');
        // Generate a random 4-digit number
        const random4 = Math.floor(1000 + Math.random() * 9000);
        return `${safeName}${random4}`;
    } else {
        // Generate a random 8-digit number
        const random8 = Math.floor(10000000 + Math.random() * 90000000);
        return `Session${random8}`;
    }
}

// Initialize OpenAI
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY,
    defaultHeaders: {
        'OpenAI-Beta': 'assistants=v2'
    }
});





// Session cleanup function
async function cleanupOldSessions() {
    try {
        const sessionsDir = 'sessions';
        if (!fs.existsSync(sessionsDir)) {
            return;
        }

        const sessionFolders = await fs.readdir(sessionsDir);
        const cutoffTime = Date.now() - SESSION_CLEANUP_MS;
        let cleanedCount = 0;

        for (const folder of sessionFolders) {
            const folderPath = path.join(sessionsDir, folder);
            const stats = await fs.stat(folderPath);
            
            if (stats.isDirectory() && stats.mtime.getTime() < cutoffTime) {
                await fs.remove(folderPath);
                console.info(`Cleaned up old session: ${folder}`);
                cleanedCount++;
            }
        }

        if (cleanedCount > 0) {
            console.info(`Session cleanup completed: ${cleanedCount} old sessions removed`);
        }
    } catch (error) {
        console.error('Error during session cleanup:', error);
    }
}

// Create sessions directory if it doesn't exist
const sessionsDir = 'sessions';
if (!fs.existsSync(sessionsDir)) {
    fs.mkdirSync(sessionsDir);
    console.info('Created sessions directory');
}

// Run initial cleanup
cleanupOldSessions();

// Configure multer for file upload
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        // Extract session name from query parameters
        const sessionName = req.query.sessionName;
        if (!sessionName) {
            return cb(new Error('Session name is required'), null);
        }
        
        console.log('Found session name from query:', sessionName);
        const validatedSessionName = validateSessionName(sessionName);
        console.log('Validated session name:', validatedSessionName);
        if (!validatedSessionName) {
            console.error('Invalid session name after validation:', sessionName);
            return cb(new Error('Invalid session name'), null);
        }
        
        const sessionDir = path.join('sessions', validatedSessionName);
        
        // Create session subdirectories
         uploadDir = path.join(sessionDir, 'uploads');
         imagesDir = path.join(sessionDir, 'images');
         tempDir = path.join(sessionDir, 'temp');
         outputDir = path.join(sessionDir, 'output');
         excelBaseDir = path.join(sessionDir, 'excelBase');
        
        // Create directories if they don't exist
        [sessionDir, uploadDir, imagesDir, tempDir, outputDir, excelBaseDir].forEach(dir => {
            if (!fs.existsSync(dir)) {
                console.log("creating directory: ", dir);
                fs.mkdirSync(dir, { recursive: true });
            }
        });
        
    
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        const uniqueSuffix = Date.now();
        cb(null, uniqueSuffix + '-' + file.originalname);
    }
});

const fileFilter = (req, file, cb) => {
    console.log('Multer fileFilter called with file:', file);
    const allowedTypes = [
        'application/pdf',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/msword',  // .doc
        'application/vnd.ms-excel'  // .xls
    ];
    
    if (allowedTypes.includes(file.mimetype)) {
        console.log('File type accepted:', file.mimetype);
        cb(null, true);
    } else {
        console.log('File type rejected:', file.mimetype);
        cb(new Error('Invalid file type. Only PDF, DOC, DOCX, XLS, and XLSX files are allowed.'), false);
    }
};

const upload = multer({ 
    storage: storage,
    fileFilter: fileFilter,
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Middleware
app.use(helmet());
app.use(compression());
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// Rate limiting
const limiter = rateLimit({
    windowMs: parseInt(process.env.RATE_LIMIT_WINDOW_MS) || 900000, // 15 minutes
    max: parseInt(process.env.RATE_LIMIT_MAX_REQUESTS) || 100,
    message: 'Too many requests from this IP, please try again later.'
});
app.use('/analyze', limiter);



// Request queue
const queue = expressQueue({
    activeLimit: parseInt(process.env.MAX_CONCURRENT_USERS) || 20,
    queuedLimit: -1,
    timeout: parseInt(process.env.QUEUE_TIMEOUT_MS) || 300000 // 5 minutes
});
app.use('/analyze', queue);

// Error handling middleware
app.use((err, req, res, next) => {
    console.error(err);
    res.status(500).json({ error: 'Internal server error' });
});

// Routes
app.post('/analyze', upload.array('documents'), async (req, res) => {
    console.log("**RECEIVED REQUEST**", req.body);
    try {
        const sessionName = validateSessionName(req.body.sessionName);
        const projectName = req.body.projectName;
        
        if (!sessionName) {
            return res.status(400).json({ error: 'Session name is required' });
        }
        
        if (!projectName) {
            return res.status(400).json({ error: 'Project name is required' });
        }

        if (!req.files || req.files.length === 0) {
            console.error('No files were uploaded');
            return res.status(400).json({ error: 'No files uploaded' });
        }

        console.info(`Received ${req.files.length} files for processing`);
        
        // Log session directory structure
        const sessionDir = path.join('sessions', sessionName);
        console.info(`Using session directory: ${sessionDir}`);
        
        const results = await Promise.all(req.files.map(async (file) => {
            try {
                console.info(`Processing file: ${file.originalname} (${file.mimetype})`);
                
                // Extract images from the file
                console.info('Extracting images from file...');
                const fileExt = path.extname(file.originalname).toLowerCase();
                console.log("extracted file ext: ", fileExt);
                
                const baseFilename = path.basename(file.originalname, fileExt);
                let extractedImages = []; // Initialize as empty array
                let imageError = null;

                // Get session directory for this project
                const sessionDir = path.join('sessions', sessionName);
                const imagesDir = path.join(sessionDir, 'images');

                // Try to extract images, but continue even if it fails
                try {
                    if (fileExt === '.pdf') {
                        console.log("extracting images from pdf");
                        extractedImages = await extractImagesFromPDF(file.path, baseFilename, imagesDir);
                    } else {
                        console.log("extracting images from docx");
                        extractedImages = await extractImages(file.path, baseFilename, fileExt, imagesDir);
                    }
                    console.info(`Extracted ${extractedImages.length} images from ${file.originalname}`);
                } catch (error) {
                    console.error(`Error extracting images from ${file.originalname}:`, error);
                    imageError = error.message;
                }
                
                let fileContent;
                let filePath = file.path;

                // Convert Excel files to HTML
                if (file.mimetype === 'application/vnd.ms-excel' || 
                    file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
                    console.info('Converting Excel file to HTML...');
                    const tempDir = path.join(sessionDir, 'temp');
                    const { htmlContent, tempFilePath } = await excelToHTML(file.path, tempDir);
                    filePath = tempFilePath;
                    console.info('Excel file converted to HTML successfully');
                }

                // Read file as binary
                fileContent = await fs.promises.readFile(filePath);
                console.info(`Successfully read file: ${file.originalname}, size: ${fileContent.length} bytes`);
                
                // Upload file to OpenAI
                console.info('Uploading file to OpenAI...');
                const uploadedFile = await openai.files.create({
                    file: fs.createReadStream(filePath),
                    purpose: 'assistants'
                });
                console.info(`File uploaded successfully with ID: ${uploadedFile.id}`);

                // Create a new thread
                console.info('Creating new thread...');
                const thread = await openai.beta.threads.create();
                console.info(`Thread created with ID: ${thread.id}`);

                // Add the file content as a message
                console.info('Adding file content to thread...');
                try {
                    await openai.beta.threads.messages.create(thread.id, {
                        role: "user",
                        content: "Please analyze this document.",
                        attachments: [{
                            file_id: uploadedFile.id,
                            tools: [{ type: "file_search" }]
                        }]
                    });
                    console.info('File content added to thread');
                } catch (error) {
                    console.error('Error adding file content to thread:', {
                        error: error.message,
                        code: error.code,
                        type: error.type,
                        status: error.status,
                        response: error.response?.data,
                        threadId: thread.id,
                        fileId: uploadedFile.id
                    });
                    throw error;
                }

                // Run the assistant
                console.info('Starting assistant run...');
                console.info('Using assistant ID:', process.env.OPENAI_ASSISTANT_ID);
                console.info('Assistant ID length:', process.env.OPENAI_ASSISTANT_ID?.length);
                const run = await openai.beta.threads.runs.create(thread.id, {
                    assistant_id: process.env.OPENAI_ASSISTANT_ID
                });
                console.info(`Run started with ID: ${run.id}`);

                // Wait for the run to complete
                console.info('Waiting for run to complete...');
                let runStatus = await openai.beta.threads.runs.retrieve(thread.id, run.id);
                while (runStatus.status === 'in_progress' || runStatus.status === 'queued') {
                    console.info(`Run status: ${runStatus.status}`);
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    runStatus = await openai.beta.threads.runs.retrieve(thread.id, run.id);
                }
                console.info(`Run completed with status: ${runStatus.status}`);

                // Get the messages
                console.info('Retrieving messages from thread...');
                const messages = await openai.beta.threads.messages.list(thread.id);
                const assistantMessage = messages.data.find(m => m.role === 'assistant');
                console.info('Messages retrieved successfully');

                // Parse the analysis content and add image properties
                let analysisContent = assistantMessage.content[0].text.value;
                let jsonObjects = [];
                
                try {
                    // Try to parse as JSON array
                    jsonObjects = JSON.parse(analysisContent);
                    if (!Array.isArray(jsonObjects)) {
                        jsonObjects = [jsonObjects];
                    }
                } catch (e) {
                    // If not valid JSON, wrap in array
                    jsonObjects = [{ content: analysisContent }];
                }

                // Add image properties to each JSON object
                const maxImagesPerItem = 10;
                let currentImageIndex = 0;

                jsonObjects.forEach((jsonObj, index) => {
                    console.log("reemplaza imagenes");
                    console.log("extractedImages: ", extractedImages);
                    // Add image values to this object if there are remaining images
                    if (currentImageIndex < extractedImages.length) {
                        console.log("extractedImages: ", extractedImages);
                        const remainingImages = extractedImages.length - currentImageIndex;
                        const imagesToAdd = Math.min(remainingImages, maxImagesPerItem);
                        
                        for (let i = 0; i < imagesToAdd; i++) {
                            console.log("adding image to jsonObj: ", extractedImages);
                            jsonObj[`IMAGE ${i + 1}`] = extractedImages[currentImageIndex].path;
                            currentImageIndex++;
                        }
                    }
                });

                // Convert back to string
                analysisContent = JSON.stringify(jsonObjects, null, 2);

                // Calculate costs
                const promptTokens = runStatus.usage?.prompt_tokens || 0;
                const completionTokens = runStatus.usage?.completion_tokens || 0;
                const promptCost = (promptTokens / 1000000) * 2.5; // $2.5 per 1M tokens
                const completionCost = (completionTokens / 1000000) * 10; // $10 per 1M tokens
                const totalCost = promptCost + completionCost;

                // Store cleanup information
                const cleanupInfo = {
                    filePath: file.path,
                    htmlFilePath: filePath !== file.path ? filePath : null,
                    uploadedFileId: uploadedFile.id,
                    threadId: thread.id
                };

                return {
                    filename: file.originalname,
                    analysis: analysisContent,
                    extractedImages: extractedImages.map(img => ({
                        filename: img.filename,
                        path: img.path
                    })),
                    imageError: imageError,
                    cost: {
                        promptTokens,
                        completionTokens,
                        totalTokens: promptTokens + completionTokens,
                        promptCost: promptCost.toFixed(6),
                        completionCost: completionCost.toFixed(6),
                        totalCost: totalCost.toFixed(6)
                    },
                    cleanupInfo
                };
            } catch (error) {
                console.error(`Error processing file ${file.originalname}:`, error);
                console.error('Error details:', {
                    message: error.message,
                    stack: error.stack,
                    code: error.code
                });
                // Clean up the file even if there's an error
                if (fs.existsSync(file.path)) {
                    fs.unlinkSync(file.path);
                    console.info(`Cleaned up file after error: ${file.originalname}`);
                }
                return {
                    filename: file.originalname,
                    error: error.message
                };
            }
        }));

        console.info('All files processed successfully');
        
        // Debug: Check results array
        console.log('Results array:', results);
        console.log('Results length:', results.length);
        results.forEach((result, index) => {
            console.log(`Result ${index}:`, {
                filename: result.filename,
                hasAnalysis: !!result.analysis,
                analysisLength: result.analysis ? result.analysis.length : 0,
                error: result.error
            });
        });
        
        // Process results for Excel
        const jsonData = results.map(result => {
            try {
                console.log(`Parsing analysis for ${result.filename}:`, result.analysis);
                return JSON.parse(result.analysis);
            } catch (e) {
                console.error('Error parsing result:', e);
                console.error('Failed analysis content:', result.analysis);
                return null;
            }
        }).filter(data => data !== null);

        console.log('Final jsonData:', jsonData);
        console.log('jsonData length:', jsonData.length);

        // Write to Excel
        
        // Ensure all session directories exist before writing Excel
        const sessionDirs = [
            sessionDir,
            path.join(sessionDir, 'uploads'),
            path.join(sessionDir, 'images'),
            path.join(sessionDir, 'temp'),
            outputDir
        ];
        
        sessionDirs.forEach(dir => {
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir, { recursive: true });
                console.info(`Created session directory: ${dir}`);
            }
        });
        
        const excelPath = await writeToExcel(jsonData, projectName, outputDir);
        
        // Replace image filenames with actual embedded images
        try {
            await replaceImagesInExcel(excelPath, sessionName);
            console.log('Images successfully embedded in Excel file');
        } catch (error) {
            console.error('Error embedding images in Excel:', error);
            // Continue without images if there's an error
        }
        
        res.json({ 
            success: true,
            results,
            excelPath
        });

        // Perform cleanup after sending response
        results.forEach(result => {
            if (result.cleanupInfo) {
                try {
                    // Clean up local files
                    if (fs.existsSync(result.cleanupInfo.filePath)) {
                        fs.unlinkSync(result.cleanupInfo.filePath);
                        console.info(`Cleaned up local file: ${result.filename}`);
                    }
                    if (result.cleanupInfo.htmlFilePath && fs.existsSync(result.cleanupInfo.htmlFilePath)) {
                        fs.unlinkSync(result.cleanupInfo.htmlFilePath);
                        console.info(`Cleaned up HTML file for: ${result.filename}`);
                    }

                    // Clean up OpenAI resources
                    openai.files.del(result.cleanupInfo.uploadedFileId)
                        .then(() => console.info(`Deleted file from OpenAI: ${result.cleanupInfo.uploadedFileId}`))
                        .catch(err => console.error(`Error deleting OpenAI file: ${err.message}`));

                    openai.beta.threads.del(result.cleanupInfo.threadId)
                        .then(() => console.info(`Deleted thread: ${result.cleanupInfo.threadId}`))
                        .catch(err => console.error(`Error deleting thread: ${err.message}`));
                } catch (error) {
                    console.error(`Error during cleanup for ${result.filename}:`, error);
                }
            }
        });
    } catch (error) {
        console.error('Error processing documents:', error);
        console.error('Error details:', {
            message: error.message,
            stack: error.stack,
            code: error.code
        });
        res.status(500).json({ error: 'Error processing documents' });
    }
});

/* // Periodic session cleanup (every day)
setInterval(async () => {
    console.info('Running periodic session cleanup...');
    await cleanupOldSessions();
}, 86400000); // Run every day */

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'healthy' });
});

// Add download endpoint
app.get('/download', (req, res) => {
    const filePath = req.query.path;
    if (!filePath) {
        return res.status(400).json({ error: 'No file path provided' });
    }

    // Check if file exists
    if (!fs.existsSync(filePath)) {
        return res.status(404).json({ error: 'File not found' });
    }

    // Send the file
    res.download(filePath, (err) => {
        if (err) {
            console.error('Error downloading file:', err);
            res.status(500).json({ error: 'Error downloading file' });
        }
    });
});

// Helper function to replace image filenames with actual images in Excel
async function replaceImagesInExcel(excelPath, sessionName) {
    try {
        console.log(`Replacing images in Excel file: ${excelPath}`);
        const ExcelJS = require('exceljs');
        const workbook = new ExcelJS.Workbook();
        
        // Read the Excel file
        await workbook.xlsx.readFile(excelPath);
        const worksheet = workbook.getWorksheet(1);
        
        // Get the images directory for this session
        const imagesDir = path.join('sessions', sessionName, 'images');
        
        // Check if images directory exists
        if (!fs.existsSync(imagesDir)) {
            console.log(`Images directory not found: ${imagesDir}`);
            return;
        }
        
        // Get list of available images
        const availableImages = fs.readdirSync(imagesDir);
        console.log(`Available images in session: ${availableImages}`);
        
        // Function to find and replace image filenames with actual images
        async function replaceImageFilenamesWithImages() {
            let replacedCount = 0;
            
            worksheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    if (
                        cell.value &&
                        typeof cell.value === "string" &&
                        (cell.value.toLowerCase().endsWith(".jpeg") ||
                         cell.value.toLowerCase().endsWith(".jpg") ||
                         cell.value.toLowerCase().endsWith(".png") ||
                         cell.value.toLowerCase().endsWith(".gif") ||
                         cell.value.toLowerCase().endsWith(".bmp"))
                    ) {
                        console.log(`Found image filename in cell:  col: ${colNumber} row: ${rowNumber}    `);
                        
                        // Check if the image exists in the session's images directory
                        const imagePath =  cell.value
                        console.log(`Looking for image at: ${imagePath}`);
                        
                        if (fs.existsSync(imagePath)) {
                            try {
                                const imageBuffer = fs.readFileSync(imagePath);
                                
                                // Determine image extension from filename
                                const extension = path.extname(cell.value).toLowerCase().substring(1);
                                
                                // Add image to workbook
                                const imageId = workbook.addImage({
                                    buffer: imageBuffer,
                                    extension: extension,
                                });
                                
                                console.log(`Added image to workbook with ID: ${imageId}`);
                                
                                // Clear the cell value (remove filename)
                                //cell.value = "";
                                
                                // Add the image to the worksheet
                                worksheet.addImage(imageId, {
                                    tl: { col: colNumber - 1, row: rowNumber - 1 },
                                    ext: { width: 100, height: 100 },
                                });
                                
                                // Adjust row height and column width to accommodate image
                                row.height = 100;
                                worksheet.getColumn(colNumber).width = 30;
                                
                                replacedCount++;
                                console.log(`Successfully replaced image: ${cell.value}`);
                            } catch (error) {
                                console.error(`Error processing image ${cell.value}:`, error);
                            }
                        } else {
                            console.log(`Image not found: ${imagePath}`);
                        }
                    }
                });
            });
            
            console.log(`Total images replaced: ${replacedCount}`);
        }
        
        // Replace image filenames with actual images
        await replaceImageFilenamesWithImages();
        
        // Save the updated workbook
        await workbook.xlsx.writeFile(excelPath);
        console.log(`Excel file updated with embedded images: ${excelPath}`);
        
    } catch (error) {
        console.error('Error replacing images in Excel:', error);
        throw error;
    }
}

// Start server
app.listen(port, () => {
    console.info(`Server running at http://localhost:${port} time: ${new Date().toISOString()}  `);
}); 