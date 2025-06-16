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
const NodeCache = require('node-cache');
const { extractImages } = require('./imageExtractor');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const xlsx = require('xlsx');
const { extractImagesFromPDF } = require('./pdfImageExtractor');

// Validate environment variables
const requiredEnvVars = ['OPENAI_API_KEY', 'OPENAI_ASSISTANT_ID'];
const missingEnvVars = requiredEnvVars.filter(envVar => !process.env[envVar]);
if (missingEnvVars.length > 0) {
    console.error(`Missing required environment variables: ${missingEnvVars.join(', ')}`);
    process.exit(1);
}

const app = express();
const port = process.env.PORT || 3000;

// Initialize OpenAI
const openai = new OpenAI({
    apiKey: process.env.OPENAI_API_KEY
});

// Initialize cache for storing thread IDs and file IDs
const threadCache = new NodeCache({ stdTTL: 3600 }); // 1 hour TTL
const fileCache = new NodeCache({ stdTTL: 3600 }); // 1 hour TTL

// Create uploads directory if it doesn't exist
const uploadDir = 'uploads';
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
    console.info('Created uploads directory');
}

// Create images directory if it doesn't exist
const imagesDir = 'images';
if (!fs.existsSync(imagesDir)) {
    fs.mkdirSync(imagesDir);
    console.info('Created images directory');
}

// Configure multer for file upload
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
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
    try {
        if (!req.files || req.files.length === 0) {
            console.error('No files were uploaded');
            return res.status(400).json({ error: 'No files uploaded' });
        }

        console.info(`Received ${req.files.length} files for processing`);
        
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

                // Try to extract images, but continue even if it fails
                try {
                    if (fileExt === '.pdf') {
                        extractedImages = await extractImagesFromPDF(file.path, baseFilename);
                    } else {
                        extractedImages = await extractImages(file.path, baseFilename, fileExt);
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
                    const { htmlContent, tempFilePath } = await excelToHTML(file.path);
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
                    analysis: assistantMessage.content[0].text.value,
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
        res.json({ results });

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

// Comment out the hourly cleanup job
/* setInterval(async () => {
    try {
        // List all files
        const files = await openai.files.list();
        
        // Delete files older than 1 hour
        const oneHourAgo = Date.now() - 3600000;
        for (const file of files.data) {
            if (file.created_at * 1000 < oneHourAgo) {
                await openai.files.del(file.id);
                console.info(`Cleaned up old file: ${file.id}`);
            }
        }
    } catch (error) {
        console.error('Error in cleanup job:', error);
    }
}, 3600000); // Run every hour */

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'healthy' });
});

// Start server
app.listen(port, () => {
    console.info(`Server running at http://localhost:${port}`);
}); 