require('dotenv').config();
const express = require('express');
const multer = require('multer');
const cors = require('cors');
const OpenAI = require('openai');
const path = require('path');
const fs = require('fs');
const rateLimit = require('express-rate-limit');
const helmet = require('helmet');
const compression = require('compression');
const expressQueue = require('express-queue');
const NodeCache = require('node-cache');
const logger = require('./logger');
const { excelToHTML } = require('./excelToHTML');
const { extractImages } = require('./imageExtractor');
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');

// Validate environment variables
const requiredEnvVars = ['OPENAI_API_KEY', 'OPENAI_ASSISTANT_ID'];
const missingEnvVars = requiredEnvVars.filter(envVar => !process.env[envVar]);
if (missingEnvVars.length > 0) {
    logger.error(`Missing required environment variables: ${missingEnvVars.join(', ')}`);
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
    logger.info('Created uploads directory');
}

// Create images directory if it doesn't exist
const imagesDir = 'images';
if (!fs.existsSync(imagesDir)) {
    fs.mkdirSync(imagesDir);
    logger.info('Created images directory');
}

// Configure multer for file upload
const storage = multer.diskStorage({
    destination: function (req, file, cb) {
        cb(null, uploadDir);
    },
    filename: function (req, file, cb) {
        cb(null, Date.now() + '-' + file.originalname);
    }
});

const upload = multer({
    storage: storage,
    limits: {
        fileSize: parseInt(process.env.MAX_FILE_SIZE) || 10485760, // 10MB default
    },
    fileFilter: (req, file, cb) => {
        const allowedTypes = [
            'application/pdf',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'text/plain',
            'text/markdown',
            'application/json',
            'text/csv'
        ];
        
        if (allowedTypes.includes(file.mimetype)) {
            cb(null, true);
        } else {
            cb(new Error('Invalid file type'));
        }
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
    logger.error(err);
    res.status(500).json({ error: 'Internal server error' });
});

// Routes
app.post('/analyze', upload.array('documents'), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            logger.error('No files were uploaded');
            return res.status(400).json({ error: 'No files uploaded' });
        }

        logger.info(`Received ${req.files.length} files for processing`);
        
        const results = await Promise.all(req.files.map(async (file) => {
            try {
                logger.info(`Processing file: ${file.originalname} (${file.mimetype})`);
                
                // Extract images from the file
                logger.info('Extracting images from file...');
                const extractedImages = await extractImages(file.path, file.originalname);
                logger.info(`Extracted ${extractedImages.length} images from ${file.originalname}`);

                let fileContent;
                let filePath = file.path;

                // Convert Excel files to HTML
                if (file.mimetype === 'application/vnd.ms-excel' || 
                    file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') {
                    logger.info('Converting Excel file to HTML...');
                    const { htmlContent, tempFilePath } = await excelToHTML(file.path);
                    filePath = tempFilePath;
                    logger.info('Excel file converted to HTML successfully');
                }

                // Read file as binary
                fileContent = await fs.promises.readFile(filePath);
                logger.info(`Successfully read file: ${file.originalname}, size: ${fileContent.length} bytes`);
                
                // Upload file to OpenAI
                logger.info('Uploading file to OpenAI...');
                const uploadedFile = await openai.files.create({
                    file: fs.createReadStream(filePath),
                    purpose: 'assistants'
                });
                logger.info(`File uploaded successfully with ID: ${uploadedFile.id}`);

                // Create a new thread
                logger.info('Creating new thread...');
                const thread = await openai.beta.threads.create();
                logger.info(`Thread created with ID: ${thread.id}`);

                // Add the file content as a message
                logger.info('Adding file content to thread...');
                try {
                    await openai.beta.threads.messages.create(thread.id, {
                        role: "user",
                        content: "Please analyze this document.",
                        attachments: [{
                            file_id: uploadedFile.id,
                            tools: [{ type: "file_search" }]
                        }]
                    });
                    logger.info('File content added to thread');
                } catch (error) {
                    logger.error('Error adding file content to thread:', {
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
                logger.info('Starting assistant run...');
                const run = await openai.beta.threads.runs.create(thread.id, {
                    assistant_id: process.env.OPENAI_ASSISTANT_ID
                });
                logger.info(`Run started with ID: ${run.id}`);

                // Wait for the run to complete
                logger.info('Waiting for run to complete...');
                let runStatus = await openai.beta.threads.runs.retrieve(thread.id, run.id);
                while (runStatus.status === 'in_progress' || runStatus.status === 'queued') {
                    logger.info(`Run status: ${runStatus.status}`);
                    await new Promise(resolve => setTimeout(resolve, 1000));
                    runStatus = await openai.beta.threads.runs.retrieve(thread.id, run.id);
                }
                logger.info(`Run completed with status: ${runStatus.status}`);

                // Get the messages
                logger.info('Retrieving messages from thread...');
                const messages = await openai.beta.threads.messages.list(thread.id);
                const assistantMessage = messages.data.find(m => m.role === 'assistant');
                logger.info('Messages retrieved successfully');

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
                logger.error(`Error processing file ${file.originalname}:`, error);
                logger.error('Error details:', {
                    message: error.message,
                    stack: error.stack,
                    code: error.code
                });
                // Clean up the file even if there's an error
                if (fs.existsSync(file.path)) {
                    fs.unlinkSync(file.path);
                    logger.info(`Cleaned up file after error: ${file.originalname}`);
                }
                return {
                    filename: file.originalname,
                    error: error.message
                };
            }
        }));

        logger.info('All files processed successfully');
        res.json({ results });

        // Perform cleanup after sending response
        results.forEach(result => {
            if (result.cleanupInfo) {
                try {
                    // Clean up local files
                    if (fs.existsSync(result.cleanupInfo.filePath)) {
                        fs.unlinkSync(result.cleanupInfo.filePath);
                        logger.info(`Cleaned up local file: ${result.filename}`);
                    }
                    if (result.cleanupInfo.htmlFilePath && fs.existsSync(result.cleanupInfo.htmlFilePath)) {
                        fs.unlinkSync(result.cleanupInfo.htmlFilePath);
                        logger.info(`Cleaned up HTML file for: ${result.filename}`);
                    }

                    // Clean up OpenAI resources
                    openai.files.del(result.cleanupInfo.uploadedFileId)
                        .then(() => logger.info(`Deleted file from OpenAI: ${result.cleanupInfo.uploadedFileId}`))
                        .catch(err => logger.error(`Error deleting OpenAI file: ${err.message}`));

                    openai.beta.threads.del(result.cleanupInfo.threadId)
                        .then(() => logger.info(`Deleted thread: ${result.cleanupInfo.threadId}`))
                        .catch(err => logger.error(`Error deleting thread: ${err.message}`));
                } catch (error) {
                    logger.error(`Error during cleanup for ${result.filename}:`, error);
                }
            }
        });
    } catch (error) {
        logger.error('Error processing documents:', error);
        logger.error('Error details:', {
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
                logger.info(`Cleaned up old file: ${file.id}`);
            }
        }
    } catch (error) {
        logger.error('Error in cleanup job:', error);
    }
}, 3600000); // Run every hour */

// Health check endpoint
app.get('/health', (req, res) => {
    res.json({ status: 'healthy' });
});

// Start server
app.listen(port, () => {
    logger.info(`Server running at http://localhost:${port}`);
}); 