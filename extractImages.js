const fs = require('fs');
const path = require('path');
const { extractImages } = require('./imageExtractor');
const logger = require('./logger');

// Create uploads directory if it doesn't exist
const uploadDir = 'uploads';
if (!fs.existsSync(uploadDir)) {
    fs.mkdirSync(uploadDir);
    logger.info('Created uploads directory');
}

// Configure multer for file upload
const multer = require('multer');
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
        fileSize: 10485760, // 10MB
    },
    fileFilter: (req, file, cb) => {
        console.log('Multer fileFilter called with file:', file);
        const allowedTypes = [
            'application/pdf',
            'application/msword',
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        ];
        
        if (allowedTypes.includes(file.mimetype)) {
            console.log('File type accepted:', file.mimetype);
            cb(null, true);
        } else {
            console.log('File type rejected:', file.mimetype);
            cb(new Error('Invalid file type'));
        }
    }
});

const express = require('express');
const app = express();
const port = 3001; // Different port from main server

app.use(express.json());
app.use(express.static('public'));
// Add static file serving for the images directory
app.use('/images', express.static(path.join(__dirname, 'images')));

// Add error handling middleware
app.use((err, req, res, next) => {
    console.error('Error in middleware:', err);
    if (err instanceof multer.MulterError) {
        console.error('Multer error details:', {
            code: err.code,
            field: err.field,
            message: err.message
        });
    }
    next(err);
});

// Route to handle file upload and image extraction
app.post('/extract', upload.any(), async (req, res) => {
    try {
        console.log('Received file upload request');
        console.log('Request files:', req.files);
        
        if (!req.files || req.files.length === 0) {
            console.log('No files in request');
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const file = req.files[0];
        console.log(`Processing file: ${file.originalname} (${file.mimetype})`);
        console.log(`File path: ${file.path}`);
        
        // Extract images
        console.log('Starting image extraction...');
        const extractedImages = await extractImages(file.path, file.originalname);
        console.log(`Extracted ${extractedImages.length} images from ${file.originalname}`);

        // Clean up the uploaded file
        try {
            fs.unlinkSync(file.path);
            console.log(`Cleaned up uploaded file: ${file.path}`);
        } catch (cleanupError) {
            console.error(`Error cleaning up uploaded file: ${cleanupError.message}`);
        }

        res.json({
            filename: file.originalname,
            extractedImages: extractedImages.map(img => ({
                filename: img.filename,
                path: img.path
            }))
        });

    } catch (error) {
        console.error('Error processing file:', error);
        console.error('Error stack:', error.stack);
        // Clean up the uploaded file even if there's an error
        if (req.files && req.files.length > 0) {
            try {
                fs.unlinkSync(req.files[0].path);
                console.log(`Cleaned up uploaded file after error: ${req.files[0].path}`);
            } catch (cleanupError) {
                console.error(`Error cleaning up uploaded file after error: ${cleanupError.message}`);
            }
        }
        res.status(500).json({ 
            error: error.message,
            stack: error.stack
        });
    }
});

// Serve the test page
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html>
        <head>
            <title>Image Extraction Test</title>
            <style>
                body { font-family: Arial, sans-serif; max-width: 800px; margin: 0 auto; padding: 20px; }
                .result { margin-top: 20px; }
                .image-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 20px; }
                .image-item { border: 1px solid #ccc; padding: 10px; }
                .image-item img { max-width: 100%; height: auto; }
                .error { color: red; background: #ffeeee; padding: 10px; border: 1px solid red; margin: 10px 0; }
                .success { color: green; background: #eeffee; padding: 10px; border: 1px solid green; margin: 10px 0; }
                .loading { color: blue; }
            </style>
        </head>
        <body>
            <h1>Image Extraction Test Server</h1>
            <p>This is a test interface running on port 3001. Use this page to test image extraction without affecting the main application.</p>
            <form id="uploadForm" enctype="multipart/form-data">
                <input type="file" name="document" accept=".pdf,.doc,.docx,.xls,.xlsx" required>
                <button type="submit">Extract Images</button>
            </form>
            <div id="result" class="result"></div>

            <script>
                document.getElementById('uploadForm').onsubmit = async (e) => {
                    e.preventDefault();
                    const formData = new FormData(e.target);
                    console.log('Form data entries:');
                    for (let pair of formData.entries()) {
                        console.log(pair[0], pair[1]);
                    }
                    const resultDiv = document.getElementById('result');
                    resultDiv.innerHTML = '<p class="loading">Processing... Please wait.</p>';

                    try {
                        console.log('Sending request to test server...');
                        const response = await fetch('/extract', {
                            method: 'POST',
                            body: formData
                        });
                        console.log('Response status:', response.status);
                        const data = await response.json();
                        console.log('Response data:', data);
                        
                        if (data.error) {
                            resultDiv.innerHTML = \`
                                <div class="error">
                                    <h3>Error:</h3>
                                    <p>\${data.error}</p>
                                    \${data.stack ? \`<pre>\${data.stack}</pre>\` : ''}
                                </div>
                            \`;
                            return;
                        }

                        let html = \`
                            <div class="success">
                                <h2>Successfully extracted \${data.extractedImages.length} images from \${data.filename}</h2>
                            </div>
                        \`;
                        
                        if (data.extractedImages.length > 0) {
                            html += '<div class="image-grid">';
                            data.extractedImages.forEach(img => {
                                html += \`
                                    <div class="image-item">
                                        <img src="/images/\${img.path}" alt="\${img.filename}" onerror="this.onerror=null; this.src='data:image/svg+xml,<svg xmlns=\\'http://www.w3.org/2000/svg\\' width=\\'200\\' height=\\'200\\'><text x=\\'50%\\' y=\\'50%\\' dominant-baseline=\\'middle\\' text-anchor=\\'middle\\'>Image not found</text></svg>'">
                                        <p>\${img.filename}</p>
                                    </div>
                                \`;
                            });
                            html += '</div>';
                        } else {
                            html += '<p>No images found in the document.</p>';
                        }
                        resultDiv.innerHTML = html;
                    } catch (error) {
                        resultDiv.innerHTML = \`
                            <div class="error">
                                <h3>Error:</h3>
                                <p>\${error.message}</p>
                            </div>
                        \`;
                    }
                };
            </script>
        </body>
        </html>
    `);
});

app.listen(port, () => {
    logger.info(`Image extraction test server running on port ${port}`);
}); 