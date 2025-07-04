const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const selectFilesBtn = document.getElementById('selectFilesBtn');
const fileList = document.getElementById('fileList');
const analyzeBtn = document.getElementById('analyzeBtn');
const loading = document.getElementById('loading');
const results = document.getElementById('results');
const xlsWarningPopup = document.getElementById('xlsWarningPopup');
const closePopupBtn = document.getElementById('closePopupBtn');
let currentFiles = new Map();

// Generate session name based on project name
function generateSessionName(projectName) {
    console.log("generating session name for project: ", projectName);
    if (projectName && projectName.trim()) {
        // Remove unsafe characters and trim
        const safeName = projectName.trim().replace(/[^a-zA-Z0-9-_]/g, '');
        // Generate a random 4-digit number
        const random4 = Math.floor(1000 + Math.random() * 9000);
        console.log("generated session name: ", `${safeName}-${random4}`);
        return `${safeName}${random4}`;
    } else {
        // Generate a random 8-digit number
        const random8 = Math.floor(10000000 + Math.random() * 90000000);
        console.log("generated session name: ", `Session${random8}`);
        return `Session${random8}`;
    }
}

// Close popup when clicking the close button
closePopupBtn.addEventListener('click', () => {
    xlsWarningPopup.style.display = 'none';
});

// Close popup when clicking outside
xlsWarningPopup.addEventListener('click', (e) => {
    if (e.target === xlsWarningPopup) {
        xlsWarningPopup.style.display = 'none';
    }
});

// Prevent default drag behaviors
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, preventDefaults, false);
    document.body.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

// Highlight drop zone when dragging over it
['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, highlight, false);
});

['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, unhighlight, false);
});

function highlight(e) {
    dropZone.classList.add('dragover');
}

function unhighlight(e) {
    dropZone.classList.remove('dragover');
}

// Handle dropped files
dropZone.addEventListener('drop', handleDrop, false);
fileInput.addEventListener('change', handleFileSelect, false);
selectFilesBtn.addEventListener('click', () => fileInput.click());

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    handleFiles(files);
}

function handleFileSelect(e) {
    const files = e.target.files;
    handleFiles(files);
}

function handleFiles(files) {
    let hasXlsFile = false;
    Array.from(files).forEach(file => {
        currentFiles.set(file.name, file);
        // Check if any of the files is an XLS file
        if (file.name.toLowerCase().endsWith('.xls')) {
            hasXlsFile = true;
        }
    });
    updateFileList();
    analyzeBtn.disabled = currentFiles.size === 0;
    
    // Show warning popup if there are XLS files
    if (hasXlsFile) {
        xlsWarningPopup.style.display = 'flex';
    }
}

function updateFileList() {
    fileList.innerHTML = '';
    currentFiles.forEach((file, name) => {
        const fileItem = document.createElement('div');
        fileItem.className = 'file-item';
        
        const fileNameSpan = document.createElement('span');
        fileNameSpan.textContent = name;
        
        const removeButton = document.createElement('button');
        removeButton.textContent = '×';
        removeButton.addEventListener('click', () => removeFile(name));
        
        fileItem.appendChild(fileNameSpan);
        fileItem.appendChild(removeButton);
        fileList.appendChild(fileItem);
    });
}

function removeFile(filename) {
    currentFiles.delete(filename);
    updateFileList();
    analyzeBtn.disabled = currentFiles.size === 0;
}

analyzeBtn.addEventListener('click', async () => {
    if (currentFiles.size === 0) return;

    const projectName = document.getElementById('projectName').value.trim();
    if (!projectName) {
        alert('Please enter a project name');
        return;
    }

    const sessionName = generateSessionName(projectName);

    const formData = new FormData();
    currentFiles.forEach(file => {
        console.log("adding file to form data: ", file);
        formData.append('documents', file);
    });
    formData.append('projectName', projectName);
    formData.append('sessionName', sessionName);

    loading.style.display = 'block';
    results.innerHTML = '';
    analyzeBtn.disabled = true;

    try {
        const response = await fetch(`/analyze?sessionName=${encodeURIComponent(sessionName)}`, {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        
        if (response.ok) {
            results.innerHTML = data.results.map(result => `
                <div class="result-item">
                    <h3>${result.filename}</h3>
                    ${result.error 
                        ? `<div class="error">Error: ${result.error}</div>`
                        : `<div class="result-content">
                            ${result.analysis}
                            ${result.cost ? `
                                <div class="cost-info">
                                    <h4>Cost Information:</h4>
                                    <p>Total Tokens: ${result.cost.totalTokens}</p>
                                    <p>Prompt Tokens: ${result.cost.promptTokens} ($${result.cost.promptCost})</p>
                                    <p>Completion Tokens: ${result.cost.completionTokens} ($${result.cost.completionCost})</p>
                                    <p>Total Cost: $${result.cost.totalCost}</p>
                                </div>
                            ` : ''}
                        </div>`
                    }
                </div>
            `).join('');

            // Add download link if Excel was generated
            if (data.excelPath) {
                results.innerHTML += `
                    <div class="excel-download">
                        <h3>Excel File Generated</h3>
                        <a href="/download?path=${encodeURIComponent(data.excelPath)}" class="download-btn">
                            Download Excel File
                        </a>
                    </div>
                `;
            }
        } else {
            results.innerHTML = `<div class="error">Error: ${data.error}</div>`;
        }
    } catch (error) {
        results.innerHTML = '<div class="error">Error: Failed to analyze documents</div>';
    } finally {
        loading.style.display = 'none';
        analyzeBtn.disabled = false;
        currentFiles.clear();
        updateFileList();
    }
}); 