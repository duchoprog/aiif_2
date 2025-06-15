const { extractImages } = require('./imageExtractor');

async function testExtraction() {
    try {
       /*  // Test with a PDF file
        console.log('Testing PDF extraction...');
        const pdfResults = await extractImages(
            'path/to/your/test.pdf',
            'test-pdf',
            '.pdf'
        );
        console.log('PDF Results:', pdfResults); */

        // Test with a DOCX file
        console.log('\nTesting DOCX extraction...');
        const docxResults = await extractImages(
            "./Lily-img.docx",
            "Lily-img",  // baseFilename
            ".docx"      // fileExt
        );
        console.log('DOCX Results:', docxResults);

        // Test with an XLSX file
        console.log('\nTesting XLSX extraction...');
        const xlsxResults = await extractImages(
            './Jimmy.xlsx',
            'Jimmy',
            '.xlsx'
        );
        console.log('XLSX Results:', xlsxResults);

    } catch (error) {
        console.error('Test failed:', error);
    }
}

// Run the test
testExtraction(); 