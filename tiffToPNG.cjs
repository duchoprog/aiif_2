/* // tiffToJPG.cjs
const sharp = require('sharp');
const fs = require('fs').promises;

async function convertTiffToPng(sourcePath, outputPath) {
  try {
    console.log('Reading and converting with sharp...');
    
  
    await sharp(sourcePath)
      .jpeg() 
      .toFile(outputPath); // Save to the output file

    console.log('\n✅ Success! Image converted successfully with sharp.');
    return outputPath;
  } catch (error) {
    console.error('\n❌ An error occurred during the sharp conversion:', error);
    throw error;
  }
}

module.exports = {
  convertTiffToPng
};
 */