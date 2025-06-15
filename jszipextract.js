// extractWithJSZip.js

const JSZip = require("jszip");
const fs = require("fs").promises;
const path = require("path");

// --- CONFIGURATION ---
const docxPath = path.join(__dirname, "Lily-img.docx"); // Your .docx file
const outputDir = path.join(__dirname, "extracted_images_jszip"); // Output folder
// --- END CONFIGURATION ---

async function extractImagesFromDocx() {
  console.log(`Reading file: ${docxPath}`);

  try {
    // 1. Read the .docx file into a buffer.
    const data = await fs.readFile(docxPath);

    // 2. Load the buffer as a zip file.
    const zip = await JSZip.loadAsync(data);

    // 3. Filter the files to only include those in the "word/media/" folder.
    const mediaFolder = zip.folder("word/media");
    if (!mediaFolder) {
      console.log("No 'word/media' folder found in the document. No images to extract.");
      return;
    }

    // Ensure the output directory exists.
    await fs.mkdir(outputDir, { recursive: true });

    let imageCounter = 0;
    const promises = [];

    // 4. Loop through each file in the "word/media" folder.
    mediaFolder.forEach((relativePath, file) => {
      // Ignore directories
      if (file.dir) {
        return;
      }

      imageCounter++;
      const outputPath = path.join(outputDir, relativePath);
      console.log(`Found image: ${relativePath}. Saving to ${outputPath}`);

      // 5. Get the image data as a Node.js Buffer and create a promise to write the file.
      const writePromise = file.async("nodebuffer").then((content) => {
        return fs.writeFile(outputPath, content);
      });
      promises.push(writePromise);
    });

    // Wait for all file-writing promises to complete.
    await Promise.all(promises);

    if (imageCounter > 0) {
      console.log(`\n✅ Success! Extracted ${imageCounter} image(s) to '${path.basename(outputDir)}'.`);
    } else {
      console.log(`\nℹ️  Process finished, but no images were found in the 'word/media' folder.`);
    }

  } catch (error) {
    console.error("\n❌ An error occurred:", error);
  }
}

// Run the extraction function
extractImagesFromDocx();