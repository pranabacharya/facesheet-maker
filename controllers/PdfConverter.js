const path = require("path");
const libre = require("libreoffice-convert");
const fs = require("fs").promises;
const { randomUUID } = require("crypto");
libre.convertAsync = require("util").promisify(libre.convert);

module.exports = async (argPath) => {
  const ext = ".pdf";
  const inputPath = argPath;
  const tempFileName = randomUUID();
  // Define the base output directory path
  const basePath = path.join(__dirname, "../output");

  try {
    await fs.access(basePath);
  } catch (error) {
    if (error.code === "ENOENT") {
      await fs.mkdir(basePath, { recursive: true });
    } else {
      throw error; // rethrow other errors
    }
  }

  // Define the output path with a unique filename
  const outputPath = path.join(`${basePath}`, `${tempFileName}${ext}`);

  // Read the input DOCX file
  const docxBuf = await fs.readFile(inputPath);

  // Convert the DOCX file to PDF
  let pdfBuffer = await libre.convertAsync(docxBuf, ext, undefined);

  // Write the converted PDF to the output path
  await fs.writeFile(outputPath, pdfBuffer);

  return outputPath;
};
