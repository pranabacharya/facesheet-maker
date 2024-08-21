const path = require("path");
const libre = require("libreoffice-convert");
const fs = require("fs").promises;
const { randomUUID } = require("crypto");

module.exports = async (argPath) => {
  const ext = ".pdf";
  const inputPath = argPath;
  const tempFileName = randomUUID();
  // eslint-disable-next-line no-undef
  const basePath = path.join(__dirname, "../output");
  if (!basePath) {
    fs.mkdir(basePath, { recursive: true });
  }
  const outputPath = path.join(`${basePath}`, `${tempFileName}${ext}`);

  const docxBuf = await fs.readFile(inputPath);

  let pdfBuffer = await libre.convertAsync(docxBuf, ext, undefined);

  await fs.writeFile(outputPath, pdfBuffer);
  return outputPath;
};
