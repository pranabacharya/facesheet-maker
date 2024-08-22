"use strict";

const process = require("process");
const generateDocx = require("./controllers/FacesheetGenerator.js");
const path = require("path");
const fs = require('fs');
const fsPromise = require('fs').promises;
const { randomUUID } = require("crypto");
const archiver = require('archiver');

let jsonFile = null;

process.argv.forEach((arg) => {
  const [key, value] = arg.split("=");
  switch (key) {
    case "--json":
      jsonFile = value;
      break;
    default:
      break;
  }
});

if (!jsonFile) {
  console.error("--json required");
  return;
} else {
  (async () => {
    
    const jsonString = await fsPromise.readFile(jsonFile, {encoding: 'utf-8'});
    const jsonObj = JSON.parse(jsonString);
    const imagePath = jsonObj.image_path;
    const tenantName = jsonObj.tenant_name;
    const borderColor = jsonObj.border_color;
    const candidates = jsonObj.candidates;
    const zipFolderName = randomUUID();

    for( let i = 0; i < candidates.length ; i++ ) {
      await generateDocx(imagePath,borderColor, tenantName, candidates[i], zipFolderName);
    }

    const zipFolder = path.join(__dirname, `${zipFolderName}.zip`);
    const dataFolder = path.join(__dirname, zipFolderName);
    await zipFolderSync(dataFolder, zipFolder);
    await fsPromise.rm(dataFolder, {force: true, recursive: true});
    console.log(zipFolder);
  })();
}

function zipFolderSync(folderPath, outputZipPath) {
  return new Promise((resolve, reject) => {
      const output = fs.createWriteStream(outputZipPath);
      const archive = archiver('zip', { zlib: { level: 9 } });

      output.on('close', () => {
          resolve();
      });

      archive.on('error', (err) => {
          reject(err);
      });

      archive.pipe(output);
      archive.directory(folderPath, path.basename(folderPath));
      archive.finalize();
  });
}
