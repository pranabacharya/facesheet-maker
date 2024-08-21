"use strict";

const libre = require("libreoffice-convert");
const process = require("process");
const PdfConverter = require("./controllers/PdfConverter.js");
const generateDocx = require("./controllers/FacesheetGenerator.js");
const path = require("path");

libre.convertAsync = require("util").promisify(libre.convert);

let commandPath = null;
let faceSheetData = null;

process.argv.forEach((arg) => {
  const [key, value] = arg.split("=");
  switch (key) {
    case "--path":
      commandPath = value;
      break;
    case "--json":
      faceSheetData = value;
      break;
    default:
      break;
  }
});

if (!commandPath || !faceSheetData) {
  console.error("--path and --json required");
  return;
} else {
  PdfConverter(commandPath)
    .then((res) => {
      console.log(res);
      // Parameters
      // eslint-disable-next-line no-undef
      const imagePath = path.join(__dirname, "demo_cv", "tcs_logo.png"); // Replace with your image path
      if (faceSheetData) {
        generateDocx(imagePath, faceSheetData);
      }
    })
    .catch((err) => console.log(err));
}
