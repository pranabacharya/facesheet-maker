const fs = require("fs");
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  WidthType,
  AlignmentType,
  BorderStyle,
  ImageRun,
} = require("docx");

// Function to generate DOCX
async function generateDocx(imagePath, jsonData) {
  jsonData = JSON.parse(jsonData);
  const doc = new Document({
    styles: {
      default: {
        font: {
          name: "Calibri", // Change the font here
          size: 18, // Change the font size here
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            size: { width: 11906, height: 16838 }, // A4 size in twips (1/20th of a point)
            margin: {
              top: 720,
              right: 1024,
              bottom: 720,
              left: 1024,
            },
          },
        },
        children: [
          // Image
          new Paragraph({
            children: [
              new ImageRun({
                data: fs.readFileSync(imagePath),
                transformation: { width: 100, height: 100 },
              }),
            ],
            alignment: AlignmentType.RIGHT,
            spacing: { after: 100 }, // space after the image
            rightTabStop: 800, // Adding some padding by using a tab stop
          }),

          // Table
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            rows: Object.entries(jsonData).map(
              ([key, value]) =>
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: key,
                              size: 28,
                              font: "Calibri",
                            }),
                          ],
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        bottom: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        left: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        right: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                      },
                      width: { size: 50, type: WidthType.PERCENTAGE },
                      alignment: AlignmentType.LEFT,
                      verticalAlign: AlignmentType.CENTER, // Center the text vertically
                      margins: {
                        top: 60,
                        bottom: 60,
                        left: 100,
                        right: 100,
                      },
                      padding: {
                        top: 200,
                        bottom: 200,
                        left: 200, // Add padding here
                        right: 200, // Add padding here
                      },
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: value,
                              size: 28,
                              font: "Calibri",
                              //bold: true,
                            }),
                          ],
                        }),
                      ],
                      borders: {
                        top: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        bottom: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        left: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                        right: {
                          style: BorderStyle.SINGLE,
                          size: 6,
                          color: "000088",
                        },
                      },
                      width: { size: 50, type: WidthType.PERCENTAGE },
                      alignment: AlignmentType.LEFT,
                      verticalAlign: AlignmentType.CENTER, // Center the text vertically
                      margins: {
                        top: 60,
                        bottom: 60,
                        left: 100,
                        right: 100,
                      },
                      padding: {
                        top: 200,
                        bottom: 200,
                        left: 200, // Add padding here
                        right: 200, // Add padding here
                      },
                    }),
                  ],
                })
            ),
            alignment: AlignmentType.CENTER,
          }),

          // Note
          new Paragraph({
            children: [
              new TextRun({
                text: "Note : This data represents the system-generated candidate information used for candidate evaluation.",
                size: 12,
                color: "808080",
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { before: 400, after: 200 }, // space before and after the note
          }),

          // Footer
          new Paragraph({
            children: [
              new TextRun({
                text: "Powered by Corporate Resources",
                size: 12,
                bold: true,
                color: "808080",
              }),
            ],
            alignment: AlignmentType.CENTER,
          }),
        ],
      },
    ],
  });

  // Create and save the document
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync("output11.docx", buffer);
  console.log("Document created successfully");
}

// Generate DOCX
module.exports = generateDocx;
