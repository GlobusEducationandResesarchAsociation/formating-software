const express = require("express");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const { PDFDocument, StandardFonts } = require("pdf-lib");
const { Document, Packer, Paragraph, TextRun } = require("docx");

const app = express();
const upload = multer({ dest: "uploads/" });
app.use(express.static("public"));

async function processDocx(filePath, outputPath) {
  const text = fs.readFileSync(filePath);
  const doc = new Document({
    sections: [{
      properties: {
        page: { size: { width: 11907, height: 16840 } },
      },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: text.toString(),
              font: "Times New Roman",
              size: 24,
            }),
          ],
          spacing: { line: 360 },
        }),
      ],
    }],
  });
  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outputPath, buffer);
}

async function processPdf(filePath, outputPath) {
  const existingPdf = await PDFDocument.load(fs.readFileSync(filePath));
  const pages = existingPdf.getPages();
  const newPdf = await PDFDocument.create();
  const font = await newPdf.embedFont(StandardFonts.TimesRoman);

  for (const page of pages) {
    const { width, height } = page.getSize();
    const newPage = newPdf.addPage([595.28, 841.89]);
    newPage.drawText("Converted to Times New Roman, 12pt, 1.5 spacing", {
      x: 50,
      y: height - 80,
      size: 12,
      font,
      lineHeight: 18,
    });
  }

  const pdfBytes = await newPdf.save();
  fs.writeFileSync(outputPath, pdfBytes);
}

app.post("/upload", upload.single("file"), async (req, res) => {
  const file = req.file;
  if (!file) return res.status(400).send("No file uploaded.");

  const ext = path.extname(file.originalname).toLowerCase();
  const outputPath = path.join(__dirname, "uploads", "formatted.pdf");

  try {
    if (ext === ".docx") {
      await processDocx(file.path, outputPath);
    } else if (ext === ".pdf") {
      await processPdf(file.path, outputPath);
    } else {
      return res.status(400).send("Please upload a .docx or .pdf file only.");
    }

    res.download(outputPath, "formatted.pdf");
  } catch (err) {
    console.error(err);
    res.status(500).send("Error processing the file.");
  }
});

app.listen(3000, () => console.log("🚀 Server running on http://localhost:3000"));
