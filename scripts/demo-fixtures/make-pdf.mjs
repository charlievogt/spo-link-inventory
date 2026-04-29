#!/usr/bin/env node
// Generate a minimal valid .pdf with text + hyperlink annotations.
//
// Usage: node make-pdf.mjs <output> "Title" "Body" '[["display","https://target"], ...]'
//
// PDFs are upload-byte-stable on SharePoint (unlike .docx), so they're the
// right format for the duplicate-hash demo.

import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';
import { writeFileSync } from 'node:fs';

const [outPath, title, body, linksJson] = process.argv.slice(2);
const links = JSON.parse(linksJson || '[]');

const pdf = await PDFDocument.create();
const page = pdf.addPage([612, 792]); // US Letter
const font = await pdf.embedFont(StandardFonts.Helvetica);
const fontBold = await pdf.embedFont(StandardFonts.HelveticaBold);

let y = 740;
page.drawText(title, { x: 50, y, font: fontBold, size: 18, color: rgb(0, 0, 0) });
y -= 30;
page.drawText(body, { x: 50, y, font, size: 11, color: rgb(0, 0, 0), maxWidth: 500 });
y -= 40;

for (const [display, url] of links) {
  page.drawText(display, { x: 50, y, font, size: 11, color: rgb(0, 0.4, 0.8) });
  const w = font.widthOfTextAtSize(display, 11);
  // Add /Link annotation with /URI action over the text bounds
  const linkAnnot = pdf.context.obj({
    Type: 'Annot',
    Subtype: 'Link',
    Rect: [50, y - 2, 50 + w, y + 11],
    Border: [0, 0, 0],
    A: {
      Type: 'Action',
      S: 'URI',
      URI: pdf.context.obj(url),
    },
  });
  const annotRef = pdf.context.register(linkAnnot);
  const existing = page.node.Annots() ?? pdf.context.obj([]);
  existing.push(annotRef);
  page.node.set(pdf.context.obj('Annots'), existing);
  y -= 22;
}

const bytes = await pdf.save({ useObjectStreams: false });
// deepcode ignore PT: local-only dev fixture generator; outPath is supplied by the developer running the script, no untrusted input source.
writeFileSync(outPath, bytes);
console.log(`wrote ${outPath} (${bytes.length} bytes, ${links.length} hyperlinks)`);
