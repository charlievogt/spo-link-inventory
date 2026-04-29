#!/usr/bin/env node
// Generate a minimal valid .docx with hyperlinks for demo content.
//
// Usage: node make-docx.mjs <output-path> "Title" "Body paragraph text" '[["hyperlink-display","https://target"], ...]'
//
// The output is a real OOXML .docx file that the func's documentLinkExtractor
// will see as a real hyperlink-bearing document.

import JSZip from 'jszip';
import { writeFileSync } from 'node:fs';

const [outPath, title, body, linksJson] = process.argv.slice(2);
if (!outPath) {
  console.error('Usage: make-docx.mjs <out> <title> <body> <linksJson>');
  process.exit(1);
}
const links = JSON.parse(linksJson || '[]'); // [["display","url"], ...]

const zip = new JSZip();

zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

zip.folder('_rels').file('.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

const linkRels = links
  .map(([, url], i) =>
    `<Relationship Id="rId${100 + i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${url
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/"/g, '&quot;')}" TargetMode="External"/>`,
  )
  .join('');

zip.folder('word').folder('_rels').file('document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${linkRels}</Relationships>`);

const linkParas = links
  .map(([display], i) => `
  <w:p>
    <w:hyperlink r:id="rId${100 + i}">
      <w:r><w:t xml:space="preserve">${display.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</w:t></w:r>
    </w:hyperlink>
  </w:p>`)
  .join('');

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t xml:space="preserve">${title}</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">${body}</w:t></w:r></w:p>${linkParas}
  </w:body>
</w:document>`;

zip.folder('word').file('document.xml', documentXml);

const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
// deepcode ignore PT: local-only dev fixture generator; outPath is supplied by the developer running the script, no untrusted input source.
writeFileSync(outPath, buf);
console.log(`wrote ${outPath} (${buf.length} bytes, ${links.length} hyperlinks)`);
