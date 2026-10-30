import { readFileSync, writeFileSync } from 'node:fs';
import { PDFDocument } from 'pdf-lib';

const page1 = `BT
/F1 12 Tf
72 740 Td
(This is a automation file.) Tj
0 -18 Td
(This is a automation file.) Tj
0 -18 Td
(This is a automation file.) Tj
ET`;

const page2 = `BT
/F1 18 Tf
72 720 Td
(Page 2) Tj
0 -32 Td
/F1 12 Tf
(Second page, used to check previous and next.) Tj
ET`;

const objects = [
  '<< /Type /Catalog /Pages 2 0 R >>',
  '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 >>',
  '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 5 0 R /Resources << /Font << /F1 6 0 R >> >> >>',
  '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 7 0 R /Resources << /Font << /F1 6 0 R >> >> >>',
  null,
  '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  null,
];

function stream(text) {
  const data = text.endsWith('\n') ? text : `${text}\n`;
  return `<< /Length ${Buffer.byteLength(data, 'latin1')} >>\nstream\n${data}endstream`;
}

objects[4] = stream(page1);
objects[6] = stream(page2);

let output = '%PDF-1.4\n';
const offsets = [0];
objects.forEach((body, index) => {
  offsets.push(Buffer.byteLength(output, 'latin1'));
  output += `${index + 1} 0 obj\n${body}\nendobj\n`;
});

const xrefAt = Buffer.byteLength(output, 'latin1');
output += `xref\n0 ${objects.length + 1}\n`;
output += '0000000000 65535 f \n';
for (let index = 1; index <= objects.length; index += 1) {
  output += `${String(offsets[index]).padStart(10, '0')} 00000 n \n`;
}
output += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefAt}\n%%EOF\n`;

const sampleUrl = new URL('../demo/sample.pdf', import.meta.url);
writeFileSync(sampleUrl, output, 'latin1');

const doc = await PDFDocument.load(readFileSync(sampleUrl));
const page = doc.getPages()[1];
const form = doc.getForm();
const name = form.createTextField('Name');
name.addToPage(page, { x: 72, y: 640, width: 200, height: 20 });
const agree = form.createCheckBox('Agree');
agree.addToPage(page, { x: 72, y: 600, width: 16, height: 16 });
form.updateFieldAppearances();
writeFileSync(sampleUrl, await doc.save());
