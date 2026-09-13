const assert = require('node:assert/strict');
const JSZip = require('jszip');
const { xml2js } = require('xml-js');
const { convertMarkdownToDocx } = require('../routes/convert');
const a = x => x == null ? {} : x._attributes || {};
(async () => {
  const data = await convertMarkdownToDocx('# Title\n\n## Subtitle\n\n## Section\nBody.\n\nTable 1\n| A | Number |\n| --- | --- |\n| x | 1 |');
  const zip = await JSZip.loadAsync(data);
  const xml = await zip.file('word/document.xml').async('string');
  assert.match(xml, /w:pgSz w:w="11905" w:h="16837"/);
  assert.match(xml, /w:pgMar w:top="1440" w:right="1474" w:bottom="1440" w:left="1587"/);
  const styles = xml2js(await zip.file('word/styles.xml').async('string'), { compact: true })['w:styles']['w:style'];
  const ids = ['ReportBody','ReportTitle','ReportSubtitle','ReportHeading1','ReportHeading2','ReportHeading3','ReportHeading4','ReportCaption','ReportTableHeader','ReportTableBody','ReportNote','ReportPageNumber'];
  for (const id of ids) assert.ok(styles.some(s => a(s)['w:styleId'] === id), id);
  assert.match(await zip.file('word/footer1.xml').async('string'), /PAGE/);
  assert.match(xml, /ReportTitle/); assert.match(xml, /ReportTableHeader/);
  console.log('DOCX style smoke test passed');
})().catch(e => { console.error(e); process.exitCode = 1; });

