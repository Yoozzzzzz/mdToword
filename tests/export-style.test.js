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
  const bodyStyle = styles.find(s => a(s)['w:styleId'] === 'ReportBody');
  assert.equal(a(bodyStyle['w:pPr']['w:spacing'])['w:line'], '360');
  assert.equal(a(bodyStyle['w:pPr']['w:spacing'])['w:lineRule'], 'auto');
  const heading2Style = styles.find(s => a(s)['w:styleId'] === 'ReportHeading2');
  assert.equal(a(heading2Style['w:pPr']['w:spacing'])['w:before'], '88');
  assert.equal(a(heading2Style['w:pPr']['w:spacing'])['w:after'], '44');
  assert.match(await zip.file('word/footer1.xml').async('string'), /PAGE/);
  assert.match(xml, /ReportTitle/); assert.match(xml, /ReportTableHeader/);
  const items = ['人工录入错误；', '漏录、错录；', '检测结果与设备对应错误；', '后期修改记录；', '检测人员及责任人员难以追溯等风险。'];
  const listZip = await JSZip.loadAsync(await convertMarkdownToDocx(
    '这一过程中存在：\n' + items.map((text, i) => `${['*', '-', '+'][i % 3]} ${text}`).join('\n') +
    '\n* **重点**风险；\n\n项目实施后，将形成：\n1. 有序项目\n\n---'
  ));
  const paragraphs = xml2js(await listZip.file('word/document.xml').async('string'), { compact: true })['w:document']['w:body']['w:p'];
  const lists = paragraphs.filter(p => (Array.isArray(p['w:pPr']['w:pStyle']) ? p['w:pPr']['w:pStyle'] : [p['w:pPr']['w:pStyle']])
    .some(style => a(style)['w:val'] === 'ReportList'));
  assert.equal(lists.length, 6);
  for (const [i, p] of lists.entries()) {
    const props = p['w:pPr'];
    assert.ok(props['w:numPr'], 'native Word bullet');
    assert.equal(a(props['w:ind'])['w:left'], '720');
    assert.equal(a(props['w:ind'])['w:hanging'], '240');
    assert.equal(a(props['w:spacing'])['w:line'], '360');
    if (i < items.length) assert.equal(p['w:r']['w:t']._text, items[i]);
  }
  assert.ok(lists[5]['w:r'][0]['w:rPr']['w:b']);
  const listStyles = xml2js(await listZip.file('word/styles.xml').async('string'), { compact: true })['w:styles']['w:style'];
  const listStyle = listStyles.find(s => a(s)['w:styleId'] === 'ReportList');
  assert.equal(a(listStyle['w:rPr']['w:b'])['w:val'], 'false');
  assert.equal(listStyle['w:pPr']['w:outlineLvl'], undefined);
  assert.equal(paragraphs.length, 9, 'separators do not become list items');
  const orderedZip = await JSZip.loadAsync(await convertMarkdownToDocx('1. 普通项目\n10. **重点**项目\n3、 保留编号'));
  const orderedParagraphs = xml2js(await orderedZip.file('word/document.xml').async('string'), { compact: true })['w:document']['w:body']['w:p'];
  for (const [i, p] of orderedParagraphs.entries()) {
    assert.equal(a(p['w:pPr']['w:pStyle'])['w:val'], 'ReportOrderedList');
    assert.equal(a(p['w:pPr']['w:ind'])['w:left'], '960');
    assert.equal(a(p['w:pPr']['w:ind'])['w:hanging'], '480');
    assert.equal(a(p['w:pPr']['w:tabs']['w:tab'])['w:pos'], '960');
    assert.equal(p['w:r'][0]['w:t']._text, ['1.', '10.', '3.'][i]);
    assert.ok('w:tab' in p['w:r'][0]);
  }
  assert.ok(orderedParagraphs[1]['w:r'][1]['w:rPr']['w:b']);
  const orderedStyle = styles.find(s => a(s)['w:styleId'] === 'ReportOrderedList');
  assert.equal(a(orderedStyle['w:rPr']['w:b'])['w:val'], 'false');
  assert.equal(orderedStyle['w:pPr']['w:outlineLvl'], undefined);
  assert.equal(a(orderedStyle['w:pPr']['w:spacing'])['w:line'], '360');
  assert.equal(a(orderedStyle['w:pPr']['w:spacing'])['w:lineRule'], 'auto');
  console.log('DOCX style smoke test passed');
})().catch(e => { console.error(e); process.exitCode = 1; });

