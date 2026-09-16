const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, Footer, PageNumber, LineRuleType, VerticalAlign, convertMillimetersToTwip } = require('docx');

const router = express.Router();

// 配置multer存储
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadsDir = path.join(__dirname, '..', 'uploads');
    cb(null, uploadsDir);
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    // 保持原始文件名（multer会自动处理编码）
    cb(null, uniqueSuffix + '-' + file.originalname);
  }
});

// 文件过滤器：只允许.md文件
const fileFilter = (req, file, cb) => {
  if (file.mimetype === 'text/markdown' || 
      file.originalname.toLowerCase().endsWith('.md')) {
    cb(null, true);
  } else {
    cb(new Error('只支持Markdown文件（.md格式）'), false);
  }
};

const upload = multer({
  storage: storage,
  fileFilter: fileFilter,
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB限制
  }
});

// 辅助函数：正确处理中文文件名
function decodeFileName(filename) {
  if (!filename) return '';
  
  // 尝试多种解码方式
  try {
    // 如果文件名包含URL编码的字符，先解码
    if (filename.includes('%')) {
      return decodeURIComponent(filename);
    }
    // 如果文件名是ISO-8859-1编码的中文（常见情况）
    // 尝试从latin1转换为UTF-8
    if (/[\x80-\xFF]/.test(filename) && !/[^\x00-\xFF]/.test(filename)) {
      // 可能是latin1编码，需要转换
      return Buffer.from(filename, 'latin1').toString('utf8');
    }
    // 直接返回（已经是UTF-8）
    return filename;
  } catch (e) {
    // 如果所有解码都失败，返回原始文件名
    return filename;
  }
}

// 转换路由
router.post('/convert', upload.single('markdownFile'), async (req, res) => {
  let uploadedFilePath = null;
  let outputFilePath = null;

  try {
    // 检查文件是否存在
    if (!req.file) {
      return res.status(400).json({
        success: false,
        error: '请上传Markdown文件'
      });
    }

    uploadedFilePath = req.file.path;
    const markdownContent = fs.readFileSync(uploadedFilePath, 'utf-8');

    // 将Markdown转换为Word文档
    const docxBuffer = await convertMarkdownToDocx(markdownContent);

    // 生成输出文件名（确保正确处理中文文件名）
    const originalFileName = decodeFileName(req.file.originalname);
    const originalName = originalFileName.replace(/\.md$/i, '');
    const outputFileName = `${originalName}_${Date.now()}.docx`;
    outputFilePath = path.join(__dirname, '..', 'outputs', outputFileName);

    // 保存Word文档
    fs.writeFileSync(outputFilePath, docxBuffer);

    // 设置响应头，返回文件供下载
    // 处理中文文件名：使用RFC 5987标准编码，兼容不同浏览器
    const fileName = `${originalName}.docx`;
    
    // 确保文件名是UTF-8编码
    const utf8FileName = Buffer.from(fileName, 'utf8').toString('utf8');
    
    // 对文件名进行URL编码（用于filename*参数）
    const encodedFileName = encodeURIComponent(utf8FileName);
    
    // 检查文件名是否包含非ASCII字符
    const hasNonAscii = /[^\x00-\x7F]/.test(utf8FileName);
    
    // 使用RFC 5987格式：
    // - filename: 对于非ASCII字符，使用Base64编码或URL编码
    // - filename*: 使用UTF-8编码（现代浏览器优先使用）
    let fallbackFileName;
    if (hasNonAscii) {
      // 对于包含中文的文件名，使用Base64编码作为fallback
      // 但更简单的方式是直接使用URL编码
      fallbackFileName = encodedFileName;
    } else {
      fallbackFileName = utf8FileName;
    }
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    // 使用RFC 5987格式，确保中文文件名正确显示
    res.setHeader('Content-Disposition', `attachment; filename="${fallbackFileName}"; filename*=UTF-8''${encodedFileName}`);
    
    // 发送文件
    res.send(docxBuffer);

    // 延迟删除临时文件（确保文件已发送）
    setTimeout(() => {
      try {
        if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
          fs.unlinkSync(uploadedFilePath);
        }
        if (outputFilePath && fs.existsSync(outputFilePath)) {
          fs.unlinkSync(outputFilePath);
        }
      } catch (err) {
        console.error('清理临时文件失败:', err);
      }
    }, 1000);

  } catch (error) {
    console.error('转换错误:', error);
    
    // 清理临时文件
    try {
      if (uploadedFilePath && fs.existsSync(uploadedFilePath)) {
        fs.unlinkSync(uploadedFilePath);
      }
      if (outputFilePath && fs.existsSync(outputFilePath)) {
        fs.unlinkSync(outputFilePath);
      }
    } catch (cleanupError) {
      console.error('清理临时文件失败:', cleanupError);
    }

    res.status(500).json({
      success: false,
      error: error.message || '转换失败，请检查Markdown文件格式'
    });
  }
});

// 固定报告规范，详见 DOCUMENT_STYLE.md。字号为半磅，距离为 twip。
const BODY_FONT = '仿宋_GB2312';
const spacing = (before = 0, after = 0) => ({
  line: 440, lineRule: LineRuleType.EXACTLY,
  before: Math.round(before * 440), after: Math.round(after * 440)
});
const reportStyles = [
  ['ReportBody', '正文', BODY_FONT, 24, false, 'left', 0, 0],
  ['ReportTitle', '文档大标题', '黑体', 44, false, 'center', 0, 1],
  ['ReportSubtitle', '报告副标题', '黑体', 30, false, 'center', 0, 0.8],
  ['ReportHeading1', '一级标题', '黑体', 32, false, 'left', 0.8, 0.5],
  ['ReportHeading2', '二级标题', '黑体', 28, false, 'left', 0.5, 0.3],
  ['ReportHeading3', '三级标题', '楷体_GB2312', 28, true, 'left', 0.3, 0],
  ['ReportHeading4', '四级列表项目标题', BODY_FONT, 24, true, 'left', 0, 0],
  ['ReportCaption', '表格标题', '黑体', 21, false, 'center', 0, 0],
  ['ReportTableHeader', '表头', '黑体', 21, false, 'center', 0, 0],
  ['ReportTableBody', '表格内容', BODY_FONT, 21, false, 'left', 0, 0],
  ['ReportNote', '表下注释', BODY_FONT, 21, false, 'left', 0, 0],
  ['ReportPageNumber', '页码', '宋体', 21, false, 'center', 0, 0]
].map(([id, name, font, size, bold, alignment, before, after]) => ({
  id, name,
  run: { font: { name: font, eastAsia: font }, size, bold, color: '000000' },
  paragraph: {
    alignment, spacing: spacing(before, after),
    // 小四号为 12 磅，首行 24 磅即两个汉字宽度。
    indent: { firstLine: id === 'ReportBody' ? 480 : 0 },
    ...(/^ReportHeading/.test(id) ? { outlineLevel: Number(id.slice(-1)) - 1 } : {})
  }
}));

async function convertMarkdownToDocx(markdownContent) {
  const doc = new Document({
    styles: {
      default: { document: {
        run: { font: { name: BODY_FONT, eastAsia: BODY_FONT }, size: 24, color: '000000' },
        paragraph: { spacing: spacing() }
      } },
      paragraphStyles: reportStyles
    },
    sections: [{
      properties: { page: {
        size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
        margin: { top: 1440, bottom: 1440,
          left: convertMillimetersToTwip(28), right: convertMillimetersToTwip(26), footer: 720 }
      } },
      footers: { default: new Footer({ children: [new Paragraph({
        style: 'ReportPageNumber', children: [new TextRun({ children: [PageNumber.CURRENT] })]
      })] }) },
      children: parseMarkdownToDocx(markdownContent)
    }]
  });
  return Packer.toBuffer(doc);
}

function reportParagraph(text, style = 'ReportBody', options = {}) {
  const parts = String(text).split(/<br\s*\/?>/gi);
  return new Paragraph({ style, children: parts.flatMap((part, index) =>
    index ? [new TextRun({ text: part, break: 1 })] : parseInlineFormatting(part)
  ), ...options });
}

function parseMarkdownToDocx(content) {
  const lines = content.replace(/\uFF5C/g, '|').split('\n');
  const children = [];
  let hasTitle = false;
  let afterTable = false;
  let i = 0;
  while (i < lines.length) {
    const rawLine = lines[i].trim();
    // Word/AI 生成的 Markdown 常把标题误写成 ▪、• 或 · 开头；它们不是报告列表。
    const line = rawLine.replace(/^[▪•·]\s*/, '');
    if (!line) { i++; continue; }
    // Markdown 分隔线不是报告正文。
    if (/^(?:[-*_])(?:\s*[-*_]){2,}$/.test(line)) { i++; continue; }
    if (isTableHeader(line) && isTableSeparator((lines[i + 1] || '').trim())) {
      const result = parseTable(lines, i);
      children.push(result.table);
      i = result.nextLineIndex;
      afterTable = true;
      continue;
    }
    if (line.startsWith('```')) {
      i++;
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        // 代码保留原文，字体和行距仍遵守报告规范。
        children.push(new Paragraph({ style: 'ReportBody', text: lines[i++] }));
      }
      if (i < lines.length) i++;
      afterTable = false;
      continue;
    }
    const heading = line.match(/^(#{1,6})\s+(.+?)(?:\s+#+)?$/);
    const text = heading ? heading[2] : line;
    const plain = text.replace(/\*\*|__/g, '').trim();
    let next = i + 1;
    while (next < lines.length && !lines[next].trim()) next++;
    const beforeTable = isTableHeader((lines[next] || '').trim()) &&
      isTableSeparator((lines[next + 1] || '').trim());
    // “《项目》项目验收报告”是常见的合并标题写法，拆成规范的两段。
    const mergedTitle = plain.match(/^(.+?)\s*项目验收报告$/);
    if (mergedTitle && !hasTitle && !beforeTable) {
      children.push(reportParagraph(mergedTitle[1].trim(), 'ReportTitle'));
      children.push(reportParagraph('项目验收报告', 'ReportSubtitle'));
      hasTitle = true;
    } else if (plain === '项目验收报告') {
      children.push(reportParagraph(text, 'ReportSubtitle'));
    } else if (/^表\s*[0-9一二三四五六七八九十]+/.test(plain) && beforeTable) {
      children.push(reportParagraph(text, 'ReportCaption'));
    } else if (afterTable && /^(?:注|说明|备注)(?:[：:]|\s|$)/.test(plain)) {
      children.push(reportParagraph(text, 'ReportNote'));
      i++;
      continue;
    } else if (/^[一二三四五六七八九十百]+[、．.]/.test(plain)) {
      children.push(reportParagraph(text, 'ReportHeading1'));
    } else if (/^[（(][一二三四五六七八九十百]+[）)]/.test(plain)) {
      children.push(reportParagraph(text, 'ReportHeading2'));
    } else if (heading) {
      if (heading[1].length === 1 && !hasTitle) {
        children.push(reportParagraph(text, 'ReportTitle'));
        hasTitle = true;
      } else {
        const level = /^\d+[.、．]/.test(plain)
          ? 3 : Math.min(4, Math.max(1, heading[1].length - 1));
        children.push(reportParagraph(text, `ReportHeading${level}`));
      }
    } else if (/^[-*]\s+|^\d+[.、．]\s+/.test(line)) {
      // 报告中的无序列表通常是“项目名称/承担单位”字段，不输出额外项目符号。
      // 有序列表保留原编号，统一采用四级列表项目标题样式。
      const listText = line.replace(/^[-*]\s+/, '').replace(/^(\d+)[.、．]\s+/, '$1. ');
      children.push(reportParagraph(listText, 'ReportHeading4'));
    } else {
      children.push(reportParagraph(line.replace(/^>\s*/, '')));
    }
    afterTable = false;
    i++;
  }
  return children.length ? children : [reportParagraph('')];
}

// 检查是否是表格头部
function isTableHeader(line) {
  return line.startsWith('|') && line.endsWith('|');
}

// 检查是否是表格分隔符行
function isTableSeparator(line) {
  // 分隔符行通常由|, -, :组成，例如: |---|:---:|---:|
  return line.startsWith('|') && line.endsWith('|') && /^[\|\-:\s]+$/.test(line);
}

// 解析表格
function parseTable(lines, startIndex) {
  const headerLine = lines[startIndex];
  const headerCells = parseTableCells(headerLine);
  const createRow = (cells, header = false) => new TableRow({
    tableHeader: header,
    children: cells.map(cell => {
      const value = cell.trim();
      const numeric = /^[+−-]?(?:[¥￥$]\s*)?\d+(?:[,，]\d{3})*(?:[.．]\d+)?\s*[%％]?$/.test(
        value.replace(/\*\*|__|`/g, '')
      );
      return new TableCell({
        verticalAlign: VerticalAlign.CENTER,
        width: { size: 100 / headerCells.length, type: WidthType.PERCENTAGE },
        children: [reportParagraph(value, header ? 'ReportTableHeader' : 'ReportTableBody', {
          alignment: header || numeric ? 'center' : 'left'
        })]
      });
    })
  });
  const rows = [createRow(headerCells, true)];
  let i = startIndex + 2;
  while (i < lines.length && isTableHeader(lines[i].trim())) {
    rows.push(createRow(parseTableCells(lines[i].trim())));
    i++;
  }
  return {
    table: new Table({ rows, width: { size: 100, type: WidthType.PERCENTAGE } }),
    nextLineIndex: i
  };
}

// 解析表格单元格
function parseTableCells(line) {
  // 将全角竖线 ｜(U+FF5C) 统一为半角 |，避免无法分割
  const normalizedLine = line.trim().replace(/\uFF5C/g, '|');
  // 移除首尾的|符号
  const trimmedLine = normalizedLine.replace(/^\|+/, '').replace(/\|+$/, '');
  // 按|分割；原正则 (?! [^`]* `) 在单元格含反引号时会误判导致整行落入第一格，改为简单分割
  return trimmedLine.split('|').map(cell => cell.trim()).filter((cell, index, cells) =>
    !(index === cells.length - 1 && cell === '')
  );
}

// 只转换展示文字；代码、网址、邮箱和数字内部标点保留，避免改变技术数据。
function chinesePunctuation(text) {
  const punctuation = { ',': '，', ':': '：', ';': '；', '!': '！', '?': '？',
    '(': '（', ')': '）', '[': '【', ']': '】', '.': '。' };
  return text.replace(/https?:\/\/[^\s<>，。；！？）]+|[\w.+-]+@[\w.-]+\.[a-z]+|\d+(?:[.,]\d+)+|[,:;!?()\[\]]|\.(?!\s)|"([^"\n]*)"|'([^'\n]*)'/gi,
    (match, doubleQuote, singleQuote) => doubleQuote !== undefined ? `“${chinesePunctuation(doubleQuote)}”` :
      singleQuote !== undefined ? `‘${chinesePunctuation(singleQuote)}’` : punctuation[match] || match);
}

function makeTextRun(text, options = {}) {
  // docx 将 TextRun 的换行文本序列化为 w:br。
  return new TextRun({ ...options, text: String(text).replace(/<br\s*\/?>/gi, '\n') });
}

// 解析内联格式（粗体、斜体、代码等）
function parseInlineFormatting(text) {
  const runs = [];
  let currentIndex = 0;
  
  // 匹配粗体 **text** 或 __text__
  const boldRegex = /\*\*(.*?)\*\*|__(.*?)__/g;
  // 匹配斜体 *text* 或 _text_
  const italicRegex = /(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.*?)(?<!_)_(?!_)/g;
  // 匹配代码 `code`
  const codeRegex = /`([^`]+)`/g;
  // 匹配删除线 ~~text~~
  const strikethroughRegex = /~~(.*?)~~/g;
  
  // 简单的实现：先处理代码，再处理粗体，最后处理斜体
  const parts = [];
  let lastIndex = 0;
  
  // 收集所有格式标记
  const markers = [];
  
  // 查找所有格式标记
  let match;
  
  // 代码标记
  while ((match = codeRegex.exec(text)) !== null) {
    markers.push({
      start: match.index,
      end: match.index + match[0].length,
      type: 'code',
      content: match[1]
    });
  }
  
  // 粗体标记
  while ((match = boldRegex.exec(text)) !== null) {
    markers.push({
      start: match.index,
      end: match.index + match[0].length,
      type: 'bold',
      content: match[1] || match[2]
    });
  }
  
  // 斜体标记
  while ((match = italicRegex.exec(text)) !== null) {
    const start = match.index;
    const end = start + match[0].length;
    // 检查是否与粗体重叠
    const overlapsBold = markers.some(m => m.type === 'bold' && start < m.end && end > m.start);
    if (!overlapsBold) {
      markers.push({
        start: start,
        end: end,
        type: 'italic',
        content: match[1] || match[2]
      });
    }
  }
  
  // 删除线标记
  while ((match = strikethroughRegex.exec(text)) !== null) {
    markers.push({
      start: match.index,
      end: match.index + match[0].length,
      type: 'strikethrough',
      content: match[1]
    });
  }
  
  // 按位置排序
  markers.sort((a, b) => a.start - b.start);
  
  // 过滤重叠的标记（优先保留代码，然后是粗体，最后是斜体）
  const filteredMarkers = [];
  for (const marker of markers) {
    let overlaps = false;
    for (const existing of filteredMarkers) {
      // 检查是否重叠
      if (marker.start < existing.end && marker.end > existing.start) {
        overlaps = true;
        break;
      }
    }
    if (!overlaps) {
      filteredMarkers.push(marker);
    }
  }
  
  // 重新排序
  filteredMarkers.sort((a, b) => a.start - b.start);
  
  // 构建runs
  let pos = 0;
  for (const marker of filteredMarkers) {
    // 添加标记前的普通文本
    if (marker.start > pos) {
      const plainText = text.substring(pos, marker.start);
      if (plainText) {
        runs.push(makeTextRun(chinesePunctuation(plainText)));
      }
    }
    
    // 添加格式化的文本
    const options = {};
    if (marker.type === 'bold') {
      options.bold = true;
    } else if (marker.type === 'italic') {
      options.italics = true;

    } else if (marker.type === 'strikethrough') {
      options.strike = true;
    }
    
    runs.push(makeTextRun(marker.type === 'code' ? marker.content : chinesePunctuation(marker.content), options));
    
    pos = marker.end;
  }
  
  // 添加剩余的普通文本
  if (pos < text.length) {
    const plainText = text.substring(pos);
    if (plainText) {
      runs.push(makeTextRun(chinesePunctuation(plainText)));
    }
  }
  
  // 如果没有找到任何格式标记，返回普通文本
  if (runs.length === 0) {
    runs.push(makeTextRun(chinesePunctuation(text)));
  }
  
  return runs;
}

module.exports = router;
module.exports.convertMarkdownToDocx = convertMarkdownToDocx;
