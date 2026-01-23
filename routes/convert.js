const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType } = require('docx');

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

// Markdown转Word转换函数
async function convertMarkdownToDocx(markdownContent) {
  // 将Markdown转换为docx元素
  const children = parseMarkdownToDocx(markdownContent);
  
  // 创建Word文档
  const doc = new Document({
    sections: [{
      properties: {},
      children: children
    }]
  });
  
  // 生成Buffer
  return await Packer.toBuffer(doc);
}

// 解析Markdown内容为docx元素
function parseMarkdownToDocx(content) {
  const lines = content.split('\n');
  const children = [];
  let i = 0;
  
  while (i < lines.length) {
    const line = lines[i].trim();
    
    if (!line) {
      // 空行
      children.push(new Paragraph({ text: '' }));
      i++;
      continue;
    }
    
    // 检查是否是表格行（包含|符号的行）
    const nextLine = i + 1 < lines.length ? lines[i + 1].trim() : '';
    if (isTableHeader(line) && isTableSeparator(nextLine)) {
      // 找到表格，解析整个表格
      const tableData = parseTable(lines, i);
      children.push(tableData.table);
      i = tableData.nextLineIndex; // 跳过已处理的表格行
      continue;
    }
    
    // 标题处理（支持标题中的内联格式）
    if (line.startsWith('# ')) {
      const titleText = line.substring(2);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_1
      }));
    } else if (line.startsWith('## ')) {
      const titleText = line.substring(3);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_2
      }));
    } else if (line.startsWith('### ')) {
      const titleText = line.substring(4);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_3
      }));
    } else if (line.startsWith('#### ')) {
      const titleText = line.substring(5);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_4
      }));
    } else if (line.startsWith('##### ')) {
      const titleText = line.substring(6);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_5
      }));
    } else if (line.startsWith('###### ')) {
      const titleText = line.substring(7);
      const runs = parseInlineFormatting(titleText);
      children.push(new Paragraph({
        children: runs,
        heading: HeadingLevel.HEADING_6
      }));
    } else if (line.startsWith('- ') || line.startsWith('* ') || /^\d+\.\s/.test(line)) {
      // 列表项
      const listText = line.replace(/^[-*]\s+/, '').replace(/^\d+\.\s+/, '');
      // 处理列表项中的内联格式（粗体、斜体等）
      const runs = parseInlineFormatting(listText);
      children.push(new Paragraph({
        children: runs,
        bullet: {
          level: 0
        }
      }));
    } else if (line.startsWith('> ')) {
      // 引用
      const quoteText = line.substring(2);
      // 处理引用中的内联格式
      const runs = parseInlineFormatting(quoteText);
      children.push(new Paragraph({
        children: runs,
        indent: {
          left: 720 // 0.5 inch
        }
      }));
    } else if (line.startsWith('```')) {
      // 代码块开始，收集代码块内容
      const codeLines = [];
      i++; // 跳过代码块标记行
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        codeLines.push(lines[i]);
        i++;
      }
      i++; // 跳过结束标记行
      const codeText = codeLines.join('\n');
      children.push(new Paragraph({
        text: codeText,
        spacing: {
          after: 200
        }
      }));
    } else {
      // 普通段落，处理内联格式
      const runs = parseInlineFormatting(line);
      children.push(new Paragraph({
        children: runs
      }));
    }
    i++;
  }
  
  return children;
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
  const separatorLine = lines[startIndex + 1];
  const headerCells = parseTableCells(headerLine);
  
  // 解析分隔符行获取对齐信息
  const alignmentInfo = parseAlignment(separatorLine);
  
  // 创建表格
  const rows = [];
  
  // 添加表头行
  const headerRow = new TableRow({
    children: headerCells.map((cell, idx) => {
      const runs = parseInlineFormatting(cell.trim());
      return new TableCell({
        children: [new Paragraph({ children: runs })],
        width: {
          size: 1000, // 平均分配宽度
          type: WidthType.PERCENTAGE
        }
      });
    })
  });
  rows.push(headerRow);
  
  // 解析数据行
  let i = startIndex + 2;
  while (i < lines.length) {
    const line = lines[i].trim();
    if (!line) {
      // 空行则结束表格解析
      break;
    }
    if (!isTableHeader(line)) {
      // 不是表格行，则结束表格解析
      break;
    }
    
    const dataCells = parseTableCells(line);
    const dataRow = new TableRow({
      children: dataCells.map((cell, idx) => {
        const runs = parseInlineFormatting(cell.trim());
        return new TableCell({
          children: [new Paragraph({ children: runs })],
          width: {
            size: 1000, // 平均分配宽度
            type: WidthType.PERCENTAGE
          }
        });
      })
    });
    rows.push(dataRow);
    i++;
  }
  
  const table = new Table({
    rows: rows
  });
  
  return {
    table: table,
    nextLineIndex: i
  };
}

// 解析表格单元格
function parseTableCells(line) {
  // 移除首尾的|符号
  const trimmedLine = line.replace(/^\|+/, '').replace(/\|+$/, '');
  // 按|分割，但要处理转义的|符号
  return trimmedLine.split(/\|(?![^`]*`)/);
}

// 解析对齐信息
function parseAlignment(separatorLine) {
  const separators = parseTableCells(separatorLine);
  return separators.map(sep => {
    sep = sep.trim();
    const startsWithColon = sep.startsWith(':');
    const endsWithColon = sep.endsWith(':');
    
    if (startsWithColon && endsWithColon) {
      return 'center'; // 居中对齐
    } else if (startsWithColon) {
      return 'left';   // 左对齐
    } else if (endsWithColon) {
      return 'right';  // 右对齐
    } else {
      return 'left';   // 默认左对齐
    }
  });
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
        runs.push(new TextRun(plainText));
      }
    }
    
    // 添加格式化的文本
    const options = {};
    if (marker.type === 'bold') {
      options.bold = true;
    } else if (marker.type === 'italic') {
      options.italics = true;
    } else if (marker.type === 'code') {
      options.font = 'Courier New';
    } else if (marker.type === 'strikethrough') {
      options.strike = true;
    }
    
    runs.push(new TextRun({
      text: marker.content,
      ...options
    }));
    
    pos = marker.end;
  }
  
  // 添加剩余的普通文本
  if (pos < text.length) {
    const plainText = text.substring(pos);
    if (plainText) {
      runs.push(new TextRun(plainText));
    }
  }
  
  // 如果没有找到任何格式标记，返回普通文本
  if (runs.length === 0) {
    runs.push(new TextRun(text));
  }
  
  return runs;
}

module.exports = router;