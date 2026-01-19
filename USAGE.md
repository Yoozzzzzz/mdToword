# 使用说明

## 安装依赖

```bash
npm install
```

## 启动服务

```bash
npm start
```

或者使用开发模式（自动重启）：

```bash
npm run dev
```

## 访问服务

启动后，在浏览器中访问：

```
http://localhost:3000
```

## 使用方法

1. 打开浏览器访问 `http://localhost:3000`
2. 点击上传区域或拖拽Markdown文件（.md格式）
3. 选择文件后，点击"开始转换"按钮
4. 等待转换完成，Word文档会自动下载

## 支持的Markdown格式

- **标题**：`# H1` 到 `###### H6`
- **粗体**：`**粗体**` 或 `__粗体__`
- **斜体**：`*斜体*` 或 `_斜体_`
- **代码**：`` `代码` ``
- **代码块**：```代码块```
- **列表**：`- 项目` 或 `1. 编号`
- **引用**：`> 引用文本`
- **删除线**：`~~删除线~~`

## API接口

### POST /api/convert

上传Markdown文件并转换为Word文档。

**请求格式**：
- Content-Type: `multipart/form-data`
- 字段名：`markdownFile`
- 文件类型：`.md`

**响应**：
- 成功：返回Word文档文件流
- 失败：返回JSON错误信息

**示例**：
```bash
curl -X POST -F "markdownFile=@example.md" http://localhost:3000/api/convert -o output.docx
```

## 注意事项

- 文件大小限制：10MB
- 临时文件会在转换完成后自动清理
- 确保Node.js版本 >= 12.0.0

