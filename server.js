const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const convertRouter = require('./routes/convert');

const app = express();
// 支持从命令行参数或环境变量读取端口号
// 命令行参数格式: node server.js --port 3001
const getPortFromArgs = () => {
  const args = process.argv.slice(2);
  const portIndex = args.indexOf('--port');
  if (portIndex !== -1 && args[portIndex + 1]) {
    return parseInt(args[portIndex + 1], 10);
  }
  return null;
};
const PORT = process.env.PORT || getPortFromArgs() || 3000;

// 创建必要的目录
const uploadsDir = path.join(__dirname, 'uploads');
const outputsDir = path.join(__dirname, 'outputs');

if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

if (!fs.existsSync(outputsDir)) {
  fs.mkdirSync(outputsDir, { recursive: true });
}

// 中间件配置
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// 静态文件服务
app.use(express.static(path.join(__dirname, 'public')));

// 转换路由
app.use('/api', convertRouter);

// 根路径重定向到index.html
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`服务器运行在 http://localhost:${PORT}`);
  console.log(`上传页面: http://localhost:${PORT}`);
});

