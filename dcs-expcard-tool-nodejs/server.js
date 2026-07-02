const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

// 检测是否在 SEA 环境中运行
const isSEA = !process.argv[1] || process.argv[1].endsWith('.exe') || process.argv[1].includes('sea');

// 加载外部配置文件（如果存在）
let EXPCARD_CONFIG;
const configPath = path.join(process.cwd(), 'config.js');

if (fs.existsSync(configPath)) {
  try {
    const configContent = fs.readFileSync(configPath, 'utf-8');
    const cleaned = configContent
      .replace(/if\s*\(typeof\s+module\s*!==\s*['"]undefined['"]\s*&&\s*module\.exports\)\s*\{[\s\S]*?\}/g, '')
      .replace(/module\.exports\s*=\s*EXPCARD_CONFIG\s*;?/g, '');
    eval(cleaned);
    EXPCARD_CONFIG = typeof EXPCARD_CONFIG !== 'undefined' ? EXPCARD_CONFIG : getDefaultConfig();
    console.log('Loaded config.js');
  } catch (err) {
    console.warn('Config load failed, using default:', err.message);
    EXPCARD_CONFIG = getDefaultConfig();
  }
} else {
  console.log('No config.js found, using default');
  EXPCARD_CONFIG = getDefaultConfig();
}

function getDefaultConfig() {
  return {
    LOGIC_OPERATORS: { '与': '且', '且': '且', '或': '或', '或取反': '或取反' },
    SPECIAL_SEPARATORS: { '或延时': '或', '与延时': '且' },
    SKIP_HEADERS: ['序号', '条件确认'],
    SECTION_HEADERS: ['试验条件', '试验恢复', '结论', '存在问题']
  };
}

const app = express();
const PORT = process.env.PORT || 3210;

// pandoc 路径
const PANDOC_PATH = isSEA 
  ? path.join(process.cwd(), 'pandoc.exe')
  : path.join(__dirname, 'pandoc.exe');

app.use(cors());
app.use(express.json({ limit: '10mb' }));

// 读取内嵌的静态文件
function readEmbeddedFile(filename) {
  try {
    return fs.readFileSync(path.join(__dirname, filename), 'utf-8');
  } catch (e) {
    return null;
  }
}

// 首页
app.get('/', (req, res) => {
  const indexHtml = readEmbeddedFile('index.html');
  if (indexHtml) {
    res.type('html').send(indexHtml);
  } else {
    res.status(404).send('index.html not found');
  }
});

// 静态文件
app.get('/convert.js', (req, res) => {
  const content = readEmbeddedFile('convert.js');
  if (content) {
    res.type('js').send(content);
  } else {
    res.status(404).send('convert.js not found');
  }
});

app.get('/config.js', (req, res) => {
  // 优先返回外部配置文件
  if (fs.existsSync(configPath)) {
    try {
      const content = fs.readFileSync(configPath, 'utf-8');
      res.type('js').send(content);
    } catch (e) {
      res.status(500).send('Error reading config.js');
    }
  } else {
    // 返回默认配置
    const defaultConfig = `const EXPCARD_CONFIG = ${JSON.stringify(EXPCARD_CONFIG, null, 2)};`;
    res.type('js').send(defaultConfig);
  }
});

// Markdown 转 Word
app.post('/api/convert-to-word', (req, res) => {
  const { markdown, filename } = req.body;

  if (!markdown) {
    return res.status(400).json({ error: 'Markdown content is empty' });
  }

  const tempDir = os.tmpdir();
  const tempMd = path.join(tempDir, `dcs_${Date.now()}.md`);
  const tempDocx = path.join(tempDir, `dcs_${Date.now()}.docx`);

  fs.writeFileSync(tempMd, markdown, 'utf-8');

  const cmd = `"${PANDOC_PATH}" "${tempMd}" -o "${tempDocx}" --from markdown --to docx`;

  exec(cmd, (error, stdout, stderr) => {
    try { fs.unlinkSync(tempMd); } catch (e) {}

    if (error) {
      console.error('Pandoc conversion failed:', stderr);
      return res.status(500).json({ error: 'Conversion failed: ' + stderr });
    }

    const docxBuffer = fs.readFileSync(tempDocx);
    try { fs.unlinkSync(tempDocx); } catch (e) {}

    const outputName = (filename || 'test').replace(/\.[^.]+$/, '') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);
    res.send(docxBuffer);
  });
});

app.listen(PORT, () => {
  console.log(`Server started: http://localhost:${PORT}`);
  console.log(`Pandoc path: ${PANDOC_PATH}`);
});
