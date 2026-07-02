const express = require('express');
const cors = require('cors');
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');
const os = require('os');

const app = express();
const PORT = process.env.PORT || 3210;

// pandoc 路径（与项目同目录）
const PANDOC_PATH = path.join(__dirname, 'pandoc.exe');

app.use(cors());
app.use(express.json({ limit: '10mb' }));
app.use(express.static(__dirname));

// Markdown 转 Word
app.post('/api/convert-to-word', (req, res) => {
  const { markdown, filename } = req.body;

  if (!markdown) {
    return res.status(400).json({ error: 'Markdown 内容为空' });
  }

  // 创建临时文件
  const tempDir = os.tmpdir();
  const tempMd = path.join(tempDir, `dcs_${Date.now()}.md`);
  const tempDocx = path.join(tempDir, `dcs_${Date.now()}.docx`);

  fs.writeFileSync(tempMd, markdown, 'utf-8');

  // 调用 pandoc 转换
  const cmd = `"${PANDOC_PATH}" "${tempMd}" -o "${tempDocx}" --from markdown --to docx`;

  exec(cmd, (error, stdout, stderr) => {
    // 清理临时 md 文件
    try { fs.unlinkSync(tempMd); } catch (e) {}

    if (error) {
      console.error('Pandoc 转换失败:', stderr);
      return res.status(500).json({ error: '转换失败: ' + stderr });
    }

    // 读取 docx 文件并返回
    const docxBuffer = fs.readFileSync(tempDocx);

    // 清理临时 docx 文件
    try { fs.unlinkSync(tempDocx); } catch (e) {}

    const outputName = (filename || '试验卡').replace(/\.[^.]+$/, '') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);
    res.send(docxBuffer);
  });
});

app.listen(PORT, () => {
  console.log(`服务器已启动: http://localhost:${PORT}`);
  console.log(`Pandoc 路径: ${PANDOC_PATH}`);
});
