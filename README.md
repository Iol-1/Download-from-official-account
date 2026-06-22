# 公众号文章保存工具

将微信公众号文章一键保存为格式化的 Word 文档，方便个人知识库整理。

## 功能特性

- 输入公众号文章链接，自动解析并保存为 `.docx` 文件
- 保留原文标题、作者、发布日期等元信息
- 支持正文排版（标题、加粗、列表、表格、引用等）
- 自动下载文章中的图片并嵌入文档
- 支持批量下载（每行一个链接）
- 内置下载历史记录与索引管理
- 支持从剪贴板智能识别链接

## 环境要求

- Python 3.8+
- Windows 系统

## 安装依赖

```bash
pip install requests beautifulsoup4 lxml python-docx PyQt5 Pillow
```

## 使用方法

### 直接运行

```bash
python wechat_to_word.py
```

### 打包为 EXE

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.png --name=公众号文章保存工具 wechat_to_word.py
```

## 操作说明

1. 启动工具后，在输入框中粘贴公众号文章链接（支持多行批量）
2. 选择保存文件夹
3. 点击"开始下载保存"
4. 下载完成后可直接打开文件或所在目录

快捷键：`Enter` 开始下载 | `Ctrl+V` 从剪贴板识别链接

## 项目结构

```
├── wechat_to_word.py      # 主程序
├── icon.png                # 应用图标
├── 启动工具.bat             # Windows 启动脚本
├── build.bat               # PyInstaller 打包脚本
├── .gitignore
└── README.md
```

## 免责声明

- 本工具仅供**个人学习和研究**使用
- 请尊重原作者的著作权，下载内容不得用于商业用途或公开传播
- 用户需自行承担使用本工具的一切法律责任
- 本项目开发者不对任何滥用行为负责

## License

[MIT License](LICENSE)
