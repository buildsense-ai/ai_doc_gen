# AI文档生成器 - Web界面

一个智能的文档生成系统，支持ChatGPT风格的聊天界面，可以处理多种文档格式并生成AI驱动的文档。

## 功能特色

- 🤖 ChatGPT风格的聊天界面
- 📁 支持多种文件格式上传（.doc, .docx, .pdf, .txt, .jpg, .png）
- 📋 项目文档管理仪表板
- 🔄 自动模板匹配和填充
- 📄 文档预览和下载
- 🧠 AI驱动的内容生成

## 安装和设置

### 1. 安装依赖
```bash
pip install -r requirements.txt
```

### 2. 设置API密钥
获取OpenRouter API密钥：https://openrouter.ai/keys

设置环境变量：
```bash
# macOS/Linux
export OPENROUTER_API_KEY='your-api-key-here'

# Windows
set OPENROUTER_API_KEY=your-api-key-here
```

或者创建 `.env` 文件：
```
OPENROUTER_API_KEY=your-api-key-here
```

### 3. 启动应用

**Web界面（推荐）:**
```bash
python main.py
```

**命令行版本:**
```bash
python main.py --cli
```

## 使用指南

### 1. 访问Web界面
启动后访问：http://localhost:8000

### 2. 上传文件类型

#### 📋 项目竣工清单
- 支持格式：.json, .txt
- 示例文件：`sample_test_list.json`
- 用途：创建项目文档清单和管理仪表板

#### 📄 多个模板
- 支持格式：.doc, .docx
- 用途：为项目提供文档模板

#### 📂 已填写的文档
- 支持格式：.doc, .docx, .pdf
- 用途：提供参考信息和历史数据

#### 📝 会议纪要和上下文信息
- 支持格式：.txt, .docx, .pdf, .jpg, .png
- 用途：提供项目背景和补充信息

### 3. 工作流程

1. **上传项目清单** → 系统创建项目仪表板
2. **上传模板** → 系统匹配模板到项目
3. **添加上下文信息** → AI提取相关数据
4. **生成文档** → 点击"生成文档"按钮
5. **预览和下载** → 完成后可预览和下载文档

### 4. 侧边栏功能

- 📊 实时显示项目状态
- 🔘 快速访问文档生成功能
- 📥 一键下载完成的文档
- 📈 项目进度跟踪

## 文件结构

```
ai_docClassify/
├── main.py              # 原始命令行程序
├── app.py               # FastAPI Web应用
├── prompt_utils.py      # AI提示工具
├── requirements.txt     # 依赖列表
├── frontend/
│   └── templates/
│       └── index.html   # Web界面
├── uploads/             # 上传文件目录
├── generated_docs/      # 生成文档目录
└── README.md           # 本文件
```

## 技术栈

- **后端**: FastAPI, Python
- **前端**: HTML5, CSS3, JavaScript
- **AI模型**: OpenRouter (Gemini Pro)
- **文档处理**: python-docx, LibreOffice
- **文件上传**: Multipart forms

## 故障排除

### 常见问题

1. **LibreOffice相关错误**
   - 确保已安装LibreOffice
   - macOS: 从官网下载安装包
   - Linux: `sudo apt-get install libreoffice`

2. **API密钥错误**
   - 检查环境变量设置
   - 确认API密钥有效性

3. **端口占用**
   - 默认端口8000，如需更改请修改app.py

4. **文件上传失败**
   - 检查文件格式是否支持
   - 确认文件大小合理

### 开发模式
```bash
# 启动开发服务器（自动重载）
uvicorn app:app --host 0.0.0.0 --port 8000 --reload
```

## 更新日志

- v1.0: 初始版本，支持基本文档生成
- v1.1: 添加Web界面和聊天功能
- v1.2: 集成项目管理仪表板

## 许可证

此项目仅供学习和研究使用。

## 联系方式

如有问题或建议，请联系开发团队。 