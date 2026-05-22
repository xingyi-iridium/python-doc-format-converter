# 📄 文件格式转换助手

一个基于 Streamlit 的智能文件格式转换工具，支持 PDF、Word、TXT 三种格式互转，集成硅基流动多模态 AI 平台实现图片内容理解。
## 项目背景
“日常工作中经常需要在不同文档格式间转换，市面上的工具要么收费，要么上传敏感文件有隐私风险。这个项目是我独立开发的本地化解决方案，核心功能完全在本地完成，无需将文件上传至第三方服务器。”

## 项目结构
file-converter/
├── agent_txt_word_pdf.py    # 主程序
├── README.md                # 项目文档
├── requirements.txt         # 依赖清单
├── .env.example             # 环境变量模板
├── uploads/                 # 临时上传目录
└── outputs/                 # 临时输出目录

## ✨ 功能特点

- ✅ 支持 PDF ↔ Word ↔ TXT 六种转换组合
- ✅ 自动检测文档中的图片
- ✅ 转换为 TXT 时智能提醒图片丢失风险
- ✅ PDF 转 Word 时提取并保留图片
- ✅ 可选 AI 图片内容分析（需配置 API）
- 🎯 简洁的 Web 界面，拖拽上传即可使用

## 📋 支持的格式

| 输入格式 | 输出格式 | 说明 |
|---------|---------|------|
| PDF | Word (.docx) | 提取文字和图片，排版需人工调整 |
| PDF | TXT | 纯文本提取，图片会丢失 |
| Word (.docx) | PDF | 需要 docx2pdf 和 Microsoft Word |
| Word (.docx) | TXT | 纯文本提取，图片会丢失 |
| TXT | PDF | 使用 reportlab 生成基础 PDF |
| TXT | Word (.docx) | 基础段落格式 |

## 🔧 环境配置

### 1. 安装依赖

```bash
pip install streamlit PyMuPDF python-docx Pillow reportlab openai python-dotenv docx2pdf
```


### 2. 配置 API 密钥（可选）

创建 `.env` 文件：

```env
SILICONFLOW_API_KEY=your_api_key_here
MULTIMODAL_MODEL=Pro/Qwen/Qwen3-VL-32B-Instruct
```


> 💡 不配置 API 也可以使用基础转换功能，仅多模态图片分析不可用

## 🚀 启动命令

```bash
streamlit run agent_txt_word_pdf.py
```


浏览器会自动打开 `http://localhost:8501`

## ⚠️ 重要说明

### Word 转 PDF 依赖限制

**Windows 系统：**
- 需要安装 `docx2pdf` 库
- **必须安装 Microsoft Word**
- 转换效果完美，保持所有格式和图片

**macOS 系统：**
- 需要安装 `docx2pdf` 库
- 需要安装 Microsoft Word 或 LibreOffice

**Linux 系统：**
- 建议使用 LibreOffice 命令行转换
- 或安装其他 PDF 渲染引擎

如果未安装依赖，会显示详细的错误提示和解决方案。

### PDF 转 Word 功能限制

⚠️ **当前实现为基础的文本和图片提取：**

- ✅ 能够识别并提取 PDF 中的文字内容
- ✅ 能够识别并提取 PDF 中的图片
- ❌ **无法保留原始排版格式**（如分栏、表格、特殊字体等）
- ❌ **无法保留复杂的页面布局**

**建议：** PDF 转 Word 后，需要人工进行排版调整和格式优化。这是技术限制，大多数 PDF 转换工具都面临同样的问题。

### 图片丢失警告

当含有图片的文档转换为 TXT 格式时，系统会自动检测并提醒图片将会丢失，建议转换为 Word 或 PDF 格式以保留图片。

## 📦 技术栈

- **前端框架**: Streamlit
- **PDF 处理**: PyMuPDF (fitz)
- **Word 处理**: python-docx
- **PDF 生成**: reportlab
- **Word 转 PDF**: docx2pdf
- **AI 平台**: 硅基流动 (SiliconFlow)
- **多模态模型**: Qwen3-VL-32B-Instruct

## 🤝 使用建议

1. **简单文档转换** - 直接使用，效果良好
2. **复杂排版文档** - 转换后需要人工调整格式
3. **含图片文档** - 建议转为 Word 或 PDF，避免转为 TXT
4. **批量处理** - 可以多次上传文件进行转换

📝 许可证
本项目仅供学习和个人使用。