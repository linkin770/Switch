# 办公文件跨格式本地化批量转换工具

> **版本：v1.0.20260512**

## 项目简介

本工具是一个基于Flask的Web应用，用于实现办公文件的跨格式批量转换。支持多种文件格式之间的转换，包括PDF、DOCX、Excel、PPT、TXT、Markdown、LaTeX、CSV等。

## 功能特性

- **多格式支持**：支持PDF、DOCX、Excel、PPT、TXT、Markdown、LaTeX、CSV等多种格式的转换
- **批量处理**：支持批量上传和转换多个文件
- **多线程处理**：使用ThreadPoolExecutor进行并行处理，提高转换效率
- **扫描PDF处理**：检测并处理扫描版PDF（无文字层）
- **跨平台支持**：在Windows上使用Office应用程序，在Linux上使用LibreOffice作为备选
- **美观界面**：采用玻璃态设计，具有现代感和良好的用户体验
- **拖放上传**：支持文件拖放上传功能
- **格式提示**：下拉菜单悬停提示，显示支持的输入格式
- **ZIP打包**：转换完成后自动打包为ZIP文件，方便下载

## 技术栈

- **后端**：Python 3.12+, Flask 3.x
- **前端**：HTML5, CSS3, JavaScript
- **转换引擎**：
  - PDF处理：PyMuPDF, pdf2docx
  - Office处理：python-docx, weasyprint, python-office (poword, poexcel, poppt)
  - 备选方案：LibreOffice, Windows Office
- **其他库**：pandas, markdown, fpdf2

## 安装与运行

### 依赖安装

```bash
# 安装Python依赖
pip install -r requirements.txt

# Linux系统还需要安装LibreOffice
sudo apt install libreoffice
```

### 运行应用

```bash
# 启动Flask应用
python app.py

# 或使用启动脚本（Linux）
bash start.sh
```

应用将运行在 http://127.0.0.1:5000/

## 项目结构

```
├── app.py              # 主应用文件
├── index.html          # 前端界面
├── requirements.txt    # 依赖文件
├── start.sh            # 启动脚本
├── static/             # 静态文件
│   ├── icon.png        # 应用图标
│   └── script.js       # 前端脚本
├── tools/              # 工具模块
│   ├── __init__.py     # 模块导入
│   ├── office_tools.py # Office文件转换
│   ├── pdf_tools.py    # PDF文件转换
│   └── zip_tools.py    # ZIP打包工具
├── tests/              # 测试文件
│   ├── test_basic_conversion.py    # 基本转换测试
│   ├── test_docx2pdf.py           # DOCX转PDF测试
│   ├── test_excel_conversion.py    # Excel转换测试
│   ├── test_full_flow.py           # 完整流程测试
│   ├── test_pdf_conversion.py      # PDF转换测试
│   └── test_ppt_conversion.py      # PPT转换测试
└── config.py           # 配置文件
```

## 配置管理

配置文件 `config.py` 用于管理不同环境的配置参数，包括：
- 文件存储路径
- 转换参数
- 线程池配置
- 其他应用设置

## 版本控制

### 版本历史

| 版本 | 日期 | 主要变更 |
|------|------|----------|
| v1.0.20260512 | 2026-05-12 | 升级依赖兼容Python 3.12，重构转换回退链，添加类型注解，修复安全与类型问题 |
| v1.0.0 | 2026-04-05 | 初始版本，实现基本文件转换功能 |

### 发布说明

- **v1.0.20260512**：
  - 升级 Flask 至 3.x，所有依赖兼容 Python 3.12
  - 重构 Office 转换回退链，修复 comtypes/win32com 嵌套 bug，Excel/PPT 转换不再依赖 poexcel
  - 为所有工具函数添加类型注解，修复 basedpyright 类型警告
  - 修复 config.py `os.cpu_count()` 返回 None 的类型问题
  - 修复 app.py `file.filename` 可能为 None 的问题
  - 添加 `.gitignore` 和 `pyrightconfig.json`
  - 从 git 追踪中移除 `__pycache__/`、`uploads/` 等运行时文件

- **v1.0.0**：
  - 实现了PDF、DOCX、Excel、PPT、TXT、Markdown、LaTeX、CSV等格式的转换
  - 支持批量处理和多线程并行转换
  - 实现了扫描PDF的检测和处理
  - 提供了美观的玻璃态界面
  - 支持文件拖放上传和格式提示
  - 实现了ZIP打包下载功能

## 团队协作

### 开发流程

1. 从主分支创建功能分支
2. 开发新功能或修复bug
3. 编写测试用例
4. 提交代码并创建Pull Request
5. 代码审查后合并到主分支

### 代码规范

- 遵循PEP 8代码风格
- 函数和变量命名清晰，使用英文
- 关键代码添加注释
- 测试用例覆盖主要功能

## 常见问题

### 转换失败

- 检查文件是否损坏
- 确保所需的依赖库已安装
- Linux系统确保已安装LibreOffice
- 扫描版PDF可能无法提取文字，会转换为图像

### 性能问题

- 大型文件转换可能需要较长时间
- 建议分批转换大量文件
- 确保系统有足够的内存和CPU资源

## 联系与支持

如有问题或建议，请联系项目维护人员。