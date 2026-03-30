# DeepFile Organizer

一个基于 Python 的专业文件处理工具，提供文件批量提取、PDF 智能重命名和高级文件整理功能。

## 界面特色

- **VS Code Dark+ 主题**: 专业深色界面，护眼且美观
- **Enterprise 布局**: 清晰的卡片式设计，操作直观
- **模态弹窗**: 居中显示，带遮罩防误触

## 功能简介

### Tab 1: 文件批量提取
- 选择源文件夹，按文件名模式（通配符）和后缀名筛选文件
- 支持多后缀名筛选（如 `pdf, doc, docx`）
- 复制到目标目录，支持冲突处理：
  - 选项1：添加来源文件夹名作为前缀
  - 选项2：自动添加时间戳重命名

### Tab 2: PDF 智能重命名
- 加载 Excel 文件，选择"关联数据列"
- 使用 **本地 OCR (RapidOCR)** 或 **火山引擎 AI OCR** 识别 PDF 内容
- 匹配 Excel 关联列值，按指定命名列重命名 PDF 文件
- 支持组合两列作为新文件名

### Tab 3: 高级文件整理 
- **多源目录管理**: 支持 1-10 个源目录同时整理
- **Excel 配置驱动**: 基于 Excel 数据自动创建目录结构
- **可视化目录编辑器**: 使用 CheckboxTreeview 直观编辑目录层级
- **智能文件分类**: 根据人名和关键字自动分类文件
- **灵活操作模式**: 支持复制或剪切文件

### Tab 4: 智能归档 (Smart Archiving) 🆕
- **双模式OCR识别**: 本地RapidOCR + 火山引擎LLM OCR兜底
- **批量OCR处理**: 先批量本地OCR，再统一LLM兜底，效率更高
- **智能文件匹配**: 文件名匹配 → OCR内容匹配 → 可选列二次匹配
- **OCR结果记录**: 
  - 本地OCR结果自动保存到 `0-OCR识别结果/本地OCR识别结果.xlsx`
  - LLM OCR结果自动保存到 `0-OCR识别结果/LLMOCR识别记录.xlsx`
- **OCR优化功能**:
  - 自动移除OCR空格（身份证号等连续数据匹配更精准）
  - 支持"识别文字量"限制（控制LLM输出长度）
  - 支持"补充提示词"（自定义LLM识别指令）
- **Token消耗追踪**: LLM OCR自动记录并汇总显示Token使用情况
- **未处理文件管理**: 未匹配文件移至 `0未处理文件夹` 并附带OCR记录

## 快速开始

### 方式一：直接运行 Python 源码
```bash
pip install -r requirements.txt
python DeepFile_Organizer.py
```

### 方式二：使用预编译 EXE（推荐）
1. 访问 [Releases](https://github.com/jameszong/DeepFile-Organizer/releases)
2. 下载最新版本的 `DeepFile_Organizer.exe`
3. **无需安装 Python**，Windows 7 x64 及以上系统可直接运行

## 自动构建

本项目使用 GitHub Actions 自动构建 Windows EXE：

### 构建触发方式
- **自动触发**: 每次推送到 `main` 分支时自动构建预发布版本
- **标签发布**: 推送 `v*` 标签（如 `v1.0.0`）自动创建正式 Release
- **手动触发**: 在 GitHub Actions 页面手动运行构建

### 构建产物
- **Artifacts**: 每次构建的 EXE 文件
- **Releases**: 正式版本和预发布版本

## 依赖说明

| 功能 | 依赖包 |
|------|--------|
| GUI 界面 | tkinter (Python 内置) |
| Excel 处理 | pandas, openpyxl |
| PDF 读取 | PyMuPDF (fitz) |
| 本地 OCR | rapidocr-onnxruntime |
| AI OCR | openai SDK (兼容火山引擎) |
| 树形控件 | ttkwidgets |

## 火山引擎 AI OCR 配置

如需使用高精度 AI OCR：
1. 在 Tab 2 的"AI 配置"区域填入 API Key 和模型 ID
2. 点击"测试连通性"验证配置
3. 选择"火山引擎 AI"识别模式

**默认模型 ID**: `doubao-seed-1-6-vision-250815`

## 系统要求

- **Windows**: Windows 7 SP1 x64 或更高版本
- **内存**: 建议 4GB+
- **存储**: OCR 模型首次运行时会自动下载（约 100MB）

## 使用技巧

### Tab 3 高级整理工作流
1. **添加源目录**: 点击"+ 添加目录"按钮添加 1-10 个源文件夹
2. **配置 Excel**: 选择包含债务人信息的 Excel 文件
3. **设计目录结构**: 
   - 点击"+ 添加目录层级"创建多级目录
   - 双击目录选择对应的 Excel 列
   - 右键或点击"+ 添加关键字"设置文件分类规则
4. **设置操作模式**: 选择复制或剪切文件
5. **开始整理**: 点击"开始整理文件"执行任务

### 文件匹配规则
- **人名匹配**: 使用 fnmatch 算法，支持多种模式（如 `张三_*`、`*_李四`）
- **关键字匹配**: 智能识别文件类型，支持模糊匹配
- **冲突处理**: 自动处理文件名冲突，确保数据完整性

## 更新日志

### v1.3.0 (2025-02-11)
-  新增 Tab 3: 高级文件整理功能
-  界面升级为 VS Code Dark+ 主题
-  优化 Tab 标签样式，采用 VS Code 风格
-  模态弹窗居中显示，添加遮罩防误触
-  更新 GitHub Actions 工作流，支持自动发布

### v1.2.0
-  新增 PDF 智能重命名功能
-  集成火山引擎 AI OCR
-  支持 Excel 数据驱动的文件重命名

### v1.1.0
-  新增文件批量提取功能
-  采用 Enterprise 风格界面设计

## 贡献

欢迎提交 Issue 和 Pull Request！

## 许可证

MIT License