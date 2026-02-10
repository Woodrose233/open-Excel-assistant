# Excel AI 助理 (Excel-AI-Assistant)

一个轻量级的开源 Excel AI 助手插件，旨在通过简单的 API Key 配置，为用户提供强大的数据分析和自动化能力。

## 🌟 项目愿景
- **开源免费**：作为 Microsoft Copilot 的开源替代方案。
- **隐私至上**：API Key 本地存储，支持连接本地模型（如 Ollama）。
- **轻量易用**：基于 Office JS 现代架构，无需安装复杂的本地环境。

## 🛠️ 技术栈
- **前端**: React + TypeScript + Fluent UI
- **插件架构**: Office Add-in (Web-based)
- **AI 集成**: 支持 OpenAI, DeepSeek, 以及任何兼容 OpenAI 格式的本地/云端 API。

## 🚀 快速开始

### 1. 克隆项目
```bash
git clone https://github.com/YourUsername/Excel-AI-Assistant.git
cd Excel-AI-Assistant/ExcelAIAssistant
```

### 2. 安装依赖
```bash
npm install
```

### 3. 启动项目
```bash
npm start
```

## 📝 路线图
- [ ] API Key 配置界面实现
- [ ] 基础聊天与单元格数据抓取
- [ ] AI 指令解析与 Excel 自动执行引擎
- [ ] 多模型切换支持 (OpenAI/DeepSeek/Ollama)

## 📄 开源协议
MIT License
