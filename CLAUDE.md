# 捣谷社区自动评论机 — 开发指引

## 项目概述

Mac 桌面工具，为捣谷社区（cp.baidu.com/forum）运营人员提供全自动化评论发布功能。批量使用多虚拟账号，自动读取帖子链接，AI 生成消费引导型评论并发布，最后在 Excel 表格中标记完成状态。

## 文件结构

```
自动评论机/
├── comment_bot.py          # 核心引擎（无界面可单独调用）
├── comment_app.py          # tkinter 图形界面
├── config.json             # 配置文件（API Key、选择器、延迟等）
├── 账号信息.xlsx           # 员工→虚拟账号映射（用户提供）
├── 捣谷社区每日打标帖.xlsx  # 帖子表格（用户提供）
├── progress.json           # 运行进度（断点续跑）
├── setup.command           # 首次安装脚本
├── 自动评论机.command       # 主启动脚本
├── docs/                   # 项目文档
│   ├── requirements.md     # 需求文档
│   ├── design.md           # 技术设计
│   └── implementation.md   # 执行步骤
└── logs/                   # 开发日志（按日期）
    └── 2026-05-28.md
```

## 开发原则

- **非技术用户友好**：图形界面简洁直观，双击即可运行
- **安全稳定**：每步保存进度，支持断点续跑，不怕断网/崩溃
- **防检测**：随机延迟、模拟真人操作节奏
- **最小变更**：不引入不必要的依赖和复杂度

## 技术栈

| 层 | 选项 |
|---|---|
| 语言 | Python 3.9+ |
| 浏览器自动化 | Playwright (Chromium) |
| Excel | openpyxl |
| AI | requests → OpenAI 兼容 API |
| GUI | tkinter (Mac 内置) |
