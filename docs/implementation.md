# 实施步骤

## 阶段一：环境验证（已完成 ✓）

- [x] 验证 cp.baidu.com 可访问性
- [x] 确认登录流程和选择器
- [x] 确认评论发布流程和选择器
- [x] 验证 AI API 可用性
- [x] 使用测试账号完整跑通一次

## 阶段二：核心功能开发（当前）

- [x] 创建账号 Excel 模板
- [x] 创建 config.json 配置文件
- [x] 实现 comment_bot.py 核心引擎
- [x] 实现 comment_app.py 图形界面
- [x] 创建启动脚本
- [ ] 端到端测试：完整流程验证（修复中）
- [ ] Bug 修复和优化

## 阶段三：完善与交付

- [ ] 编写使用指南（使用指南.md）
- [ ] 用户验收测试（用用户提供的真实账号测试）
- [ ] 根据反馈优化
- [ ] 最终交付

## 依赖清单

```
# Python 包
openpyxl   # Excel 读写
playwright # 浏览器自动化
requests   # HTTP 请求（AI API）

# 系统
Python 3.9+ (Mac 内置)
Playwright Chromium 浏览器
```

## 如何开发/修改

1. 核心逻辑在 `comment_bot.py`，修改后可直接 `python3 comment_bot.py` 测试
2. 界面在 `comment_app.py`，修改后运行 `python3 comment_app.py` 查看效果
3. 网页选择器如变更，修改 `comment_bot.py` 中 `BrowserAutomator` 类的方法
4. AI 提示词调整，修改 `comment_bot.py` 中 `AIClient.generate()` 方法

## 注意事项

- 发布按钮必须用 JS evaluate 触发，普通 click 无效
- 登录后建议刷新页面，否则可能有弹窗残留影响操作
- 账号切换间隔建议 >= 3 秒，避免触发风控
- Excel 保存频率不宜过高（默认每 5 行），避免频繁 IO
