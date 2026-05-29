# 技术设计文档

## 架构概览

```
┌─────────────────────────────────┐
│        comment_app.py           │  ← tkinter GUI（启动、停止、日志）
├─────────────────────────────────┤
│        comment_bot.py           │  ← 核心编排器
├──────────┬──────────┬───────────┤
│ExcelHandler│BrowserAuto│AIClient │  ← 三大功能模块
│(openpyxl) │(Playwright)│(requests)│
├──────────┴──────────┴───────────┤
│     config.json / progress.json  │  ← 配置与状态持久化
└─────────────────────────────────┘
```

## 登录流程（已验证）

cp.baidu.com 使用百度通行证（passport.baidu.com）进行身份认证。

1. 访问任意帖子页面（`cp.baidu.com/forum/p?id=xxx`）
2. 点击 `.input-box`（"说说你的想法"）触发登录弹窗
3. 登录弹窗默认显示微信扫码，点击"账号登录"切换
4. 填写 `#TANGRAM__PSP_11__userName` 和 `#TANGRAM__PSP_11__password`
5. 勾选协议 `#TANGRAM__PSP_11__isAgree`
6. 点击 `#TANGRAM__PSP_11__submit` 提交
7. 成功后 URL 参数出现 `from_login=1`

## 评论发布流程（已验证）

1. 登录后重新加载帖子页面（清除弹窗残留）
2. 点击 `.input-box` → 打开评论弹窗 `.comment-publish-modal`
3. 在 `textarea.comment-input` 中填入评论文本
4. 用 JavaScript 触发发布按钮：`document.querySelector('.ant-modal .ant-btn-primary').click()`
   - **注意：** 普通 `page.click()` 无效，必须用 `page.evaluate()` JS 触发
5. 弹窗关闭 = 发布成功

## 帖子内容提取

- 标题：`.title` CSS 选择器
- 正文：`.content` CSS 选择器

## Excel 读取策略

### 帖子表格自动列识别

- URL 列：优先搜索表头含"URL"（不区分大小写），其次匹配 `cp.baidu.com/forum/p?id=` 格式
- 员工列：表头文字与账号文件中的员工名精确匹配
- 完成标记：检测"1"、"已"或任何非空值

### 账号文件格式

| 员工名 | 账号1_用户名 | 账号1_密码 | 账号2_用户名 | 账号2_密码 | ... | 账号5_用户名 | 账号5_密码 |
|---|---|---|---|---|---|---|---|

## 进度持久化

`progress.json` 记录四维位置：`{sheet, row, employee, account_index}`

每完成一条评论后立即更新，支持精确到"某员工的第几个账号"的中断恢复。

## 错误处理

| 场景 | 策略 |
|---|---|
| 登录失败（密码错） | 跳过该账号，继续下一个 |
| 页面加载超时 | 重试 2 次，间隔 10 秒 |
| AI API 错误 | 重试 2 次，间隔 5 秒 |
| 评论发布失败 | 记录错误，继续下一个 |
| Excel 保存失败 | 保留 progress.json，提示用户 |
