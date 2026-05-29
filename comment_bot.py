#!/usr/bin/env python3
"""
自动评论机 — 核心引擎
自动登录虚拟账号，在捣谷社区帖子下生成并发布评论。
"""

import json
import os
import random
import re
import time
import traceback
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# ---------- 路径 ----------
BASE_DIR = Path(__file__).parent
CONFIG_PATH = BASE_DIR / "config.json"
PROGRESS_PATH = BASE_DIR / "progress.json"


# ============================================================
# 1. 配置管理
# ============================================================
class ConfigManager:
    def __init__(self, config_path=CONFIG_PATH):
        self.config_path = config_path
        self.data = self._load()

    def _load(self):
        if self.config_path.exists():
            with open(self.config_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def save(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value
        self.save()


# ============================================================
# 2. 账号管理
# ============================================================
class AccountManager:
    """从 Excel 读取员工→账号映射"""

    def __init__(self, accounts_path):
        self.accounts_path = accounts_path
        self.accounts = {}  # {员工名: [{username, password}, ...]}
        self._load()

    def _load(self):
        wb = openpyxl.load_workbook(self.accounts_path, data_only=True)
        ws = wb.active
        # 读表头，检测格式
        headers = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, ws.max_column + 1)]
        # 简单格式：2列（账号名, 密码），所有账号归属一个员工
        if len(headers) >= 2 and "账号名" in headers[0] and "密码" in headers[1] and "员工名" not in str(headers):
            acc_list = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                username = str(row[0]).strip() if row[0] else ""
                password = str(row[1]).strip() if len(row) > 1 and row[1] else ""
                if username and password:
                    acc_list.append({"username": username, "password": password})
            # 用文件名推断员工名
            import re
            stem = Path(self.accounts_path).stem
            emp_match = re.search(r"(欣怡|瑞婧|税越|小笛|冉静|珽何)", stem)
            emp_name = emp_match.group(1) if emp_match else "默认"
            if acc_list:
                self.accounts[emp_name] = acc_list
        else:
            # 多员工格式：A列员工名，后续每两列一组
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue
                name = str(row[0]).strip()
                acc_list = []
                for i in range(1, min(len(row), 11), 2):
                    username = row[i]
                    password = row[i + 1] if i + 1 < len(row) else None
                    if username and password:
                        acc_list.append({
                            "username": str(username).strip(),
                            "password": str(password).strip(),
                        })
                if acc_list:
                    self.accounts[name] = acc_list
        wb.close()

    def get_employee_names(self):
        return list(self.accounts.keys())

    def get_accounts(self, employee_name):
        return self.accounts.get(employee_name, [])


# ============================================================
# 3. Excel 读写
# ============================================================
class ExcelHandler:
    """读帖子表格、写完成标记"""

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.wb = None

    def open(self):
        self.wb = openpyxl.load_workbook(self.excel_path)

    def save(self):
        if self.wb:
            self.wb.save(self.excel_path)

    def close(self):
        if self.wb:
            self.wb.close()
            self.wb = None

    def get_sheet_names(self):
        return self.wb.sheetnames

    def _find_url_column(self, ws):
        """自动识别 URL 列"""
        # 方法1：搜索表头含"URL"（不区分大小写）
        for col in range(1, ws.max_column + 1):
            header = str(ws.cell(row=1, column=col).value or "").strip()
            if "url" in header.lower():
                return col
        # 方法2：搜索第2行中含 cp.baidu.com 链接的单元格
        for col in range(1, ws.max_column + 1):
            val = str(ws.cell(row=2, column=col).value or "")
            if "cp.baidu.com/forum/p?id=" in val:
                return col
        return None

    def _find_employee_columns(self, ws, employee_names):
        """自动识别员工列：{员工名: 列号}"""
        emp_cols = {}
        headers = []
        for col in range(1, ws.max_column + 1):
            h = str(ws.cell(row=1, column=col).value or "").strip()
            headers.append((col, h))
        for col, h in headers:
            if h in employee_names:
                emp_cols[h] = col
        return emp_cols

    def is_completed(self, ws, row, col):
        """判断单元格是否已标记完成"""
        val = str(ws.cell(row=row, column=col).value or "").strip()
        return val in ("1", "已")

    def mark_completed(self, ws, row, col):
        """标记完成"""
        ws.cell(row=row, column=col).value = "1"

    def get_posts(self, sheet_name, employee_names, progress_state=None):
        """
        获取某个 Sheet 中需要处理的帖子列表。
        progress_state: 用于断点续跑的位置信息。

        返回: [(row_num, url, {员工名: col_index}), ...]
        """
        ws = self.wb[sheet_name]
        url_col = self._find_url_column(ws)
        if url_col is None:
            return []

        emp_cols = self._find_employee_columns(ws, employee_names)
        if not emp_cols:
            return []

        posts = []
        for row in range(2, ws.max_row + 1):
            url_val = str(ws.cell(row=row, column=url_col).value or "").strip()
            if not url_val.startswith("https://cp.baidu.com/forum/p?id="):
                continue

            # 检查哪些员工还没完成
            pending_employees = {}
            for emp_name, col_idx in emp_cols.items():
                if not self.is_completed(ws, row, col_idx):
                    pending_employees[emp_name] = col_idx

            if pending_employees:
                posts.append((row, url_val, pending_employees))

        # 断点续跑过滤
        if progress_state and progress_state.get("sheet") == sheet_name:
            start_row = progress_state.get("row", 0)
            posts = [p for p in posts if p[0] >= start_row]

        return posts


# ============================================================
# 4. 进度管理
# ============================================================
class ProgressManager:
    def __init__(self, progress_path=PROGRESS_PATH):
        self.progress_path = progress_path
        self.state = self._load()

    def _load(self):
        if self.progress_path.exists():
            with open(self.progress_path, "r", encoding="utf-8") as f:
                return json.load(f)
        return {}

    def save(self):
        with open(self.progress_path, "w", encoding="utf-8") as f:
            json.dump(self.state, f, ensure_ascii=False, indent=2)

    def update(self, sheet, row, employee, account_index):
        self.state.update({
            "sheet": sheet,
            "row": row,
            "employee": employee,
            "account_index": account_index,
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        })
        self.save()

    def clear(self):
        self.state = {}
        if self.progress_path.exists():
            self.progress_path.unlink()

    def get_resume_point(self):
        if not self.state:
            return None
        return {
            "sheet": self.state.get("sheet"),
            "row": self.state.get("row", 0),
            "employee": self.state.get("employee"),
            "account_index": self.state.get("account_index", 0),
        }


# ============================================================
# 5. AI 评论生成
# ============================================================
class AIClient:
    """调用 OpenAI 兼容 API 生成评论"""

    # 5种完全不同的评论风格，每个账号固定一种
    STYLES = [
        {
            "name": "简单直夸型",
            "instruction": "你说话简短自然，像朋友圈随手评论。不要用颜文字或字母表情。遇到长篇文字帖，优先夸文笔或催更：\"老师写得好棒\"\"高质高产给我火\"。遇到二次元IP表达同好，不点名作品。遇到真人/明星只夸画面。不确定角色时直接夸画面/氛围/文笔，不追问角色名。",
            "examples": [
                "好甜",
                "这也太会写了吧",
                "老师文笔好好",
                "老师写得好棒啊，高质高产给我火！",
                "太权威了",
                "哇老师你也玩这个游戏呀",
                "写得真好",
                "这口糖我吃了",
            ],
        },
        {
            "name": "激动感叹型",
            "instruction": "情绪激动，用感叹号或语气词开头（啊啊啊/我去/不是/妈呀/天哪）。不要用TAT、qwq、233。遇到长篇文字帖可以激动催更：\"老师快快更新哇写得太好了\"。遇到二次元IP可以表达意外和兴奋，但不点名作品。遇到不认识的角色直接对画面/文笔表达激动。",
            "examples": [
                "啊啊啊甜死我了！！",
                "我去这也太绝了",
                "不是……这谁能顶得住啊",
                "妈呀这里也能遇到同好！！",
                "天哪太会写了吧",
                "啊啊啊啊藕饼99！！",
                "老师快快更新哇，真的写的好好！入住在评论区了",
                "天老爷这是我能看的吗",
            ],
        },
        {
            "name": "引用细节型",
            "instruction": "你读得认真，会提到文中或图中某个真正有记忆点的细节来夸。注意：只挑真正有意思的细节（金句、反转、名场面），不要随便挑两个平凡词汇硬夸——\"你的手\"这种普通词不值得被引用。直接夸那个细节，不要用\"[文字]配[画面]\"这种句式。不确定时夸视觉元素（色调、构图、氛围）就好。",
            "examples": [
                "混天绫打死结……这是我们能看的吗",
                "\"我在\"这两个字给我看哭了",
                "尾巴冷要抱抱这句谁顶得住",
                "这张画得也太有感觉了",
                "递姜茶那段写得也太细了",
                "改词致歉……好有画面感",
                "这颗红耳饰好特别",
                "蓝白配色也太会了吧",
            ],
        },
        {
            "name": "提问互动型",
            "instruction": "问作者一个问题或表达好奇，像跟朋友聊天。可以问创作相关（怎么画的/灵感来源/后续计划）。遇到长篇小说帖，优先催更：\"老师你的草稿箱里一定还有10000字对不对\"\"后面还有吗老师\"。遇到认识的二次元IP问\"老师也玩这个呀\"\"你也看这个呀\"，不说具体作品名。遇到不认识的画面可以问\"老师画的是哪个角色呀\"。注意：只有你才适合问角色名，其他四个风格不问。",
            "examples": [
                "作者是怎么写出来的呀",
                "老师你的草稿箱里一定还有10000字对不对？赶紧放出来让我看看",
                "老师扩列吗",
                "这个梗是哪来的求指路",
                "老师也玩这个游戏吗",
                "什么时候更新下一章",
                "后面还有吗老师，看不够",
            ],
        },
        {
            "name": "俏皮调侃型",
            "instruction": "语气调皮，像跟好朋友开玩笑，会拿内容调侃。偶尔用233或省略号，但整条评论只用一处。遇到长篇小说帖可以催更：\"更新求踢\"\"入住评论区了\"\"等我睡醒你还没更新我就再睡一次\"。遇到二次元IP可以感叹遇到同好。遇到不认识的画面走幽默路线，不追问角色。",
            "examples": [
                "哇塞小情侣真会玩",
                "头七还在夸好美笑死我了",
                "更新求踢",
                "哪吒你嘴硬手软还挺会啊",
                "入住在评论区了",
                "白楚年你搁这儿养鱼呢",
                "这哪是糖葫芦这是狗粮吧",
                "老师我们来玩一个小游戏吧，等我睡醒一次你就更新一篇……我醒了",
            ],
        },
    ]

    BANNED_PATTERNS = [
        "先收藏了", "想追后续", "会持续关注", "期待更新",
        "已关注", "求更新", "期待更多", "会一直支持",
        "值得推荐", "强烈推荐", "受益匪浅",
        "TAT", "qwq", "QwQ", "QWQ", "tat",
        # 严禁质疑缺内容、缺图片
        "图片", "没图", "看不到", "没看到", "没发图",
        "怎么判断", "咋判断", "看不出", "只有标题",
        "只有两字", "就两字", "内容空", "没内容",
        "能不能发", "有没有图", "配张图", "发个图",
        "是哪种风格", "什么风格", "啥风格",
    ]

    def __init__(self, config):
        self.api_url = config.get("api_url")
        self.api_key = config.get("api_key")
        self.api_model = config.get("api_model", "gpt-5.4")

    def generate(self, title, content, account_index=0, image_description="", confidence="中"):
        style = self.STYLES[account_index % len(self.STYLES)]
        banned = "、".join(self.BANNED_PATTERNS)

        img_hint = ""
        if image_description:
            img_hint = f"\n帖子配图信息：{image_description}"

        # 根据置信度决定IP表达策略
        if confidence == "高":
            ip_rule = "识别置信度高，可以叫角色昵称，也可以提作品名（如\"原神\"这种大IP），像粉丝之间自然聊天一样"
        elif confidence == "中":
            ip_rule = "识别置信度中等，可以叫角色昵称，但不要提作品全名，用\"这个游戏\"\"这个动漫\"等模糊说法"
        else:
            ip_rule = "识别置信度低，不要提任何角色名或作品名，直接夸画面和文笔就好"

        prompt = f"""你是捣谷社区的活跃用户，正在看一篇帖子。

帖子标题：{title}

帖子正文：
{content[:2000]}{img_hint}

写一条评论，要求：
- 你属于「{style["name"]}」
- {style["instruction"]}
- 5-25字，简短口语化，像真人随口说的
- 结合正文、配图、角色信息一起评论
- 如果正文很长（超过200字），这是长篇小说帖，优先走催更/鼓励/入住评论区路线，不要硬分析细节或逐句引用
- 引用细节时只挑真正有记忆点的金句名场面，不要随便挑平凡词汇硬夸（\"你的手\"这种普通词不值得引）
- {ip_rule}
- 二次元IP → 表达同好；真人/明星 → 只夸画面，不要假装认识明星本人
- 不要用\"[文字]配[画面]\"这种句式，真人不会这么说话
- 遇到不认识的角色：只有「提问互动型」适合问角色名，其他风格直接夸画面/氛围/文笔即可
- 严禁说\"加个好友\"或任何交友/加好友相关的话
- 严禁使用以下词语：{banned}
- 只输出评论本身，一行，不要引号不要解释

参考风格（不要照抄，根据帖子内容灵活变通）：
{chr(10).join(f'- "{ex}"' for ex in style["examples"])}

评论："""

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.api_model,
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 80,
            "temperature": 0.95,
            "top_p": 0.95,
            "frequency_penalty": 0.5,
            "presence_penalty": 0.3,
        }

        try:
            resp = requests.post(
                self.api_url, headers=headers, json=payload, timeout=90
            )
            resp.raise_for_status()
            result = resp.json()
            comment = result["choices"][0]["message"]["content"].strip()
            # 去掉可能的引号包裹
            comment = comment.strip("\"'\"")
            return comment
        except Exception as e:
            raise RuntimeError(f"AI API 调用失败: {e}")

    def describe_image(self, image_base64, title="", content=""):
        """用视觉模型描述帖子中的图片（结合标题正文识别角色和IP，含置信度）"""
        context = ""
        if title or content:
            context = f"\n补充信息：帖子标题为「{title}」，帖子正文前200字为「{content[:200]}」。请结合这些文字信息判断图中内容属于哪个作品/IP。"

        payload = {
            "model": self.api_model,
            "messages": [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": f"请完成以下任务：\n1. 简要描述图片内容（人物/场景/风格）\n2. 识别图中人物/作品的IP来源，并标注你的把握程度\n   - 高：角色特征极其明显+标题正文佐证（如芙宁娜、钟离、魈等辨识度极高的角色）\n   - 中：有较明显特征但可能存在混淆（如两个作品画风接近）\n   - 低：只能猜测，不确定\n\n{context}\n\n用中文回复，格式：\n【画面】xxx\n【角色】角色名（如有）\n【IP】作品名\n【置信度】高/中/低"},
                        {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_base64}"}},
                    ],
                }
            ],
            "max_tokens": 200,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        try:
            resp = requests.post(self.api_url, headers=headers, json=payload, timeout=90)
            resp.raise_for_status()
            return resp.json()["choices"][0]["message"]["content"].strip()
        except Exception:
            return ""  # 识图失败不阻塞流程


# ============================================================
# 6. 浏览器自动化
# ============================================================
class BrowserAutomator:
    """Playwright 浏览器控制"""

    def __init__(self, headless=True):
        self.headless = headless
        self.playwright = None
        self.browser = None

    def start(self):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(headless=self.headless)

    def stop(self):
        if self.browser:
            self.browser.close()
        if self.playwright:
            self.playwright.stop()

    def login(self, username, password, post_url=None):
        """登录 cp.baidu.com，返回带登录态的 page"""
        page = self.browser.new_page()
        try:
            if post_url:
                page.goto(post_url, timeout=30000)
            else:
                page.goto("https://cp.baidu.com/forum/", timeout=30000)
            page.wait_for_timeout(5000)

            # 用 JS 点击评论区触发登录弹窗（绕过可见性检查）
            page.evaluate("document.querySelector('.input-box').click()")
            page.wait_for_timeout(6000)

            # 切换到账号登录
            page.click("text=账号登录")
            page.wait_for_timeout(3000)

            # 填写账号密码
            page.fill("#TANGRAM__PSP_11__userName", username)
            page.fill("#TANGRAM__PSP_11__password", password)
            try:
                page.check("#TANGRAM__PSP_11__isAgree")
            except Exception:
                pass

            page.click("#TANGRAM__PSP_11__submit")
            page.wait_for_timeout(12000)

            # 重载页面清除登录弹窗残留
            page.goto(page.url, timeout=30000)
            page.wait_for_timeout(3000)

            return page
        except Exception:
            page.close()
            raise

    def extract_post(self, page, url):
        """打开帖子并提取标题、内容和截图"""
        page.goto(url, timeout=30000)
        page.wait_for_timeout(5000)

        title = ""
        content = ""
        screenshot_b64 = ""

        title_el = page.query_selector(".title")
        if title_el:
            title = title_el.inner_text().strip()

        content_el = page.query_selector(".content")
        if content_el:
            content = content_el.inner_text().strip()

        # 截取帖子内容区域（包含图片）
        try:
            post_el = page.query_selector(".post-detail, .post-content, .article-content, .main")
            if post_el:
                img_bytes = post_el.screenshot()
            else:
                img_bytes = page.screenshot()
            import base64
            screenshot_b64 = base64.b64encode(img_bytes).decode()
        except Exception:
            screenshot_b64 = ""

        return title, content, screenshot_b64

    def post_comment(self, page, comment_text):
        """在当前帖子页面发布评论"""
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(1000)
        except Exception:
            pass

        # 用 JS 点击打开评论弹窗（绕过可见性检查）
        page.evaluate("document.querySelector('.input-box').click()")
        page.wait_for_timeout(3000)

        # 输入评论
        page.fill("textarea.comment-input", comment_text)
        page.wait_for_timeout(2000)

        # 用 JS 触发发布按钮点击
        page.evaluate(
            "document.querySelector('.ant-modal .ant-btn-primary').click()"
        )
        page.wait_for_timeout(6000)

    def already_commented(self, page, username):
        """检查当前账号是否已经在这个帖子下发过评论"""
        try:
            # 等评论区加载
            page.wait_for_timeout(2000)
            comments = page.query_selector_all('.comment-item-wrap')
            for item in comments:
                text = item.inner_text()
                if username in text:
                    return True
            return False
        except Exception:
            return False

    def close_page(self, page):
        try:
            page.close()
        except Exception:
            pass


# ============================================================
# 7. 核心编排器
# ============================================================
class CommentBot:
    def __init__(self, config, log_callback=None, progress_callback=None):
        self.config = config
        self.log = log_callback or print
        self.progress_cb = progress_callback or (lambda msg: None)

        self.excel = ExcelHandler(config.get("last_post_file"))
        self.ai = AIClient(config)
        self.browser = BrowserAutomator(headless=config.get("headless", True))
        self.progress = ProgressManager()
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self, accounts_path, employee_filter=None, sheet_filter=None,
            min_row=None, max_row=None):
        """主执行流程（账号优先：一个账号登录一次，刷完所有帖子再换号）"""
        # 加载账号
        self.log("正在加载账号信息...")
        account_mgr = AccountManager(accounts_path)
        employee_names = account_mgr.get_employee_names()
        if not employee_names:
            self.log("错误：未找到任何账号信息")
            return
        self.log(f"找到 {len(employee_names)} 名员工的账号：{', '.join(employee_names)}")
        if employee_filter:
            if employee_filter not in employee_names:
                self.log(f"错误：未找到员工「{employee_filter}」")
                return
            employee_names = [employee_filter]

        # 打开 Excel
        self.log("正在加载帖子表格...")
        self.excel.open()
        sheets = self.excel.get_sheet_names()
        self.log(f"共 {len(sheets)} 个 Sheet")

        # 启动浏览器
        self.log("正在启动浏览器...")
        self.browser.start()

        total_done = 0
        total_errors = 0

        try:
            for sheet_name in sheets:
                if sheet_filter and sheet_name != sheet_filter:
                    continue
                if self._cancelled:
                    break

                self.log(f"\n--- 开始处理 Sheet: {sheet_name} ---")
                self.progress_cb(f"Sheet: {sheet_name}")

                posts = self.excel.get_posts(sheet_name, employee_names, progress_state=None)
                if not posts:
                    self.log("该 Sheet 没有待处理的帖子")
                    continue

                # 行范围过滤
                if min_row:
                    posts = [(r, u, e) for r, u, e in posts if r >= min_row]
                if max_row:
                    posts = [(r, u, e) for r, u, e in posts if r <= max_row]
                if not posts:
                    self.log("行范围过滤后没有待处理的帖子")
                    continue

                # Phase 1: 提取所有帖子内容（用第一个账号登录一次）
                self.log(f"正在提取 {len(posts)} 个帖子的内容...")
                post_data = {}  # row_num -> {title, content, screenshot, image_desc, confidence}
                first_acc = None
                for emp_name in employee_names:
                    accs = account_mgr.get_accounts(emp_name)
                    if accs:
                        first_acc = accs[0]
                        break

                if first_acc:
                    page = None
                    for row_num, url, pending_emps in posts:
                        if self._cancelled:
                            break
                        try:
                            page = self.browser.login(
                                first_acc["username"], first_acc["password"],
                                post_url=url,
                            )
                            title, post_content, screenshot_b64 = self.browser.extract_post(page, url)
                            image_desc = ""
                            confidence = "中"
                            if screenshot_b64:
                                try:
                                    image_desc = self.ai.describe_image(screenshot_b64, title, post_content)
                                    if image_desc:
                                        if "置信度】高" in image_desc or "置信度：高" in image_desc:
                                            confidence = "高"
                                        elif "置信度】低" in image_desc or "置信度：低" in image_desc:
                                            confidence = "低"
                                except Exception:
                                    pass
                            post_data[row_num] = {
                                "title": title, "content": post_content,
                                "image_desc": image_desc, "confidence": confidence,
                            }
                            self.log(f"  行{row_num}: {title[:40]}")
                        except Exception as e:
                            self.log(f"  行{row_num}: 提取失败 - {e}")
                            post_data[row_num] = {
                                "title": "", "content": "",
                                "image_desc": "", "confidence": "中",
                            }
                        finally:
                            if page:
                                self.browser.close_page(page)
                                page = None

                # Phase 2: 每个账号登录一次，刷完所有帖子
                for emp_name in employee_names:
                    if self._cancelled:
                        break
                    accounts = account_mgr.get_accounts(emp_name)
                    if not accounts:
                        continue

                    # 记录每个帖子已完成评论的账号数
                    done_per_post = {}

                    for acc_idx, acc in enumerate(accounts):
                        if self._cancelled:
                            break
                        self.log(f"\n  {emp_name} 账号{acc_idx+1}/{len(accounts)}: {acc['username']}")

                        # 登录一次
                        page = None
                        try:
                            page = self.browser.login(
                                acc["username"], acc["password"],
                                post_url=posts[0][1],  # 用第一个帖子触发登录
                            )
                        except Exception as e:
                            self.log(f"    登录失败: {e}")
                            continue

                        # 遍历所有帖子
                        for row_num, url, pending_emps in posts:
                            if self._cancelled:
                                break
                            if row_num not in post_data:
                                continue
                            data = post_data[row_num]

                            try:
                                # 跳转到帖子页面
                                page.goto(url, timeout=30000)
                                page.wait_for_timeout(3000)

                                # 检查已评论
                                if self.browser.already_commented(page, acc["username"]):
                                    self.log(f"    行{row_num}: 已评论过，跳过")
                                    done_per_post[row_num] = done_per_post.get(row_num, 0) + 1
                                    continue

                                # 生成评论
                                comment = self.ai.generate(
                                    data["title"], data["content"], acc_idx,
                                    data["image_desc"], data["confidence"],
                                )
                                self.log(f"    行{row_num} \"{comment}\"")

                                # 发布
                                self.browser.post_comment(page, comment)
                                self.log(f"      ✓ 发布成功")
                                total_done += 1
                                done_per_post[row_num] = done_per_post.get(row_num, 0) + 1

                                # 保存进度
                                self.progress.update(sheet_name, row_num, emp_name, acc_idx)

                                # 随机延迟
                                delay = random.uniform(
                                    self.config.get("delay_min", 1),
                                    self.config.get("delay_max", 3),
                                )
                                time.sleep(delay)

                            except Exception as e:
                                self.log(f"    行{row_num}: 失败 - {e}")
                                total_errors += 1
                                self.progress.update(sheet_name, row_num, emp_name, acc_idx)

                        # 登出
                        if page:
                            self.browser.close_page(page)
                            page = None

                    # 所有账号完成后，标记完成的行
                    for row_num, url, pending_emps in posts:
                        if done_per_post.get(row_num, 0) >= len(accounts):
                            try:
                                col_idx = pending_emps.get(emp_name)
                                if col_idx:
                                    self.excel.mark_completed(
                                        self.excel.wb[sheet_name], row_num, col_idx
                                    )
                                    self.log(f"  行{row_num}: 标记完成 -> 1")
                                self.log(f"  行{row_num}: 标记完成")
                            except Exception:
                                pass

                    # 保存 Excel
                    try:
                        self.excel.save()
                        self.log("  [Excel 已保存]")
                    except Exception:
                        pass

                self.log(f"\n--- Sheet {sheet_name} 处理完毕 ---")

                self.excel.save()
                self.log(f"--- Sheet {sheet_name} 处理完毕 ---")

        finally:
            # 最终保存
            try:
                self.excel.save()
            except Exception:
                pass
            self.excel.close()
            self.browser.stop()
            self.progress.clear()

        self.log(f"\n===== 全部完成 =====")
        self.log(f"成功评论: {total_done} 条")
        self.log(f"失败: {total_errors} 条")
        self.progress_cb(f"完成！成功{total_done}条，失败{total_errors}条")


# ============================================================
# 命令行入口
# ============================================================
if __name__ == "__main__":
    import sys

    config = ConfigManager()
    # 支持命令行参数指定账号文件
    accounts_file = sys.argv[1] if len(sys.argv) > 1 else "账号信息.xlsx"
    accounts_path = BASE_DIR / accounts_file

    if not accounts_path.exists():
        print(f"错误：找不到账号文件 {accounts_path}")
        sys.exit(1)

    bot = CommentBot(config)
    bot.run(str(accounts_path))
