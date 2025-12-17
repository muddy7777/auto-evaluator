"""把 main.ipynb 的流程合并到一个可运行的脚本里。

依赖：见 requirements.txt

运行：
1) 在环境变量或 .env 中设置：AI_API_KEY（必需）
   可选：AI_BASE_URL / HOMEWORK_URL / MODEL_NAME
2) 执行：python main.py
3) 浏览器打开后手动登录，回到终端按回车继续。

行为与 Notebook 对齐：
- 下载：点击 field_5 打开详情弹窗，滚动到底部找下载按钮，只下 .cpp；点击后固定等待 2s，再判断下载完成（兼容 .tmp/.crdownload）
- 回填：AntD 弹窗里点“修改”→点“请选择”→在 listbox(role=option) 里点分数→“提交”→右上角 Close
- 批量：field_11（教师评分）已有数字则跳过
- 性能：关闭 implicit wait，避免与显式等待叠加导致回填很慢
"""

from __future__ import annotations

import glob
import os
import re
import time

from dotenv import load_dotenv
from openai import OpenAI
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    StaleElementReferenceException,
    TimeoutException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


load_dotenv()


HOMEWORK_URL = os.getenv("HOMEWORK_URL") or "https://next.jinshuju.net/forms/NHaQvT/entries"
MODEL_NAME = os.getenv("MODEL_NAME") or "gpt-5-mini"

# 只从环境变量读取（不允许硬编码）
API_KEY = os.getenv("AI_API_KEY")
BASE_URL = os.getenv("AI_BASE_URL") or "https://api.openai-proxy.org/v1"
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")


SCORING_CRITERIA = """
你是C++作业评分助教，给大一的学生批改c++作业，按以下标准评分（满分10分，平均分8分）,但不要太严格，在任意评分维度上做的很好即可打高分：
1. 代码逻辑正确性：是否符合作业需求，逻辑无漏洞；
2. 代码规范性：命名规范、缩进整齐、结构清晰；
3. 注释完整性：关键步骤有注释，便于理解；
4. 代码简洁性：无冗余代码，实现高效。
评分输出格式：
第一行：分数（仅数字，例如：8.5）
第二行：简短评语（例如：代码逻辑正确，命名规范，注释完整，建议优化循环结构以提升简洁性）
"""


def setup_driver():
    chrome_options = Options()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)

    service = Service(ChromeDriverManager().install())
    d = webdriver.Chrome(service=service, options=chrome_options)
    # 默认关闭 implicit wait，避免与显式等待叠加导致整体变慢。
    d.implicitly_wait(0)
    return d


def wait_for_grid(driver):
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.CLASS_NAME, "ag-root")))
    viewport = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.CLASS_NAME, "ag-body-viewport"))
    )
    return viewport


def get_visible_rows(driver):
    return driver.find_elements(
        By.XPATH,
        "//div[contains(@class, 'ag-center-cols-container')]//div[@role='row']",
    )


def clear_download_dir():
    for f in glob.glob(os.path.join(DOWNLOAD_DIR, "*")):
        try:
            os.remove(f)
        except Exception:
            pass


def wait_download_complete(timeout=60, poll_interval=0.5, settle_rounds=3):
    """等待下载完成。

    兼容两类临时文件：
    - Chrome 的 *.crdownload
    - 站点/浏览器可能出现的 *.tmp

    规则：
    - 忽略 *.crdownload / *.tmp
    - 找到最新的“非临时文件”后，要求文件大小连续 settle_rounds 次不变才认为完成
    """

    start = time.time()
    last_path = None
    stable_count = 0
    last_size = None

    while time.time() - start < timeout:
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*"))
        candidates = [
            p
            for p in files
            if not p.lower().endswith(".crdownload") and not p.lower().endswith(".tmp")
        ]

        if candidates:
            candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
            path = candidates[0]

            try:
                size = os.path.getsize(path)
            except OSError:
                time.sleep(poll_interval)
                continue

            if path == last_path and size == last_size:
                stable_count += 1
            else:
                stable_count = 0
                last_path = path
                last_size = size

            if stable_count >= settle_rounds:
                return path

        time.sleep(poll_interval)

    return None


def _extract_filename_from_href(href: str) -> str:
    if not href:
        return ""
    m = re.search(r"(?:\\?|&)(?:attname)=([^&]+)", href)
    if not m:
        return ""
    try:
        from urllib.parse import unquote

        return unquote(m.group(1))
    except Exception:
        return m.group(1)


def _contains_cpp_hint(s: str) -> bool:
    if not s:
        return False
    return bool(re.search(r"(?i)\\.cpp(\\b|$)", s))


def _is_cpp_download_link(a) -> bool:
    href = (a.get_attribute("href") or "").strip()
    dl = (a.get_attribute("download") or "").strip()
    title = (a.get_attribute("title") or "").strip()
    aria = (a.get_attribute("aria-label") or "").strip()
    text = (a.text or "").strip()
    attname = _extract_filename_from_href(href).strip()

    candidates = [dl, attname, title, aria, text, href]
    return any(_contains_cpp_hint(c) for c in candidates)


def _get_top_visible_ant_modal(driver):
    """返回当前最上层、可见的弹层容器。

    兼容：
    - AntD Modal: .ant-modal
    - AntD Drawer: .ant-drawer
    - 兜底：任意 role=dialog 或 aria-modal=true

    说明：为了尽量少改动原有代码，仍沿用旧函数名。
    """

    # 1) modal
    try:
        modals = driver.find_elements(By.CSS_SELECTOR, ".ant-modal")
    except Exception:
        modals = []

    for m in reversed(modals):
        try:
            if m.is_displayed():
                return m
        except Exception:
            continue

    # 2) drawer
    try:
        drawers = driver.find_elements(By.CSS_SELECTOR, ".ant-drawer")
    except Exception:
        drawers = []

    for d in reversed(drawers):
        try:
            if d.is_displayed():
                return d
        except Exception:
            continue

    # 3) fallback dialog
    try:
        dialogs = driver.find_elements(By.XPATH, "//*[@role='dialog' or @aria-modal='true']")
    except Exception:
        dialogs = []

    for dlg in reversed(dialogs):
        try:
            if dlg.is_displayed():
                return dlg
        except Exception:
            continue

    return None


def _get_ant_modal_body(modal):
    if not modal:
        return None

    # modal
    try:
        return modal.find_element(By.CSS_SELECTOR, ".ant-modal-body")
    except Exception:
        pass

    # drawer
    try:
        return modal.find_element(By.CSS_SELECTOR, ".ant-drawer-body")
    except Exception:
        pass

    # fallback: 有些 dialog 结构不是 AntD
    try:
        return modal
    except Exception:
        return None


def _scroll_ant_modal_to_bottom(driver, modal, steps=10, pause=0.2):
    """把弹窗内容区域滚动到最底部（用于触发懒加载/显示底部下载按钮）。"""
    body = _get_ant_modal_body(modal)
    if not body:
        return

    last_top = None
    for _ in range(steps):
        driver.execute_script("arguments[0].scrollTop = arguments[0].scrollHeight;", body)
        time.sleep(pause)
        try:
            top = driver.execute_script("return arguments[0].scrollTop;", body)
        except Exception:
            top = None
        if top is not None and top == last_top:
            break
        last_top = top


def _click_modal_close(driver, modal, timeout=10):
    """点击右上角关闭按钮并等待弹层消失（兼容 modal/drawer）。"""
    if not modal:
        return False

    close_btn = None
    for sel in ("button.ant-modal-close", "button.ant-drawer-close"):
        try:
            close_btn = modal.find_element(By.CSS_SELECTOR, sel)
            break
        except Exception:
            continue

    if close_btn is None:
        # 兜底：aria-label=Close
        try:
            close_btn = modal.find_element(
                By.XPATH,
                ".//button[@type='button' and @aria-label='Close']",
            )
        except Exception:
            return False

    try:
        driver.execute_script("arguments[0].click();", close_btn)
    except Exception:
        try:
            close_btn.click()
        except Exception:
            return False

    def _gone(_):
        try:
            return not modal.is_displayed()
        except Exception:
            return True

    WebDriverWait(driver, timeout).until(_gone)
    return True


def _find_cell_by_row_index_and_col_id(driver, row_index: str, col_id: str):
    """用 row-index + col-id 定位可见 cell（兼容 pinned/center 多套 row）。"""
    cells = driver.find_elements(
        By.XPATH,
        f"//div[@role='row' and @row-index='{row_index}']//div[@col-id='{col_id}']",
    )
    for c in cells:
        try:
            if c.is_displayed():
                return c
        except Exception:
            continue
    return cells[0] if cells else None


def _click_open_detail(driver, cell) -> bool:
    """尽量点击到真正的可点击控件（有些 cell 内部是 a/button）。"""
    if cell is None:
        return False

    # 若 cell 内有链接/按钮，优先点它
    for xp in (".//a", ".//button", ".//*[@role='button']"):
        try:
            targets = cell.find_elements(By.XPATH, xp)
            targets = [t for t in targets if t.is_displayed()]
            if targets:
                driver.execute_script("arguments[0].click();", targets[0])
                return True
        except Exception:
            continue

    # 否则点 cell 本身
    try:
        driver.execute_script("arguments[0].click();", cell)
        return True
    except Exception:
        try:
            cell.click()
            return True
        except Exception:
            return False


def _get_top_visible_select_listbox(driver):
    """返回当前最上层、可见的自定义下拉选项容器（role=listbox）。"""
    try:
        boxes = driver.find_elements(
            By.XPATH,
            "//div[@role='listbox' and contains(@class,'SelectOptions-module')]",
        )
    except Exception:
        return None

    for b in reversed(boxes):
        try:
            if b.is_displayed():
                return b
        except Exception:
            continue
    return None


def _get_top_visible_listbox_with_options(driver):
    """返回当前最上层、可见且“有选项”的 listbox。

    listbox 出现但内部可能还没渲染 optionLabel，
    这里用 role=option 作为更稳定的判断/选择依据。
    """
    try:
        boxes = driver.find_elements(
            By.XPATH,
            "//div[@role='listbox' and contains(@class,'SelectOptions-module')]",
        )
    except Exception:
        return None

    for b in reversed(boxes):
        try:
            if not b.is_displayed():
                continue
            opts = b.find_elements(By.XPATH, ".//*[@role='option']")
            if opts:
                return b
        except Exception:
            continue
    return None


def _extract_option_text(option_el):
    """从 option 元素里尽量提取用于匹配分数的文本。"""
    try:
        label = option_el.find_elements(
            By.XPATH, ".//*[contains(@class,'SelectOptions-module__optionLabel')]"
        )
        if label:
            t = (label[0].text or "").strip()
            if t:
                return t
    except Exception:
        pass
    try:
        return (option_el.text or "").strip()
    except Exception:
        return ""


def download_homework_file(
    driver,
    row,
    row_index,
    post_click_wait=2.0,
    open_attempts=4,
    per_attempt_wait=8,
):
    """新版页面：
    1) 先点击该行的 field_5 单元格打开详情/弹窗
    2) 弹窗里会出现多个下载按钮（a 标签）
    3) 只下载以 .cpp 结尾的附件（若有多个，按顺序逐个尝试，直到下载到 .cpp）

    关键适配：
    - 点击详情后，需要把弹窗内容下滑到最底部，才会显示下载按钮
    - 站点下载时可能先生成 *.tmp，必须等待其转为最终文件
    - 点击下载后固定等待 2s（post_click_wait）再开始轮询
    """

    current_row_index = row.get_attribute("row-index")

    # 若已存在弹层，直接复用（避免因为上一个未关闭导致等待失败）
    modal = _get_top_visible_ant_modal(driver)

    if modal is None:
        # 定位 field_5 单元格（优先用全局定位，避开 pinned/center 差异）
        cell = _find_cell_by_row_index_and_col_id(driver, current_row_index, "field_5")
        if cell is None:
            print(f"第 {row_index + 1} 行：未找到 field_5 单元格，跳过")
            return None

        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                cell,
            )
        except Exception:
            pass
        time.sleep(0.2)

        # 多次尝试点击打开详情（每次短等待，避免单行卡死）
        modal = None
        for _attempt in range(1, open_attempts + 1):
            ok = _click_open_detail(driver, cell)
            time.sleep(0.2)
            if not ok:
                continue
            try:
                WebDriverWait(driver, per_attempt_wait).until(
                    lambda drv: _get_top_visible_ant_modal(drv) is not None
                )
                modal = _get_top_visible_ant_modal(driver)
                break
            except TimeoutException:
                try:
                    driver.execute_script(
                        "arguments[0].scrollIntoView({block: 'center', inline: 'center'});",
                        cell,
                    )
                except Exception:
                    pass

        if modal is None:
            print(
                f"第 {row_index + 1} 行：点击 field_5 后仍未出现弹窗/抽屉（已重试 {open_attempts} 次），跳过"
            )
            return None

    start = time.time()
    all_links = []
    while time.time() - start < 20:
        modal = _get_top_visible_ant_modal(driver) or modal
        if modal:
            _scroll_ant_modal_to_bottom(driver, modal)
            try:
                # 兼容 a 或 button；同时保留 class=download 的旧线索
                all_links = modal.find_elements(
                    By.XPATH,
                    ".//a[contains(@href,'download')] | .//button[contains(.,'下载')] | .//*[contains(@class,'download')]",
                )
                all_links = [a for a in all_links if a.is_displayed()]
            except Exception:
                all_links = []

        if all_links:
            break
        time.sleep(0.2)

    cpp_links = [a for a in all_links if _is_cpp_download_link(a)]

    if not cpp_links:
        print(f"第 {row_index + 1} 行：弹窗中未找到带 .cpp 提示的下载按钮（将不下载）")
        return None

    if len(cpp_links) > 1:
        print(
            f"第 {row_index + 1} 行：发现 {len(cpp_links)} 个可能的 .cpp 附件，依次尝试下载直到拿到 .cpp："
        )
        for a in cpp_links:
            href = a.get_attribute("href") or ""
            dl = a.get_attribute("download") or ""
            att = _extract_filename_from_href(href)
            txt = (a.text or "").strip()
            print("  -", dl or att or txt or href)

    for idx, target in enumerate(cpp_links, start=1):
        clear_download_dir()

        file_name_hint = (
            (target.get_attribute("download") or "").strip()
            or _extract_filename_from_href(target.get_attribute("href") or "")
            or (target.get_attribute("title") or "").strip()
            or (target.text or "").strip()
        )

        print(
            f"下载第 {row_index + 1} 行（候选 {idx}/{len(cpp_links)}）: {file_name_hint or '(cpp 附件)'}"
        )

        try:
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", target)
            time.sleep(0.2)
        except Exception:
            pass

        try:
            driver.execute_script("arguments[0].click();", target)
        except Exception:
            try:
                target.click()
            except Exception:
                print("点击下载失败，尝试下一个候选")
                continue
        time.sleep(post_click_wait)

        downloaded = wait_download_complete(timeout=60)
        if not downloaded:
            print("下载超时，尝试下一个候选")
            continue

        base = os.path.basename(downloaded)
        print("下载完成:", base)

        if base.lower().endswith(".cpp"):
            return downloaded

        if _contains_cpp_hint(file_name_hint):
            print("注意：下载文件扩展名不是 .cpp，但链接提示是 .cpp，将尝试按文本读取：", base)
            return downloaded

        print("下载到的不是 .cpp，清理后尝试下一个候选:", base)

    return None


def read_cpp_file(file_path, max_bytes=2_000_000):
    """尽量把下载到的 C/C++ 源码按文本读出来。

    乱码通常来自编码识别错误。这里采用“多编码候选 + 质量打分”选最优解码。
    """

    if not file_path or not os.path.exists(file_path):
        print("read_cpp_file: 文件不存在:", file_path)
        return None

    size = os.path.getsize(file_path)
    print("文件大小:", size, "bytes")

    with open(file_path, "rb") as f:
        data = f.read(max_bytes + 1)

    if len(data) > max_bytes:
        data = data[:max_bytes]
        print(f"注意：文件过大，已截断到前 {max_bytes} bytes 读取")

    if data.startswith(b"PK\x03\x04"):
        print("read_cpp_file: 看起来像 ZIP/Office 文档（可能是 docx/xlsx），不是源码文本")
        return None
    if data.startswith(b"%PDF"):
        print("read_cpp_file: 看起来像 PDF，不是源码文本")
        return None

    def _score_text(text: str) -> tuple:
        if not text:
            return (10**9, 10**9, 10**9, 0)

        length = len(text)
        repl = text.count("�")
        nul = text.count("\x00")
        ctrl = sum(1 for c in text if ord(c) < 32 and c not in ("\n", "\r", "\t"))

        tokens = ["#include", "int", "main", "std::", "using", "return", ";", "{", "}"]
        token_hits = sum(1 for t in tokens if t in text)

        repl_ratio = repl / max(1, length)
        nul_ratio = nul / max(1, length)

        penalty = 0
        if repl_ratio > 0.02:
            penalty += int(repl_ratio * 10_000)
        if nul_ratio > 0.001:
            penalty += int(nul_ratio * 10_000)

        return (repl, ctrl, penalty, -token_hits)

    encodings = [
        "utf-8-sig",
        "utf-8",
        "gb18030",
        "gbk",
        "cp936",
        "utf-16",
        "utf-16-le",
        "utf-16-be",
        "big5",
    ]

    best = None
    best_enc = None
    best_score = None

    for enc in encodings:
        try:
            text = data.decode(enc, errors="replace")
        except Exception:
            continue

        if text.count("\x00") > max(50, len(text) // 10):
            continue

        sc = _score_text(text)
        if best is None or sc < best_score:
            best = text
            best_enc = enc
            best_score = sc

    if best is None:
        best = data.decode("utf-8", errors="replace")
        best_enc = "utf-8 (fallback)"
        best_score = _score_text(best)

    print("读取编码:", best_enc, "score=", best_score)

    if best.count("�") > max(10, len(best) // 50):
        print("警告：文本可能仍存在乱码（替换字符较多）。建议检查该作业源文件实际编码。")

    return best


def score_homework_with_ai(cpp_code):
    if not API_KEY:
        return None, "缺少 AI_API_KEY（环境变量/.env）"
    if not cpp_code or not cpp_code.strip():
        return None, "文件内容为空"

    client = OpenAI(api_key=API_KEY, base_url=BASE_URL)
    resp = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[
            {"role": "system", "content": SCORING_CRITERIA},
            {"role": "user", "content": f"请评分以下C++代码：\n{cpp_code}"},
        ],
        timeout=30,
    )

    result = resp.choices[0].message.content.strip()
    lines = result.split("\n")
    score = None
    comment = ""
    for line in lines:
        m = re.search(r"\d+(?:\.\d+)?", line)
        if m and score is None:
            score = m.group()
        elif line.strip() and not line.strip().isdigit():
            comment += line.strip() + " "
    return score, comment.strip()


def fill_score_and_comment(driver, row, score, comment=None):
    """回填（新版弹窗 + 自定义选择框）。"""

    score_str = str(score).strip() if score is not None else ""
    if not score_str:
        raise ValueError("score 为空，无法回填")

    modal = _get_top_visible_ant_modal(driver)

    def _find_edit():
        m = _get_top_visible_ant_modal(driver) or modal
        if m:
            try:
                btn = m.find_element(By.XPATH, ".//button[.//span[normalize-space()='修改']]")
                if btn.is_displayed() and btn.is_enabled():
                    return btn
            except Exception:
                pass
        try:
            btn = driver.find_element(By.XPATH, "//button[.//span[normalize-space()='修改']]")
            if btn.is_displayed() and btn.is_enabled():
                return btn
        except Exception:
            return None
        return None

    edit_btn = WebDriverWait(driver, 10).until(lambda _: _find_edit())
    try:
        edit_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", edit_btn)

    time.sleep(1)
    modal = _get_top_visible_ant_modal(driver) or modal

    def _find_score_input():
        m = _get_top_visible_ant_modal(driver) or modal
        if not m:
            return None

        xpaths = [
            ".//input[@placeholder='请选择' and not(@disabled)]",
            ".//input[contains(@class,'ant-select-selection-search-input') and not(@disabled)]",
        ]
        for xp in xpaths:
            try:
                el = m.find_element(By.XPATH, xp)
                if el.is_displayed() and el.is_enabled():
                    return el
            except Exception:
                continue
        return None

    score_input = WebDriverWait(driver, 10).until(lambda _: _find_score_input())
    score_input.click()
    time.sleep(0.2)

    WebDriverWait(driver, 10).until(lambda _: _get_top_visible_listbox_with_options(driver) is not None)
    listbox = _get_top_visible_listbox_with_options(driver) or _get_top_visible_select_listbox(driver)

    if not listbox:
        raise RuntimeError("未找到评分下拉（listbox）")

    options = listbox.find_elements(By.XPATH, ".//*[@role='option']")
    if not options:
        time.sleep(0.5)
        options = listbox.find_elements(By.XPATH, ".//*[@role='option']")

    if not options:
        try:
            html = listbox.get_attribute("outerHTML")
            print("DEBUG: listbox outerHTML (truncated)=", (html or "")[:500])
        except Exception:
            pass
        raise RuntimeError("未找到可用的评分选项（listbox 已出现，但无 role=option）")

    parsed = []
    for opt in options:
        txt = _extract_option_text(opt)
        if not txt:
            continue
        parsed.append((opt, txt))

    if not parsed:
        raise RuntimeError("未找到可用的评分选项（option 存在，但无法提取文本）")

    chosen = None
    for opt, txt in parsed:
        if txt == score_str:
            chosen = opt
            break

    if chosen is None:
        try:
            target_val = float(score_str)
        except Exception:
            target_val = None

        if target_val is not None:
            best = None
            for opt, txt in parsed:
                try:
                    v = float(txt)
                except Exception:
                    continue
                diff = abs(v - target_val)
                if best is None or diff < best[0]:
                    best = (diff, opt, txt)
            if best is not None:
                chosen = best[1]
                print(f"评分 {score_str} 不在下拉中，改选最接近的：{best[2]}")

    if chosen is None:
        print("警告：无法解析分数选项文本，将默认选择第一个 option：", parsed[0][1])
        chosen = parsed[0][0]

    driver.execute_script("arguments[0].click();", chosen)
    time.sleep(0.2)

    def _find_submit():
        m = _get_top_visible_ant_modal(driver) or modal
        if m:
            try:
                btn = m.find_element(By.XPATH, ".//button[.//span[normalize-space()='提交']]")
                if btn.is_displayed() and btn.is_enabled():
                    return btn
            except Exception:
                pass
        try:
            btn = driver.find_element(By.XPATH, "//button[.//span[normalize-space()='提交']]")
            if btn.is_displayed() and btn.is_enabled():
                return btn
        except Exception:
            return None
        return None

    submit_btn = WebDriverWait(driver, 10).until(lambda _: _find_submit())
    try:
        submit_btn.click()
    except Exception:
        driver.execute_script("arguments[0].click();", submit_btn)

    time.sleep(0.5)
    modal = _get_top_visible_ant_modal(driver) or modal
    closed = _click_modal_close(driver, modal, timeout=10)
    if not closed:
        print("已提交，但未找到/未能点击关闭按钮（请手动关闭弹窗）")
    else:
        print("回填完成（已提交并关闭弹窗）")

    if comment and str(comment).strip():
        pass

    # 当前回填逻辑不依赖 row，但保留参数以兼容批量调用
    _ = row


def _get_ag_row_cell_text(row, col_id: str) -> str:
    """从 AG Grid 的某行某列拿到可见文本（用于判断是否已有评分）。"""
    if not row or not col_id:
        return ""
    try:
        cell = row.find_element(By.XPATH, f".//div[@col-id='{col_id}']")
    except Exception:
        return ""

    try:
        val = cell.find_element(By.XPATH, ".//div[contains(@class,'ag-cell-value')]")
        txt = (val.text or "").strip()
    except Exception:
        txt = (cell.text or "").strip()

    if not txt:
        try:
            txt = (cell.get_attribute("title") or "").strip()
        except Exception:
            txt = ""
    return txt


def _row_has_teacher_score(row, score_col_id: str = "field_11") -> tuple[bool, str]:
    """判断该行是否已有教师评分。返回：(是否已评分, 原始文本)。"""
    txt = _get_ag_row_cell_text(row, score_col_id)
    if not txt:
        return False, ""
    if re.search(r"\d", txt):
        return True, txt
    return False, txt


def process_all_visible_then_scroll(
    driver,
    viewport,
    max_loops=9999,
    skip_if_scored=True,
    score_col_id="field_11",
):
    processed: set[int] = set()

    def _find_center_row_by_index(row_index: int):
        return driver.find_element(
            By.XPATH,
            f"//div[contains(@class,'ag-center-cols-container')]//div[@role='row' and @row-index='{row_index}']",
        )

    for _ in range(max_loops):
        # 先“快照”当前可见的 row-index 列表（不要把 row WebElement 长期保存）
        row_indices: list[int] = []
        for r in get_visible_rows(driver):
            try:
                idx_str = r.get_attribute("row-index")
            except StaleElementReferenceException:
                continue
            except Exception:
                continue

            if not idx_str:
                continue
            try:
                row_indices.append(int(idx_str))
            except ValueError:
                continue

        new_rows = 0

        for idx in row_indices:
            if idx in processed:
                continue

            processed.add(idx)
            new_rows += 1

            # 每次处理前都重新定位“新鲜”的 row 元素
            try:
                r = _find_center_row_by_index(idx)
            except Exception:
                continue

            # 跳过：已有教师评分的行（field_11）
            if skip_if_scored:
                try:
                    has_score, raw = _row_has_teacher_score(r, score_col_id=score_col_id)
                except StaleElementReferenceException:
                    try:
                        r = _find_center_row_by_index(idx)
                        has_score, raw = _row_has_teacher_score(r, score_col_id=score_col_id)
                    except Exception:
                        continue

                if has_score:
                    print(f"\n--- 跳过第 {idx + 1} 份作业：已有教师评分 {raw} ---")
                    continue

            print(f"\n--- 处理第 {idx + 1} 份作业 ---")

            try:
                downloaded = download_homework_file(driver, r, idx)
            except StaleElementReferenceException:
                # 行被重渲染：跳过本行，下一轮滚动/刷新时再碰到就会处理
                print("行元素已失效（stale），跳过本行，继续...")
                continue

            if not downloaded:
                print("下载失败，跳过")
                continue

            cpp_code = read_cpp_file(downloaded)
            if not cpp_code:
                print("读取失败（可能下载到的不是源码文件），跳过")
                continue

            score, comment = score_homework_with_ai(cpp_code)
            if not score:
                print("评分失败，跳过：", comment)
                continue

            print("score =", score)
            print("comment =", comment)

            try:
                fill_score_and_comment(driver, r, score, comment)
            except StaleElementReferenceException:
                # 提交/关闭弹窗后 grid 重渲染是正常的，忽略即可
                print("回填后行元素变 stale（正常），继续...")

        is_bottom = driver.execute_script(
            "return arguments[0].scrollTop + arguments[0].clientHeight >= arguments[0].scrollHeight - 50;",
            viewport,
        )

        if is_bottom and new_rows == 0:
            print("已到底部，结束。总处理:", len(processed))
            break

        print("向下滚动加载更多...")
        driver.execute_script("arguments[0].scrollTop += arguments[0].clientHeight;", viewport)
        time.sleep(2)

    return processed


def main():
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)
    print("DOWNLOAD_DIR =", DOWNLOAD_DIR)
    print("AI_API_KEY present =", bool(API_KEY))

    if not API_KEY:
        print("错误：缺少 AI_API_KEY（请在环境变量或 .env 中设置）")
        raise SystemExit(1)

    driver = setup_driver()

    try:
        driver.get(HOMEWORK_URL)
        print("已打开页面：", HOMEWORK_URL)
        input("请在浏览器中完成登录，然后回到这里按回车继续... ")

        viewport = wait_for_grid(driver)
        print("AG Grid 已就绪")

        processed = process_all_visible_then_scroll(driver, viewport)
        print("处理完成，总计行数：", len(processed))
    finally:
        input("按回车关闭浏览器... ")
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    main()
