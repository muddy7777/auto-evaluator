# selenium_operator

一个用于“金数据（Jinshuju）作业表格”批量下载、AI 评分、并回填教师评分的 Selenium 自动化脚本/Notebook。

当前项目同时提供：
- `main.ipynb`：分步骤调试用（适合边看页面边改 XPath）
- `main.py`：将 notebook 的逻辑整理成可直接运行的脚本

## 功能概述

- 进入金数据表单条目列表（AG Grid）
- 对每一行作业：
  - 点击 `field_5` 打开详情弹层（兼容 Ant Design Modal/Drawer）
  - 在弹层内滚动到底部，找到下载入口
  - 只下载 `.cpp` 附件（点击后固定等待 2 秒；并等待 `.tmp/.crdownload` 消失 + 文件大小稳定）
  - 读取源码文本并做“多编码候选 + 质量打分”解码，尽量降低乱码
  - 调用 OpenAI 生成分数与简短评语
  - 回填：点击“修改”→点“请选择”→在 listbox 里点 `role=option` 分数→“提交”→右上角关闭
- 跳过已评分行：如果 `field_11`（教师评分）已有数字则跳过
- 批量滚动处理：通过 `row-index` 重新定位行，规避滚动后元素 stale

## 环境要求

- Windows/macOS/Linux
- Python 3.10+（建议 3.11/3.12）
- Google Chrome（与 ChromeDriver 版本匹配由 `webdriver_manager` 自动处理）

## 安装

在项目目录下执行：

```bash
pip install -r requirements.txt
```

依赖在 [requirements.txt](requirements.txt)：
- `selenium`
- `webdriver_manager`
- `openai`
- `python-dotenv`

## 配置（只从环境变量读取）

脚本不会硬编码 API Key，请通过环境变量或 `.env` 配置。

必须：
- `AI_API_KEY`：OpenAI API Key

可选：
- `AI_BASE_URL`：OpenAI 兼容网关地址（默认：`https://api.openai-proxy.org/v1`）
- `HOMEWORK_URL`：金数据 entries 页地址（默认写在代码里）
- `MODEL_NAME`：模型名（默认：`gpt-5-mini`）

示例 `.env`：

```env
AI_API_KEY=你的key
# AI_BASE_URL=https://api.openai.com/v1
# HOMEWORK_URL=https://next.jinshuju.net/forms/xxxx/entries
# MODEL_NAME=gpt-5-mini
```

## 运行（脚本版）

```bash
python main.py
```

运行后流程：
1. 打开 Chrome 并跳转到 `HOMEWORK_URL`
2. 你需要在浏览器里手动登录
3. 回到终端按回车继续，脚本开始批量处理

下载文件默认保存到项目内的 [downloads/](downloads/)。

## 运行（Notebook 调试版）

打开 [main.ipynb](main.ipynb) 并按顺序执行：
1. 导入 & 配置
2. 函数定义（集中）
3. 初始化浏览器并手动登录
4. 等待表格加载
5. 先做单行下载/评分/回填测试
6. 最后再运行批量滚动处理

## 常见问题

### 1) 批量中某行“点击后不弹出详情”，然后超时

已做处理：脚本会对 `field_5` 的打开动作做多次点击重试（每次短等待），仍失败则跳过该行，避免卡死。

如果仍频繁出现：
- 可能是页面结构变化（`field_5` 里真正可点击元素变了）
- 可能是弹层不再是 Ant Design（需更新弹层识别选择器）

### 2) 下载过程中出现 `.tmp` / `.crdownload`

这是正常的临时文件行为。脚本会：
- 点击下载后固定 `sleep(2)`
- 然后轮询下载目录，忽略 `.tmp/.crdownload`，并等待最终文件大小稳定

### 3) 回填分数不能直接输入

页面的分数控件是下拉 listbox（`role=listbox` / `role=option`），需要“点击 option”，不能用键盘输入。

### 4) 运行很慢

脚本默认关闭 Selenium `implicitly_wait`，避免与 `WebDriverWait` 叠加导致每次查找都被放大。

## 免责声明

- 本项目用于自动化你有权限访问的数据页面。
- 目标网站更新 UI/DOM 后，XPath/CSS 选择器可能需要调整。
