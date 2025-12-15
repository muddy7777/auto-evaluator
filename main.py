from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import time
from openai import OpenAI
from dotenv import load_dotenv
import os
import glob
import re

# 加载环境变量（存储大模型API密钥）
load_dotenv()

# ---------------------- 配置参数 ----------------------
HOMEWORK_URL = os.getenv("HOMEWORK_URL")
MODEL_NAME = os.getenv("MODEL_NAME")  # 大模型型号（可替换为gpt-4、glm-4等）

# 优先从环境变量读取 API Key，可在 .env 中设置 AI_API_KEY=xxxx
API_KEY = os.getenv("AI_API_KEY")
BASE_URL = os.getenv("AI_BASE_URL")
DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")  # 下载目录

# 评分标准（可根据实际作业要求修改）
SCORING_CRITERIA = """
你是C++作业评分助教，按以下标准评分（满分10分，平均分8分）：
1. 代码逻辑正确性：是否符合作业需求，逻辑无漏洞；
2. 代码规范性：命名规范、缩进整齐、结构清晰；
3. 注释完整性：关键步骤有注释，便于理解；
4. 代码简洁性：无冗余代码，实现高效。
评分输出格式：
第一行：分数（仅数字，例如：8.5）
第二行：简短评语（例如：代码逻辑正确，命名规范，注释完整，建议优化循环结构以提升简洁性）
"""

# ---------------------- 1. 初始化浏览器驱动 ----------------------
def setup_driver():
    """设置Chrome浏览器驱动，配置下载目录"""
    # 确保下载目录存在
    os.makedirs(DOWNLOAD_DIR, exist_ok=True)

    chrome_options = Options()
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    # 关键修改：使用 webdriver_manager 自动管理 ChromeDriver
    # 如果本机没有驱动或版本不匹配，会自动下载合适版本
    service = Service(ChromeDriverManager().install())

    # 显式传入 service 和 options
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.implicitly_wait(10)
    return driver

# ---------------------- 2. 获取表格中的作业行 ----------------------
def get_homework_rows(driver):
    """获取表格中所有作业行"""
    try:
        # 等待表格加载完成 (AG Grid)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ag-root"))
        )

        # 查找表格行 (AG Grid rows)
        # 关键修改：只获取 ag-center-cols-container 中的行，避免获取到左侧固定列(pinned-left)的行
        # 左侧固定列的行不包含中间滚动的列（如 field_5），会导致找不到元素
        rows = driver.find_elements(By.XPATH, "//div[contains(@class, 'ag-center-cols-container')]//div[@role='row']")
        
        # 如果中间区域没有行（可能是没有数据或者加载问题），尝试获取所有行作为回退
        if not rows:
            print("中间区域未找到行，尝试获取所有行...")
            rows = driver.find_elements(By.XPATH, "//div[@role='row' and contains(@class, 'ag-row') and not(contains(@class, 'ag-header-row'))]")
            
        print(f"找到 {len(rows)} 行作业数据")
        return rows
    except TimeoutException:
        print("未找到表格，请检查页面结构")
        return []

# ---------------------- 3. 下载单个作业文件 ----------------------
def download_homework_file(driver, row, row_index):
    """下载指定行的作业文件"""
    try:
        # 清空下载目录中的旧文件
        for file in glob.glob(os.path.join(DOWNLOAD_DIR, "*")):
            try:
                os.remove(file)
            except:
                pass

        # 获取当前行的 row-index，用于精确查找
        current_row_index = row.get_attribute("row-index")
        
        # 查找程序文件列的下载链接
        # col-id 为 "field_5"
        try:
            # 尝试直接在当前 row 元素下查找
            # 如果 row 是 pinned row，这里会失败，但我们已经更新了 get_homework_rows
            cell = row.find_element(By.XPATH, ".//div[@col-id='field_5']")
        except NoSuchElementException:
            # 如果在当前 row 找不到（可能是 row 引用问题），尝试使用 row-index 全局查找
            print(f"当前行元素未找到 field_5，尝试全局查找 row-index={current_row_index}...")
            try:
                cell = driver.find_element(By.XPATH, f"//div[@role='row' and @row-index='{current_row_index}']//div[@col-id='field_5']")
            except NoSuchElementException:
                print(f"第 {row_index + 1} 行 (row-index={current_row_index}) 未找到程序文件单元格(col-id='field_5')")
                return None

        # 滚动到可见区域
        driver.execute_script("arguments[0].scrollIntoView({block: 'center', inline: 'center'});", cell)
        time.sleep(0.5)

        # 在单元格内查找链接
        try:
            # 查找包含 href 的 a 标签
            download_link = cell.find_element(By.XPATH, ".//a[@href]")
            
            # 打印文件名以便调试
            file_name = download_link.get_attribute("download") or "未知文件名"
            print(f"正在下载第 {row_index + 1} 行的作业文件: {file_name}")
            
            # 使用 JavaScript 点击
            driver.execute_script("arguments[0].click();", download_link)
        except NoSuchElementException:
            # 尝试直接点击单元格
            print(f"第 {row_index + 1} 行未找到 a 标签，尝试点击单元格")
            driver.execute_script("arguments[0].click();", cell)

        # 等待下载完成
        timeout = 30
        start_time = time.time()
        downloaded_file = None

        while time.time() - start_time < timeout:
            files = glob.glob(os.path.join(DOWNLOAD_DIR, "*"))
            # 过滤掉临时下载文件（.crdownload）
            complete_files = [f for f in files if not f.endswith('.crdownload')]

            if complete_files:
                downloaded_file = complete_files[0]
                break
            time.sleep(1)

        if downloaded_file:
            print(f"下载完成: {os.path.basename(downloaded_file)}")
            return downloaded_file
        else:
            print(f"下载超时，第 {row_index + 1} 行作业文件下载失败")
            return None

    except NoSuchElementException:
        print(f"第 {row_index + 1} 行未找到下载链接")
        return None
    except Exception as e:
        print(f"下载第 {row_index + 1} 行作业文件时发生错误: {str(e)}")
        return None

# ---------------------- 4. 读取C++文件内容 ----------------------
def read_cpp_file(file_path):
    """读取C++文件内容"""
    try:
        # 尝试多种编码格式
        encodings = ['utf-8', 'gbk', 'gb2312', 'latin-1']

        for encoding in encodings:
            try:
                with open(file_path, 'r', encoding=encoding) as f:
                    content = f.read()
                print(f"成功读取文件 (编码: {encoding})")
                return content
            except UnicodeDecodeError:
                continue

        print(f"无法读取文件 {file_path}，尝试了多种编码格式")
        return None

    except Exception as e:
        print(f"读取文件时发生错误: {str(e)}")
        return None

# ---------------------- 5. AI评分功能 ----------------------
def score_homework_with_ai(cpp_code):
    """使用AI对C++代码进行评分"""
    if not cpp_code or not cpp_code.strip():
        return None, "文件内容为空"

    try:
        # 初始化OpenAI客户端
        
        client = OpenAI(api_key=API_KEY, base_url=BASE_URL)

        response = client.chat.completions.create(
            model=MODEL_NAME,
            messages=[
                {"role": "system", "content": SCORING_CRITERIA},
                {"role": "user", "content": f"请评分以下C++代码：\n{cpp_code}"}
            ],
            timeout=30
        )

        result = response.choices[0].message.content.strip()

        # 解析评分结果
        lines = result.split('\n')
        score = None
        comment = ""

        for line in lines:
            # 提取数字分数
            score_match = re.search(r'\d+', line)
            if score_match and not score:
                score = score_match.group()
            elif line.strip() and not line.strip().isdigit():
                comment += line.strip() + " "

        return score, comment.strip()

    except Exception as e:
        print(f"AI评分失败: {str(e)}")
        return None, f"评分失败: {str(e)}"

# ---------------------- 6. 填入评分结果（处理弹框） ----------------------
def fill_score_and_comment(driver, row, score, comment):
    """点击单元格，处理弹框，填入评分和评语"""
    try:
        # 填入教师评分
        try:
            # 教师评分列 col-id="field_11"
            score_cell = row.find_element(By.XPATH, ".//div[@col-id='field_11']")
            score_cell.click()
            time.sleep(1)

            # 等待弹框出现并点击修改按钮
            # 注意：如果没有弹框直接编辑，这里需要调整。假设点击后会出现弹框或进入编辑模式。
            # 如果是 AG Grid 双击编辑，可能需要 ActionChains(driver).double_click(score_cell).perform()
            
            # 尝试查找“修改”按钮，如果找不到可能已经进入编辑模式
            try:
                edit_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '修改')]"))
                )
                edit_button.click()
                time.sleep(1)
            except TimeoutException:
                # 可能是直接编辑，或者弹框结构不同
                pass

            # 在弹框中的输入框输入分数
            # 查找可见的输入框
            score_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'modal')]//input[@type='text' or @type='number'] | //input[contains(@class, 'ag-input-field-input')]"))
            )
            # 清除原有内容 (AG Grid 输入框可能需要特殊处理)
            score_input.send_keys(Keys.CONTROL + "a")
            score_input.send_keys(Keys.DELETE)
            score_input.send_keys(str(score))

            # 点击完成按钮或回车
            try:
                complete_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '完成') or contains(text(), '确定')]"))
                )
                complete_button.click()
            except TimeoutException:
                # 尝试回车
                score_input.send_keys(Keys.ENTER)
            
            time.sleep(1)
            print(f"成功填入评分: {score}")

        except Exception as e:
            print(f"填入评分失败: {str(e)}")

        # 填入教师答复
        try:
            # 教师答复列 col-id="field_12"
            comment_cell = row.find_element(By.XPATH, ".//div[@col-id='field_12']")
            comment_cell.click()
            time.sleep(1)

            # 等待弹框出现并点击修改按钮
            try:
                edit_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '修改')]"))
                )
                edit_button.click()
                time.sleep(1)
            except TimeoutException:
                pass

            # 在弹框中的文本框输入评语
            comment_input = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'modal')]//textarea | //div[contains(@class, 'modal')]//input[@type='text'] | //textarea"))
            )
            comment_input.send_keys(Keys.CONTROL + "a")
            comment_input.send_keys(Keys.DELETE)
            comment_input.send_keys(comment)

            # 点击完成按钮
            try:
                complete_button = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), '完成') or contains(text(), '确定')]"))
                )
                complete_button.click()
            except TimeoutException:
                comment_input.send_keys(Keys.ENTER)
                
            time.sleep(1)
            print(f"成功填入评语: {comment[:50]}...")

        except Exception as e:
            print(f"填入评语失败: {str(e)}")

    except Exception as e:
        print(f"填入评分结果时发生错误: {str(e)}")

# ---------------------- 7. 主流程 ----------------------
def process_homeworks():
    """主流程：下载作业、评分、填入结果"""
    driver = setup_driver()

    try:
        # 打开作业页面
        driver.get(HOMEWORK_URL)
        print("页面加载完成，请在浏览器中完成登录...")
        input("登录完成后按回车继续...")

        # 查找滚动容器 (AG Grid viewport)
        try:
            viewport = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.CLASS_NAME, "ag-body-viewport"))
            )
        except TimeoutException:
            print("未找到表格滚动区域")
            return

        processed_indices = set()
        
        while True:
            # 获取当前可视区域的行
            # 注意：只获取 ag-center-cols-container 中的行
            rows = driver.find_elements(By.XPATH, "//div[contains(@class, 'ag-center-cols-container')]//div[@role='row']")
            
            new_rows_count = 0
            for row in rows:
                # 获取行号 (AG Grid 的 row-index 属性)
                row_index_str = row.get_attribute("row-index")
                if not row_index_str:
                    continue
                
                row_index = int(row_index_str)
                
                # 如果已经处理过，跳过
                if row_index in processed_indices:
                    continue
                
                # 标记为已处理
                processed_indices.add(row_index)
                new_rows_count += 1
                
                print(f"\n--- 处理第 {row_index + 1} 份作业 ---")

                # 1. 下载作业文件
                downloaded_file = download_homework_file(driver, row, row_index)
                if not downloaded_file:
                    print(f"跳过第 {row_index + 1} 份作业（下载失败）")
                    continue

                # 2. 读取文件内容
                cpp_code = read_cpp_file(downloaded_file)
                if not cpp_code:
                    print(f"跳过第 {row_index + 1} 份作业（文件读取失败）")
                    continue

                # 3. AI评分
                score, comment = score_homework_with_ai(cpp_code)
                if not score:
                    print(f"跳过第 {row_index + 1} 份作业（AI评分失败）")
                    continue

                print(f"评分结果: {score}分")
                print(f"评语: {comment}")

                # 4. 填入评分结果
                fill_score_and_comment(driver, row, score, comment)

                print(f"第 {row_index + 1} 份作业处理完成")
                # time.sleep(1) # 稍微等待，避免操作过快

            # 检查是否滚动到底部
            # scrollTop + clientHeight >= scrollHeight - tolerance
            is_bottom = driver.execute_script(
                "return arguments[0].scrollTop + arguments[0].clientHeight >= arguments[0].scrollHeight - 50;", 
                viewport
            )
            
            if is_bottom and new_rows_count == 0:
                print("已滚动到底部，所有作业处理完成")
                break
            
            # 向下滚动 (滚动约一屏的高度)
            print("向下滚动加载更多作业...")
            driver.execute_script("arguments[0].scrollTop += arguments[0].clientHeight;", viewport)
            time.sleep(2) # 等待加载

        print(f"\n共处理了 {len(processed_indices)} 份作业！")

    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")

    finally:
        input("按回车键关闭浏览器...")
        driver.quit()

# ---------------------- 主函数 ----------------------
if __name__ == "__main__":
    # 检查API密钥
    if not API_KEY:
        print("错误: 请在.env文件中设置 AI_API_KEY")
        exit(1)

    # 开始处理作业
    process_homeworks()
