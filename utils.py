import pandas as pd
import logging
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import os
import requests
import zipfile
import json
import time
from PyQt6.QtCore import QSettings

def read_categories_from_excel(file_path):
    try:
        # 读取Excel文件
        df = pd.read_excel(file_path)
        # 假设类别在第一列，可以根据实际情况调整
        categories = df.iloc[:, 0].dropna().tolist()
        return categories if categories else []
    except Exception as e:
        logging.error(f"读取Excel类别时出错: {str(e)}")
        return []

def read_sheet_names_from_excel(file_path):
    try:
        # 读取Excel文件的所有sheet名称
        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names
        return sheet_names if sheet_names else []
    except Exception as e:
        logging.error(f"读取Excel工作表名称时出错: {str(e)}")
        return []

def get_chrome_version():
    try:
        chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
        if os.path.exists(chrome_path):
            from win32com.client import Dispatch
            parser = Dispatch('Scripting.FileSystemObject')
            version = parser.GetFileVersion(chrome_path)
            return version
    except Exception as e:
        logging.error(f"获取Chrome版本失败: {str(e)}")
    return None

def get_chromedriver_from_mirrors(version):
    """从多个镜像源尝试下载ChromeDriver"""
    mirrors = [
        {
            'name': '淘宝镜像',
            'version_url': f"https://registry.npmmirror.com/-/binary/chromedriver/{version}",
            'download_url': lambda v: f"https://registry.npmmirror.com/-/binary/chromedriver/{v}/chromedriver_win32.zip"
        },
        {
            'name': '中科大镜像',
            'version_url': f"https://mirrors.ustc.edu.cn/chromedriver/{version}/",
            'download_url': lambda v: f"https://mirrors.ustc.edu.cn/chromedriver/{v}/chromedriver_win32.zip"
        },
        {
            'name': '官方源',
            'version_url': "https://chromedriver.storage.googleapis.com/LATEST_RELEASE_" + version,
            'download_url': lambda v: f"https://chromedriver.storage.googleapis.com/{v}/chromedriver_win32.zip"
        }
    ]
    
    for mirror in mirrors:
        try:
            logging.info(f"尝试使用{mirror['name']}下载ChromeDriver...")
            response = requests.get(mirror['version_url'], timeout=10)
            
            if response.status_code == 200:
                if mirror['name'] == '淘宝镜像':
                    versions = response.json()
                    if versions:
                        latest_version = versions[-1]
                else:
                    latest_version = response.text.strip()
                
                download_url = mirror['download_url'](latest_version)
                logging.info(f"从{mirror['name']}下载 ChromeDriver {latest_version}")
                
                response = requests.get(download_url, timeout=30)
                if response.status_code == 200:
                    return response.content
                    
        except Exception as e:
            logging.warning(f"{mirror['name']}下载失败: {str(e)}")
            continue
    
    return None

def download_chromedriver():
    try:
        chrome_version = get_chrome_version()
        if not chrome_version:
            return None
            
        major_version = chrome_version.split('.')[0]
        
        # 尝试从镜像下载
        content = get_chromedriver_from_mirrors(major_version)
        if not content:
            logging.error("所有镜像源下载失败")
            return None
            
        # 保存并解压
        with open("chromedriver.zip", "wb") as f:
            f.write(content)
        
        # 如果存在旧的chromedriver，先删除
        if os.path.exists("chromedriver.exe"):
            os.remove("chromedriver.exe")
        
        with zipfile.ZipFile("chromedriver.zip", "r") as zip_ref:
            zip_ref.extractall()
            
        os.remove("chromedriver.zip")
        logging.info("ChromeDriver下载并解压成功")
        return "chromedriver.exe"
            
    except Exception as e:
        logging.error(f"下载ChromeDriver失败: {str(e)}")
    return None

def open_browser(driver_path=None, user_data_dir=None):
    try:
        options = webdriver.ChromeOptions()
        if user_data_dir:
            options.add_argument(f'user-data-dir={user_data_dir}')
        
        # 添加其他必要的选项
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        # 添加代理支持
        proxy = os.environ.get('HTTP_PROXY') or os.environ.get('HTTPS_PROXY')
        if proxy:
            options.add_argument(f'--proxy-server={proxy}')
        
        settings = QSettings('ImportifyApp', 'Settings')
        auto_download = settings.value('auto_download', True, type=bool)
        
        # 如果设置为自动下载或没有指定driver_path
        if auto_download or not driver_path:
            logging.info("尝试自动下载ChromeDriver...")
            downloaded_driver = download_chromedriver()
            if downloaded_driver:
                driver_path = downloaded_driver
            elif not driver_path:  # 如果下载失败且没有手动指定路径
                logging.error("无法自动下载ChromeDriver，请手动指定路径")
                return None
        
        if not os.path.exists(driver_path):
            logging.error(f"ChromeDriver路径不存在: {driver_path}")
            return None
            
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        logging.info("成功创建Chrome浏览器实例")
        return driver
        
    except Exception as e:
        logging.error(f"创建浏览器实例时出错: {str(e)}")
        if "This version of ChromeDriver only supports Chrome version" in str(e):
            logging.info("检测到版本不匹配，尝试重新下载...")
            if settings.value('auto_download', True, type=bool):
                driver_path = download_chromedriver()
                if driver_path:
                    try:
                        service = Service(driver_path)
                        driver = webdriver.Chrome(service=service, options=options)
                        logging.info("使用新下载的ChromeDriver成功创建浏览器实例")
                        return driver
                    except Exception as new_e:
                        logging.error(f"使用新下载的ChromeDriver仍然失败: {str(new_e)}")
            else:
                logging.error("ChromeDriver版本不匹配，请手动下载正确版本")
        return None

def process_link(driver, link, category, sheet_name):
    max_retries = 3
    retry_count = 0
    
    while retry_count < max_retries:
        try:
            # 检查driver是否还有效
            try:
                driver.current_url
            except:
                logging.error("浏览器已关闭，需要重新初始化")
                driver = open_browser()
                if not driver:
                    return 0
            
            # 访问链接并切换到主窗口
            logging.info(f"访问页面: {link}")
            driver.get(link)
            driver.switch_to.window(driver.window_handles[0])
            
            # 处理搜索和产品列表 - 传递sheet_name参数
            success_count = handle_product_detail(driver, category, sheet_name)
            if success_count is None:
                retry_count += 1
                continue
                
            return success_count
            
        except Exception as e:
            logging.error(f"处理类别 '{category}' 出错: {str(e)}")
            retry_count += 1
            if retry_count >= max_retries:
                return 0
            
            try:
                driver.quit()
            except:
                pass
            driver = open_browser()
            if not driver:
                return 0
            
            time.sleep(2)
    
    return 0

def handle_product_detail(driver, category, sheet_name):
    try:
        # 等待搜索框加载
        search_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input.search-bar-input.util-ellipsis'))
        )
        search_input.clear()
        search_input.send_keys(category)
        
        # 点击搜索按钮
        search_button = driver.find_element(By.CSS_SELECTOR, 'button.fy23-icbu-search-bar-inner-button')
        search_button.click()
        
        # 等待产品列表加载
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CLASS_NAME, "organic-list"))
        )
        
        # 滚动加载所有产品
        last_height = driver.execute_script("return document.body.scrollHeight")
        scroll_attempts = 0
        max_scrolls = 3  # 限制滚动次数
        
        while scroll_attempts < max_scrolls:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)
            new_height = driver.execute_script("return document.body.scrollHeight")
            if new_height == last_height:
                break
            last_height = new_height
            scroll_attempts += 1
        
        # 获取产品列表
        product_list = driver.find_elements(By.CLASS_NAME, "fy23-search-card")
        num_products = len(product_list)
        logging.info(f"找到 {num_products} 个产品")
        
        # 处理每个产品
        success_count = 0
        for product in product_list:
            try:
                # 获取产品标题
                product_title = product.find_element(By.CLASS_NAME, "search-card-e-title")
                logging.info(f"当前产品标题: {product_title.text}")
                
                # 获取产品链接并打开
                product_link = product.find_element(By.TAG_NAME, "a").get_attribute("href")
                driver.execute_script(f"window.open('{product_link}')")
                
                # 获取所有窗口句柄
                original_window = driver.current_window_handle
                new_window = [handle for handle in driver.window_handles if handle != original_window][0]
                
                # 切换到新窗口
                driver.switch_to.window(new_window)
                
                # 等待产品详情页加载
                product_title = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.TAG_NAME, "h1"))
                )
                
                # 处理产品详情页操作 - 传递sheet_name参数
                success_count = handle_product_actions(driver, category, success_count, sheet_name)
                
                # 关闭新窗口并切回原窗口
                driver.close()
                driver.switch_to.window(original_window)
                
            except Exception as e:
                logging.error(f"处理产品时出错: {str(e)}")
                try:
                    driver.close()
                    driver.switch_to.window(original_window)
                except:
                    pass
                continue
        
        return success_count
        
    except Exception as e:
        logging.error(f"处理产品详情时出错: {str(e)}")
        return 0

def handle_product_actions(driver, category, success_count, sheet_name):
    try:
        logging.info(f"处理产品详情页操作: {category}, {sheet_name}")
        
        # 点击添加按钮
        add_btn_con = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//*[@id="addBtnCon"]')))
        add_btn_con.click()
        logging.info("点击了添加按钮")

        # 处理Draft元素
        try:
            draft_element = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.XPATH, '//span[@class="inactive" and text()="Draft"]'))
            )
            logging.info("成功加载 Draft 元素")
            actions = ActionChains(driver)
            actions.move_to_element(draft_element).perform()
            draft_element.click()
            logging.info("成功点击 Draft 元素")
            time.sleep(2)
        except Exception as e:
            logging.error(f"等待和点击 Draft 元素时出错：{e}")
            return success_count

        # 检查区域限制
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Sorry, this product can\'t be shipped to your region.")]'))
            )
            logging.info("检测到产品无法配送到当前区域，跳过处理")
            return success_count
        except TimeoutException:
            logging.info("未检测到区域限制消息，继续处理")

        # 检查产品是否已存在
        try:
            success_message = driver.find_element(By.XPATH, '//div[@class="textcontainer centeralign home-content "]/p[1]')
            if success_message.text == "This product is already in your store, what would you like to do?":
                logging.info("产品已存在，不再处理")
                return success_count
        except NoSuchElementException:
            pass

        # 选择类别
        select_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '//button[@class="ms-choice"]'))
        )
        logging.info("等待并点击选择按钮")
        select_button.click()

        # 处理下拉选项并等待完成
        try:
            fetch_dropdown_options(driver, sheet_name)
        except Exception as e:
            logging.error(f"选择类别失败: {e}")
            return success_count

        # 不需要再点击描述标签，因为已经在fetch_dropdown_options中处理了
        time.sleep(3)

        # 处理变体
        try:
            variants_button = driver.find_element(By.CSS_SELECTOR, 'button.accordion-tab[data-actab-group="0"][data-actab-id="2"]')
            variants_button.click()
            logging.info("点击了变体按钮")

            # 选择变体选项
            all_variants_radio = driver.find_element(By.ID, 'all_variants')
            all_variants_radio.click()
            time.sleep(3)

            price_switch_radio = driver.find_element(By.ID, 'price_switch')
            price_switch_radio.click()
            time.sleep(3)
        except Exception as e:
            logging.error(f"处理变体时出错：{e}")
            return success_count

        # 处理图片
        images_button = driver.find_element(By.XPATH, '//button[@class="accordion-tab accordion-custom-tab" and @data-actab-group="0" and @data-actab-id="3"]')
        images_button.click()
        logging.info("点击了图片按钮")
        time.sleep(3)

        # 添加到商店
        add_to_store_button = driver.find_element(By.ID, 'addBtnSec')
        driver.execute_script("arguments[0].scrollIntoView(true);", add_to_store_button)
        add_to_store_button.click()
        logging.info("点击了添加到商店按钮")

        # 等待导入完成
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, 'importify-app-container'))
            )
            logging.info("产品正在导入中...")

            # 等待成功消息
            timeout = 100
            start_time = time.time()
            while time.time() - start_time < timeout:
                try:
                    success_message = driver.find_element(By.XPATH, '//div[@class="textcontainer centeralign home-content "]/p[1]')
                    if success_message.text == "We have successfully created the product page.":
                        logging.info(f"产品导入成功, 总数: {success_count + 1}")
                        success_count += 1
                        break
                except:
                    time.sleep(5)

        except Exception as e:
            logging.error(f"导入过程出错: {e}")

        return success_count

    except Exception as e:
        logging.error(f"处理产品操作时出错: {e}")
        return success_count

def fetch_dropdown_options(driver, sheet_name):
    try:
        # 如果sheet_name是列表，取第一个元素
        if isinstance(sheet_name, list):
            sheet_name = sheet_name[0]
            
        logging.info(f"处理下拉选项: {sheet_name}")

        # 等待下拉菜单加载
        dropdown = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.CLASS_NAME, 'ms-drop'))
        )

        # 查找并填写搜索框
        search_box = dropdown.find_element(By.CSS_SELECTOR, '.ms-search input[type="text"]')
        search_box.clear()
        search_box.send_keys(sheet_name.lower())
        
        # 等待搜索结果加载
        time.sleep(2)

        # 取消现有选择
        driver.execute_script("""
            var checkboxes = document.querySelectorAll('input[data-name="selectItem"]');
            checkboxes.forEach(function(checkbox) {
                if (checkbox.checked) {
                    checkbox.click();
                }
            });
        """)
        
        # 等待取消选择完成
        time.sleep(1)

        # 选择匹配的选项
        selected = False
        max_attempts = 3
        attempt = 0
        
        while not selected and attempt < max_attempts:
            try:
                # 获取所有可见的选项
                options = WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.ms-drop li:not(.hide) span'))
                )
                
                # 遍历选项寻找最佳匹配
                target_text = sheet_name.lower().strip()
                for option in options:
                    option_text = option.text.lower().strip()
                    if option_text == target_text or target_text in option_text:
                        # 找到匹配的选项，点击对应的checkbox
                        checkbox = option.find_element(By.XPATH, './preceding-sibling::input[@type="checkbox"]')
                        driver.execute_script("arguments[0].click();", checkbox)
                        selected = True
                        logging.info(f"成功选择类别: {option_text}")
                        
                        # 直接点击描述标签，不关闭下拉菜单
                        description_tab = WebDriverWait(driver, 5).until(
                            EC.element_to_be_clickable((By.XPATH, '//*[@id="description_tab_button"]'))
                        )
                        description_tab.click()
                        logging.info("直接进入描述标签")
                        break
                
                if not selected:
                    attempt += 1
                    if attempt == max_attempts:
                        logging.error(f"无法找到匹配的类别: {sheet_name}")
                        # 记录所有可用选项供调试
                        available_options = [opt.text for opt in options]
                        logging.info(f"可用选项: {available_options}")
                    time.sleep(2)
                    
            except Exception as e:
                attempt += 1
                if attempt == max_attempts:
                    logging.error(f"选择类别时出错: {str(e)}")
                time.sleep(2)

        if not selected:
            raise Exception("未能成功选择类别")

    except Exception as e:
        logging.error(f"处理下拉选项时出错: {e}")
        raise  # 抛出异常以便上层函数知道下拉选项处理失败
