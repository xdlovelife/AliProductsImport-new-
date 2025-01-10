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
        df = pd.read_excel(file_path, header=None)  # 不使用第一行作为表头
        # 获取第一列的所有值，包括第一行
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
        # 优先使用指定的ChromeDriver路径
        default_driver_path = r'D:\chromedriver-win64\chromedriver.exe'
        chrome_driver_path = driver_path if driver_path else default_driver_path
        
        if not os.path.exists(chrome_driver_path):
            logging.error(f"ChromeDriver不存在于路径: {chrome_driver_path}")
            return None
            
        # 设置Chrome选项
        chrome_options = Options()
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-features=TranslateUI')
        chrome_options.add_argument('--disable-translate')
        chrome_options.add_argument('--lang=en-US')
        chrome_options.add_argument('--remote-allow-origins=*')
        
        # 如果提供了用户数据目录，则添加相应选项
        if user_data_dir and os.path.exists(user_data_dir):
            chrome_options.add_argument(f'--user-data-dir={user_data_dir}')
            
        # 创建Service对象
        service = Service(executable_path=chrome_driver_path)
        
        # 创建WebDriver实例
        driver = webdriver.Chrome(service=service, options=chrome_options)
        logging.info("成功创建Chrome浏览器实例")
        return driver
        
    except Exception as e:
        logging.error(f"创建浏览器实例失败: {str(e)}")
        return None

def process_link(driver, category, sheet_name):
    max_retries = 3
    retry_count = 0
    
    logging.info(f"开始处理类别: {category}, 目标分类: {sheet_name}")
    
    while retry_count < max_retries:
        try:
            # 检查driver是否还有效
            try:
                driver.current_url
            except:
                logging.error("浏览器已关闭，需要重新初始化")
                # 检查是否处于暂停状态
                if hasattr(process_link, 'is_paused') and process_link.is_paused:
                    logging.info("系统处于暂停状态，等待继续...")
                    while process_link.is_paused:
                        time.sleep(1)
                driver = open_browser()
                if not driver:
                    return 0
            
            # 验证sheet_name
            if not sheet_name:
                logging.error(f"类别 '{category}' 没有对应的目标分类")
                return 0
            
            # 访问阿里巴巴主页
            url = "https://www.alibaba.com/"
            logging.info(f"访问页面: {url}")
            driver.get(url)
            
            # 等待页面加载完成
            try:
                WebDriverWait(driver, 20).until(
                    lambda d: d.execute_script('return document.readyState') == 'complete'
                )
                logging.info("页面加载完成")
            except TimeoutException:
                logging.error("页面加载超时")
                return 0
                
            driver.switch_to.window(driver.window_handles[0])
            
            # 等待主要元素出现
            try:
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'fy23-icbu-search-bar-inner'))
                )
                logging.info("搜索栏加载完成")
            except TimeoutException:
                logging.error("搜索栏加载超时")
                return 0
            
            # 处理搜索和产品列表
            try:
                # 等待搜索框加载
                logging.info("等待搜索框加载...")
                search_input = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'input.search-bar-input.util-ellipsis'))
                )
                
                # 确保搜索框可以交互
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, 'input.search-bar-input.util-ellipsis'))
                )
                
                # 清除搜索框并输入
                search_input.clear()
                time.sleep(1)  # 等待清除完成
                search_input.send_keys(category)
                time.sleep(1)  # 等待输入完成
                
                # 验证输入是否成功
                if search_input.get_attribute('value') != category:
                    logging.error(f"搜索关键词输入失败，预期: {category}, 实际: {search_input.get_attribute('value')}")
                    return 0
                    
                logging.info(f"输入搜索关键词: {category}")
                
                # 点击搜索按钮
                search_button = driver.find_element(By.CSS_SELECTOR, 'button.fy23-icbu-search-bar-inner-button')
                search_button.click()
                logging.info("点击搜索按钮")
                
                # 等待搜索结果加载
                try:
                    # 等待加载动画消失
                    WebDriverWait(driver, 10).until_not(
                        EC.presence_of_element_located((By.CLASS_NAME, "loading-mask"))
                    )
                    logging.info("搜索结果加载中...")
                    
                    # 等待搜索结果出现
                    WebDriverWait(driver, 20).until(
                        EC.presence_of_element_located((By.CLASS_NAME, "organic-list"))
                    )
                    logging.info("搜索结果加载完成")
                    
                    # 等待一下确保所有元素都加载完成
                    time.sleep(3)
                    
                    # 验证是否在搜索结果页面
                    current_url = driver.current_url
                    if "alibaba.com/trade/search" not in current_url:
                        logging.error("未能正确跳转到搜索结果页面")
                        return 0
                        
                except TimeoutException:
                    logging.error("等待搜索结果超时")
                    return 0
                
                # 滚动加载所有产品
                logging.info("开始滚动加载更多产品...")
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
                    logging.info(f"完成第 {scroll_attempts} 次滚动")
                
                # 获取产品列表
                product_list = driver.find_elements(By.CLASS_NAME, "fy23-search-card")
                num_products = len(product_list)
                logging.info(f"找到 {num_products} 个产品")
                
                if num_products == 0:
                    logging.warning(f"类别 '{category}' 没有找到任何产品")
                    return 0
                
                # 处理每个产品
                success_count = 0
                
                for product in product_list:
                    try:
                        # 获取产品标题和URL
                        product_title = product.find_element(By.CLASS_NAME, "search-card-e-title")
                        product_url = product.find_element(By.TAG_NAME, "a").get_attribute("href")
                        
                        logging.info(f"当前产品标题: {product_title.text}")
                        
                        # 打开新窗口
                        driver.execute_script(f"window.open('{product_url}')")
                        
                        # 获取所有窗口句柄
                        original_window = driver.current_window_handle
                        new_window = [handle for handle in driver.window_handles if handle != original_window][0]
                        
                        # 切换到新窗口
                        driver.switch_to.window(new_window)
                        
                        # 等待产品详情页加载
                        product_title = WebDriverWait(driver, 60).until(
                            EC.presence_of_element_located((By.TAG_NAME, "h1"))
                        )
                        
                        # 处理产品详情页操作
                        current_sheet_name = sheet_name
                        if isinstance(sheet_name, list):
                            if sheet_name:  # 确保列表不为空
                                current_sheet_name = sheet_name[0]
                            else:
                                logging.error("sheet_name列表为空")
                                current_sheet_name = None
                                
                        result = handle_product_actions(driver, category, success_count, current_sheet_name)
                        
                        if result > success_count:
                            success_count = result
                        
                        # 关闭新窗口并切回原窗口
                        try:
                            # 直接切换回原窗口
                            driver.switch_to.window(original_window)
                            
                            # 关闭其他所有窗口，只保留原窗口
                            for handle in driver.window_handles:
                                if handle != original_window:
                                    driver.switch_to.window(handle)
                                    driver.close()
                            
                            # 最后确保回到原窗口
                            driver.switch_to.window(original_window)
                            
                        except Exception as e:
                            logging.error(f"处理窗口关闭时出错: {str(e)}")
                            # 如果出错，尝试切换到任何可用窗口
                            try:
                                if driver.window_handles:
                                    driver.switch_to.window(driver.window_handles[0])
                            except:
                                pass
                            continue
                        
                    except Exception as e:
                        logging.error(f"处理产品时出错: {str(e)}")
                        try:
                            # 检查当前窗口是否还存在
                            try:
                                driver.current_url
                                driver.close()
                            except:
                                pass
                            
                            # 尝试切换回原窗口
                            try:
                                driver.switch_to.window(original_window)
                            except:
                                # 如果切换失败，尝试切换到任何可用窗口
                                available_windows = driver.window_handles
                                if available_windows:
                                    driver.switch_to.window(available_windows[0])
                            
                        except Exception as close_error:
                            logging.error(f"清理窗口时出错: {close_error}")
                        continue
                
                return success_count
                
            except Exception as e:
                logging.error(f"处理产品列表时出错: {str(e)}")
                retry_count += 1
                if retry_count >= max_retries:
                    return 0
                time.sleep(2)
                continue
                
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

def handle_product_actions(driver, category, success_count, sheet_name):
    try:
        logging.info(f"处理产品详情页操作: {category}, {sheet_name}")
        
        # 检查窗口是否存在
        def check_window():
            try:
                driver.current_url
                return True
            except:
                logging.error("窗口已关闭，需要重新处理")
                return False

        # 点击添加按钮
        max_retries = 3
        for retry in range(max_retries):
            try:
                if not check_window():
                    return success_count
                    
                add_btn_con = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="addBtnCon"]')))
                    
                # 再次检查窗口
                if not check_window():
                    return success_count
                    
                add_btn_con.click()
                logging.info("点击了添加按钮")
                break
            except Exception as e:
                if retry == max_retries - 1:
                    logging.error(f"点击添加按钮失败: {str(e)}")
                    return success_count
                time.sleep(2)
                
                # 检查窗口是否还存在
                if not check_window():
                    return success_count

        # 处理Draft元素
        try:
            if not check_window():
                return success_count
                
            # 等待Draft元素可见和可点击，减少等待时间
            draft_element = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, '//span[@class="inactive" and text()="Draft"]'))
            )
            
            if not check_window():
                return success_count
                
            WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.XPATH, '//span[@class="inactive" and text()="Draft"]'))
            )
            logging.info("成功加载 Draft 元素")
            
            # 再次检查窗口
            if not check_window():
                return success_count
                
            # 使用JavaScript点击，更可靠
            driver.execute_script("arguments[0].click();", draft_element)
            logging.info("成功点击 Draft 元素")
            time.sleep(1)  # 减少等待时间
        except Exception as e:
            logging.error(f"等待和点击 Draft 元素时出错：{e}")
            return success_count

        # 检查区域限制
        try:
            # 减少等待时间到3秒，因为区域限制信息通常会很快显示
            region_restriction = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(text(), "Sorry, this product can\'t be shipped to your region.")]'))
            )
            if region_restriction.is_displayed():
                logging.info("检测到产品无法配送到当前区域，跳过处理")
                return success_count
        except TimeoutException:
            logging.info("未检测到区域限制消息，继续处理")

        # 检查产品是否已存在（同样减少等待时间）
        try:
            WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, '//div[@class="textcontainer centeralign home-content "]'))
            )
            success_message = driver.find_element(By.XPATH, '//div[@class="textcontainer centeralign home-content "]/p[1]')
            if success_message.text == "This product is already in your store, what would you like to do?":
                logging.info("产品已存在，不再处理")
                return success_count
        except (TimeoutException, NoSuchElementException):
            pass

        # 检查sheet_name是否有效
        if not sheet_name:
            logging.error("没有有效的目标分类，跳过类别选择")
            return success_count

        # 选择类别
        try:
            # 首先检查当前选择的类别
            try:
                current_selection = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'button.ms-choice span'))
                )
                if current_selection.text.strip() == sheet_name:
                    logging.info(f"当前已选择正确的类别: {sheet_name}，直接进入下一步")
                    # 直接点击描述标签
                    try:
                        description_tab = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.XPATH, '//*[@id="description_tab_button"]'))
                        )
                        time.sleep(1)  # 等待一下确保元素完全加载
                        
                        # 使用JavaScript点击
                        driver.execute_script("arguments[0].click();", description_tab)
                        logging.info("已点击描述标签")
                        
                        # 等待一下确保切换完成
                        time.sleep(2)
                    except Exception as e:
                        logging.error(f"点击描述标签失败: {str(e)}")
                        return success_count
                else:
                    # 等待并点击选择按钮
                    for retry in range(max_retries):
                        try:
                            select_button = WebDriverWait(driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, '//button[@class="ms-choice"]'))
                            )
                            select_button.click()
                            logging.info("等待并点击选择按钮")
                            break
                        except Exception as e:
                            if retry == max_retries - 1:
                                raise
                            time.sleep(2)

                    # 处理下拉选项并等待完成
                    try:
                        fetch_dropdown_options(driver, sheet_name)
                    except ValueError as ve:
                        logging.error(f"选择类别失败: {ve}")
                        return success_count
                    except Exception as e:
                        logging.error(f"选择类别失败: {e}")
                        return success_count

                    time.sleep(3)
            except Exception as e:
                logging.error(f"检查当前选择时出错: {e}")
                return success_count

            # 处理变体
            try:
                # 首先点击Variants按钮
                try:
                    variants_button = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, '//button[@class="accordion-tab accordion-custom-tab" and @data-actab-group="0" and @data-actab-id="2"]'))
                    )
                    variants_button.click()
                    logging.info("点击了Variants按钮")
                    time.sleep(2)  # 等待变体面板展开
                except Exception as e:
                    logging.error(f"点击Variants按钮时出错：{e}")
                    return success_count

                # 点击"Select which variants to include"按钮
                try:
                    # 等待radio按钮可见和可点击
                    select_variants_radio = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'price_switch'))
                    )
                    
                    # 确保按钮在视图中
                    driver.execute_script("arguments[0].scrollIntoView(true);", select_variants_radio)
                    time.sleep(1)  # 等待滚动完成
                    
                    # 使用JavaScript点击radio按钮
                    driver.execute_script("arguments[0].click();", select_variants_radio)
                    logging.info("选择了'Select which variants to include'选项")
                    time.sleep(2)  # 等待变体选择界面加载
                except Exception as e:
                    logging.error(f"选择变体选项时出错：{e}")
                    return success_count

                # 处理变体选择
                try:
                    # 等待变体表格完全加载
                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'var_price'))
                    )
                    time.sleep(2)  # 等待表格完全渲染

                    # 只取消第一个之外的其他变体
                    try:
                        variant_checkboxes = driver.find_elements(By.CSS_SELECTOR, '.include_variant')
                        for i, checkbox in enumerate(variant_checkboxes):
                            if i > 0 and checkbox.is_selected():  # 跳过第一个变体，只取消其他选中的变体
                                driver.execute_script("arguments[0].click();", checkbox)
                                time.sleep(0.5)
                        logging.info("已取消第一个之外的其他变体")
                        
                        # 确保第一个变体被选中
                        first_variant = variant_checkboxes[0]
                        if not first_variant.is_selected():
                            driver.execute_script("arguments[0].click();", first_variant)
                            logging.info("已选中第一个变体")
                            
                    except Exception as e:
                        logging.error(f"处理变体选择时出错：{e}")

                except Exception as e:
                    logging.error(f"处理变体时出错：{e}")
                    return success_count

            except Exception as e:
                logging.error(f"处理变体时出错：{e}")
                return success_count

            # 处理图片
            try:
                images_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//button[@class="accordion-tab accordion-custom-tab" and @data-actab-group="0" and @data-actab-id="3"]'))
                )
                images_button.click()
                logging.info("点击了图片按钮")
                time.sleep(2)
            except Exception as e:
                logging.error(f"处理图片时出错：{e}")
                return success_count

            # 添加到商店
            try:
                add_to_store_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.ID, 'addBtnSec'))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", add_to_store_button)
                time.sleep(1)  # 等待滚动完成
                add_to_store_button.click()
                logging.info("点击了添加到商店按钮")

                # 等待导入完成
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'importify-app-container'))
                )
                logging.info("产品正在导入中...")

                # 等待成功消息
                success = False
                timeout = 100
                start_time = time.time()
                while time.time() - start_time < timeout:
                    try:
                        success_message = WebDriverWait(driver, 5).until(
                            EC.presence_of_element_located((By.XPATH, '//div[@class="textcontainer centeralign home-content "]/p[1]'))
                        )
                        if success_message.text == "We have successfully created the product page.":
                            success_count += 1
                            logging.info(f"产品导入成功, 总数: {success_count}")
                            success = True
                            break
                    except:
                        time.sleep(5)
                
                if not success:
                    logging.warning("等待成功消息超时")

            except Exception as e:
                logging.error(f"添加到商店时出错: {e}")

        except Exception as e:
            logging.error(f"选择类别按钮时出错: {e}")

        return success_count

    except Exception as e:
        logging.error(f"处理产品操作时出错: {e}")
        return success_count

def fetch_dropdown_options(driver, sheet_name):
    try:
        # 处理sheet_name为空列表或None的情况
        if not sheet_name:
            logging.error("sheet_name为空，无法选择类别")
            raise ValueError("sheet_name不能为空")
            
        # 如果sheet_name是列表，取第一个元素
        if isinstance(sheet_name, list):
            if not sheet_name:  # 如果是空列表
                logging.error("sheet_name列表为空，无法选择类别")
                raise ValueError("sheet_name列表不能为空")
            sheet_name = sheet_name[0]
            
        logging.info(f"处理下拉选项: {sheet_name}")

        try:
            # 等待下拉菜单完全加载
            dropdown = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'ms-drop'))
            )
            time.sleep(1)  # 额外等待以确保完全加载

            # 查找并填写搜索框
            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.ms-search input[type="text"]'))
            )
            
            # 清除搜索框并输入
            search_box.clear()
            time.sleep(0.5)
            search_box.send_keys(sheet_name.lower())
            time.sleep(2)  # 等待搜索结果

            # 选择匹配的选项
            target_text = sheet_name.lower().strip()
            
            # 获取所有可见的选项
            options = WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, '.ms-drop li:not(.hide) span'))
            )
            
            if not options:
                logging.warning(f"未找到任何匹配的选项: {sheet_name}")
                raise Exception("没有找到任何可选择的选项")

            # 找到完全匹配的选项
            for option in options:
                option_text = option.text.lower().strip()
                if option_text == target_text:
                    try:
                        checkbox = option.find_element(By.XPATH, './preceding-sibling::input[@type="checkbox"]')
                        if not checkbox.is_selected():
                            driver.execute_script("arguments[0].click();", checkbox)
                            logging.info(f"成功选择类别: {option.text}")
                            time.sleep(1)
                            break
                    except Exception as e:
                        logging.error(f"选择选项时出错: {str(e)}")
                        raise

            # 点击描述标签
            description_tab = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="description_tab_button"]'))
            )
            time.sleep(1)  # 等待一下确保元素完全加载
            
            # 使用JavaScript点击
            driver.execute_script("arguments[0].click();", description_tab)
            logging.info("已点击描述标签")
            
            # 等待一下确保切换完成
            time.sleep(2)

        except Exception as e:
            logging.error(f"处理下拉选项时出错: {e}")
            raise

    except Exception as e:
        logging.error(f"处理下拉选项时出错: {e}")
        raise  # 抛出异常以便上层函数知道下拉选项处理失败 