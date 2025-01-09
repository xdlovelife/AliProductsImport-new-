import sys
import logging
from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from utils import (
    read_categories_from_excel, 
    read_sheet_names_from_excel, 
    open_browser,
    process_link
)
import os
import ctypes
import time

class ImportifyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.worker = None
        self.thread = None
        
        # 设置日志处理
        self.log_handler = QTextEditLogger(self.log_text)
        logging.getLogger().addHandler(self.log_handler)
        logging.getLogger().setLevel(logging.INFO)
        
        # 加载上次的Excel文件路径
        self.load_last_excel_path()

    def init_ui(self):
        self.setWindowTitle("阿里巴巴产品导入工具")
        self.setMinimumSize(800, 600)  # 设置最小窗口大小
        
        # 设置应用图标
        icon = QIcon("xdlovelife.ico")
        self.setWindowIcon(icon)
        
        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(10, 10, 10, 10)  # 设置边距
        layout.setSpacing(10)  # 设置间距
        
        # 创建菜单栏
        self.create_menu_bar()
        
        # 创建主要区域
        self.create_main_area(layout)
        
        # 创建状态栏
        self.statusBar().showMessage('就绪')

    def create_menu_bar(self):
        menubar = self.menuBar()
        
        # 文件菜单
        file_menu = menubar.addMenu('文件')
        open_action = QAction('打开Excel', self)
        open_action.setShortcut('Ctrl+O')
        open_action.triggered.connect(self.browse_file)
        file_menu.addAction(open_action)
        
        # 设置菜单
        settings_menu = menubar.addMenu('设置')
        config_action = QAction('配置', self)
        config_action.triggered.connect(self.show_settings)
        settings_menu.addAction(config_action)

    def create_main_area(self, layout):
        # 文件选择区域
        file_group = QGroupBox("Excel文件选择")
        file_layout = QHBoxLayout()
        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("请选择Excel文件...")
        browse_button = QPushButton("浏览")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path)
        file_layout.addWidget(browse_button)
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # 创建分割器
        splitter = QSplitter(Qt.Orientation.Vertical)
        layout.addWidget(splitter)
        
        # 创建选项卡
        self.tabs = QTabWidget()
        
        # 日志标签页
        self.log_tab = QWidget()
        log_layout = QVBoxLayout()
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        # 设置最小高度，允许调节大小
        self.log_text.setMinimumHeight(200)
        log_layout.addWidget(self.log_text)
        self.log_tab.setLayout(log_layout)
        
        # 数据预览标签页
        self.preview_tab = QWidget()
        preview_layout = QVBoxLayout()
        self.preview_table = QTableWidget()
        # 设置表格自适应大小
        self.preview_table.horizontalHeader().setStretchLastSection(True)
        self.preview_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        preview_layout.addWidget(self.preview_table)
        self.preview_tab.setLayout(preview_layout)
        
        # 添加标签页
        self.tabs.addTab(self.log_tab, "运行日志")
        self.tabs.addTab(self.preview_tab, "数据预览")
        splitter.addWidget(self.tabs)
        
        # 进度和控制区域
        control_group = QGroupBox("控制面板")
        control_layout = QVBoxLayout()
        
        # 进度信息
        progress_info = QHBoxLayout()
        self.progress = QProgressBar()
        self.progress_label = QLabel("0/0")
        progress_info.addWidget(self.progress, stretch=4)
        progress_info.addWidget(self.progress_label, stretch=1)
        control_layout.addLayout(progress_info)
        
        # 控制按钮
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("开始导入")
        self.pause_button = QPushButton("暂停")
        self.pause_button.setEnabled(False)
        
        self.start_button.clicked.connect(self.start_import)
        self.pause_button.clicked.connect(self.toggle_pause)
        
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.pause_button)
        control_layout.addLayout(button_layout)
        
        control_group.setLayout(control_layout)
        splitter.addWidget(control_group)
        
        # 设置分割器的初始大小
        splitter.setSizes([600, 100])
        
        # 设置分割器可以调整大小
        splitter.setHandleWidth(5)
        splitter.setChildrenCollapsible(False)

    def show_settings(self):
        dialog = SettingsDialog(self)
        dialog.exec()

    def browse_file(self):
        settings = QSettings('ImportifyApp', 'Settings')
        last_directory = os.path.dirname(settings.value('last_excel_path', ''))
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel文件",
            last_directory,  # 从上次的目录开始
            "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path.setText(file_path)
            self.save_excel_path(file_path)  # 保存新的文件路径
            self.load_preview_data(file_path)

    def load_preview_data(self, file_path):
        try:
            categories = read_categories_from_excel(file_path)
            sheet_names = read_sheet_names_from_excel(file_path)
            
            # 更新预览表格
            self.preview_table.setRowCount(len(categories))
            self.preview_table.setColumnCount(2)
            self.preview_table.setHorizontalHeaderLabels(["产品类别", "目标分类"])
            
            for i, category in enumerate(categories):
                self.preview_table.setItem(i, 0, QTableWidgetItem(category))
                if i < len(sheet_names):
                    self.preview_table.setItem(i, 1, QTableWidgetItem(sheet_names[i]))
            
            self.preview_table.resizeColumnsToContents()
            
        except Exception as e:
            logging.error(f"加载预览数据时出错: {str(e)}")

    def start_import(self):
        if not self.file_path.text():
            QMessageBox.warning(self, "警告", "请先选择Excel文件")
            return
            
        self.start_button.setEnabled(False)
        self.pause_button.setEnabled(True)
        self.pause_button.setText("暂停")
        
        # 创建工作线程
        self.thread = QThread()
        self.worker = ImportWorker(self.file_path.text())
        self.worker.moveToThread(self.thread)
        
        # 连接信号
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.thread.finished.connect(self.on_import_finished)
        
        self.worker.progress.connect(self.update_progress)
        self.worker.total_updated.connect(self.update_total)
        self.worker.status_changed.connect(self.update_pause_button)
        
        # 启动线程
        self.thread.start()

    def toggle_pause(self):
        if self.worker:
            if not self.worker.is_paused:
                self.worker.pause()
                self.pause_button.setText("继续")
            else:
                self.worker.resume()
                self.pause_button.setText("暂停")

    def update_pause_button(self, is_paused):
        self.pause_button.setText("继续" if is_paused else "暂停")

    def update_progress(self, current):
        total = self.progress.maximum()
        self.progress.setValue(current)
        self.progress_label.setText(f"{current}/{total}")

    def update_total(self, total):
        self.progress.setMaximum(total)
        self.progress_label.setText(f"0/{total}")

    def on_import_finished(self):
        self.start_button.setEnabled(True)
        self.pause_button.setEnabled(False)
        self.pause_button.setText("暂停")
        self.progress.setValue(0)
        self.progress_label.setText("0/0")
        logging.info("导入任务完成")

    def closeEvent(self, event):
        """处理窗口关闭事件"""
        if self.thread and self.thread.isRunning():
            self.worker.is_running = False
            self.thread.wait()  # 等待线程结束
        event.accept()

    def load_last_excel_path(self):
        """加载上次使用的Excel文件路径"""
        settings = QSettings('ImportifyApp', 'Settings')
        last_excel_path = settings.value('last_excel_path', '')
        if last_excel_path and os.path.exists(last_excel_path):
            self.file_path.setText(last_excel_path)
            self.load_preview_data(last_excel_path)

    def save_excel_path(self, file_path):
        """保存Excel文件路径到设置"""
        settings = QSettings('ImportifyApp', 'Settings')
        settings.setValue('last_excel_path', file_path)

class ImportWorker(QObject):
    progress = pyqtSignal(int)
    total_updated = pyqtSignal(int)
    finished = pyqtSignal()
    status_changed = pyqtSignal(bool)  # True表示暂停，False表示继续

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path
        self.is_running = True
        self.is_paused = False

    def pause(self):
        self.is_paused = True
        # 设置全局暂停状态
        from utils import process_link
        process_link.is_paused = True
        self.status_changed.emit(True)
        logging.info("导入任务已暂停")

    def resume(self):
        self.is_paused = False
        # 清除全局暂停状态
        from utils import process_link
        process_link.is_paused = False
        self.status_changed.emit(False)
        logging.info("导入任务继续进行")

    def run(self):
        try:
            # 读取设置
            settings = QSettings('ImportifyApp', 'Settings')
            driver_path = settings.value('driver_path', '')
            user_data_dir = settings.value('user_data_dir', '')
            wait_time = int(settings.value('wait_time', 10))

            # 读取Excel文件
            categories = read_categories_from_excel(self.file_path)
            sheet_names = read_sheet_names_from_excel(self.file_path)
            
            if not categories:
                logging.error("没有找到有效的类别")
                return
                
            if not sheet_names:
                logging.error("没有找到有效的目标分类")
                return
                
            logging.info(f"总共读取到 {len(categories)} 个类别")
            logging.info(f"类别列表: {categories}")
            logging.info(f"目标分类列表: {sheet_names}")
            
            # 使用第一个sheet_name作为所有类别的目标分类
            target_sheet_name = sheet_names[0]
            logging.info(f"使用 '{target_sheet_name}' 作为所有类别的目标分类")
            
            total = len(categories)
            self.total_updated.emit(total)
            total_success_count = 0
            
            # 创建浏览器实例
            driver = None
            max_browser_retries = 3
            for retry in range(max_browser_retries):
                try:
                    driver = open_browser(driver_path, user_data_dir)
                    if driver:
                        break
                except Exception as e:
                    logging.error(f"创建浏览器实例失败 (尝试 {retry + 1}/{max_browser_retries}): {str(e)}")
                    if retry == max_browser_retries - 1:
                        raise
                    time.sleep(2)
            
            if driver is None:
                logging.error("无法创建浏览器实例")
                return
            
            try:
                url = "https://www.alibaba.com/"
                logging.info(f"访问页面: {url}")
                driver.get(url)
                
                search_bar = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CLASS_NAME, 'fy23-icbu-search-bar-inner'))
                )

                # 遍历处理每个类别
                for index, category in enumerate(categories, 1):
                    if not self.is_running:
                        logging.info("任务被用户停止")
                        break

                    # 检查是否暂停
                    while self.is_paused:
                        time.sleep(1)
                        if not self.is_running:  # 如果在暂停时被停止
                            break
                    
                    if not self.is_running:  # 再次检查是否被停止
                        break
                    
                    logging.info(f"正在处理第 {index}/{len(categories)} 个类别: {category}")
                    
                    try:
                        success_count = process_link(driver, category, target_sheet_name)
                        if success_count > 0:
                            total_success_count += success_count
                            logging.info(f"成功处理类别 '{category}'，当前总成功数: {total_success_count}")
                        else:
                            logging.warning(f"处理类别 '{category}' 未成功导入任何产品")
                        
                        self.progress.emit(index)
                        
                    except Exception as e:
                        logging.error(f"处理类别 '{category}' 出错: {e}")
                        # 尝试重新初始化浏览器
                        try:
                            driver.quit()
                        except:
                            pass
                            
                        driver = open_browser(driver_path, user_data_dir)
                        if not driver:
                            logging.error("重新创建浏览器实例失败，终止处理")
                            break
                            
                        driver.get(url)
                        continue

                # 输出最终统计信息
                logging.info(f"处理完成，共处理 {len(categories)} 个类别，成功导入 {total_success_count} 个产品")

            finally:
                try:
                    driver.quit()  # 确保浏览器被关闭
                except:
                    pass

        except Exception as e:
            logging.error(f"导入过程出错: {str(e)}")
        finally:
            self.finished.emit()

class SettingsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.init_ui()
        self.load_settings()

    def get_default_user_data_dir(self):
        """获取Chrome默认用户数据目录"""
        if os.name == 'nt':  # Windows
            return os.path.join(os.environ['LOCALAPPDATA'], 'Google', 'Chrome', 'User Data')
        elif os.name == 'posix':  # macOS
            return os.path.expanduser('~/Library/Application Support/Google/Chrome')
        else:  # Linux
            return os.path.expanduser('~/.config/google-chrome')

    def init_ui(self):
        self.setWindowTitle("设置")
        layout = QFormLayout()
        
        # Chrome驱动路径组
        driver_group = QGroupBox("ChromeDriver设置")
        driver_layout = QVBoxLayout()
        
        # 自动下载选项
        self.auto_download = QCheckBox("自动下载ChromeDriver")
        self.auto_download.stateChanged.connect(self.toggle_driver_path)
        driver_layout.addWidget(self.auto_download)
        
        # 手动选择路径
        path_layout = QHBoxLayout()
        self.driver_path = QLineEdit()
        self.driver_path.setPlaceholderText("请选择ChromeDriver路径...")
        browse_button = QPushButton("浏览")
        browse_button.clicked.connect(self.browse_driver)
        path_layout.addWidget(self.driver_path)
        path_layout.addWidget(browse_button)
        driver_layout.addLayout(path_layout)
        
        driver_group.setLayout(driver_layout)
        layout.addRow(driver_group)
        
        # 用户数据目录组
        user_data_group = QGroupBox("Chrome用户数据目录")
        user_data_layout = QVBoxLayout()
        
        # 使用默认目录选项
        self.use_default_dir = QCheckBox("使用默认用户数据目录")
        self.use_default_dir.stateChanged.connect(self.toggle_user_data_dir)
        user_data_layout.addWidget(self.use_default_dir)
        
        # 手动选择用户数据目录
        user_dir_layout = QHBoxLayout()
        self.user_data_dir = QLineEdit()
        self.user_data_dir.setPlaceholderText("请选择用户数据目录...")
        browse_user_dir_button = QPushButton("浏览")
        browse_user_dir_button.clicked.connect(self.browse_user_dir)
        user_dir_layout.addWidget(self.user_data_dir)
        user_dir_layout.addWidget(browse_user_dir_button)
        user_data_layout.addLayout(user_dir_layout)
        
        user_data_group.setLayout(user_data_layout)
        layout.addRow(user_data_group)
        
        # 等待时间
        self.wait_time = QSpinBox()
        self.wait_time.setRange(1, 60)
        self.wait_time.setValue(10)  # 默认值
        layout.addRow("等待时间(秒):", self.wait_time)
        
        # 下载提示
        tip_label = QLabel("提示：如果自动下载失败，请手动下载ChromeDriver并指定路径")
        tip_label.setStyleSheet("color: gray;")
        layout.addRow(tip_label)
        
        # 按钮
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | 
            QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self.save_settings)
        buttons.rejected.connect(self.reject)
        layout.addRow(buttons)
        
        self.setLayout(layout)

    def toggle_driver_path(self, state):
        """切换ChromeDriver路径输入状态"""
        self.driver_path.setEnabled(not state)
        if state:
            self.driver_path.clear()

    def browse_driver(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择ChromeDriver",
            "",
            "ChromeDriver (chromedriver.exe);;所有文件 (*.*)"
        )
        if file_path:
            self.driver_path.setText(file_path)
            self.auto_download.setChecked(False)

    def browse_user_dir(self):
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "选择用户数据目录"
        )
        if dir_path:
            self.user_data_dir.setText(dir_path)
            self.use_default_dir.setChecked(False)

    def toggle_user_data_dir(self, state):
        """切换用户数据目录输入状态"""
        self.user_data_dir.setEnabled(not state)
        if state:
            self.user_data_dir.setText(self.get_default_user_data_dir())
        else:
            self.user_data_dir.clear()

    def load_settings(self):
        settings = QSettings('ImportifyApp', 'Settings')
        self.auto_download.setChecked(settings.value('auto_download', True, type=bool))
        self.driver_path.setText(settings.value('driver_path', ''))
        
        # 加载用户数据目录设置
        use_default = settings.value('use_default_dir', True, type=bool)
        self.use_default_dir.setChecked(use_default)
        if use_default:
            self.user_data_dir.setText(self.get_default_user_data_dir())
        else:
            self.user_data_dir.setText(settings.value('user_data_dir', ''))
        
        self.wait_time.setValue(int(settings.value('wait_time', 10)))
        self.toggle_driver_path(self.auto_download.isChecked())
        self.toggle_user_data_dir(self.use_default_dir.isChecked())

    def save_settings(self):
        settings = QSettings('ImportifyApp', 'Settings')
        settings.setValue('auto_download', self.auto_download.isChecked())
        settings.setValue('driver_path', self.driver_path.text())
        settings.setValue('use_default_dir', self.use_default_dir.isChecked())
        settings.setValue('user_data_dir', self.user_data_dir.text())
        settings.setValue('wait_time', self.wait_time.value())
        self.accept()

class QTextEditLogger(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget
        self.widget.setReadOnly(True)
        self.setFormatter(logging.Formatter('%(asctime)s || %(message)s'))
        
        # 创建定时器用于更新日志
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_log)
        self.update_timer.start(100)  # 每100毫秒更新一次
        self.pending_records = []

    def emit(self, record):
        msg = self.format(record)
        # 将消息添加到待处理列表
        self.pending_records.append(msg)

    def update_log(self):
        if self.pending_records:
            # 批量处理所有待处理的记录
            for msg in self.pending_records:
                self.widget.append(msg)
            
            # 强制滚动到底部
            self.widget.verticalScrollBar().setValue(
                self.widget.verticalScrollBar().maximum()
            )
            
            # 清空待处理列表
            self.pending_records.clear()

def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    app_icon = QIcon("xdlovelife.ico")
    app.setWindowIcon(app_icon)
    
    # 设置Windows任务栏图标
    if os.name == 'nt':  # Windows系统
        myappid = 'mycompany.myproduct.subproduct.version'  # 任意字符串
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    
    # 设置应用样式
    app.setStyle("Fusion")
    
    window = ImportifyApp()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()

