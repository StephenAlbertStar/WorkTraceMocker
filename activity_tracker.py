import subprocess
import time
import random
import os
import sys
import ctypes
import json
import threading
import datetime

# 尝试导入pyautogui和win32com，这些是exe环境中最容易出问题的模块
try:
    import pyautogui
    # 设置pyautogui的安全设置
    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 1  # 每个操作后暂停1秒
    PYAUTOGUI_AVAILABLE = True
except Exception as e:
    print(f"Warning: pyautogui import failed: {e}")
    PYAUTOGUI_AVAILABLE = False
    # 创建一个模拟的pyautogui对象
    class MockPyautogui:
        def getAllWindows(self): return []
        def press(self, key): pass
        def hotkey(self, *keys): pass
        def click(self, x=None, y=None): pass
    pyautogui = MockPyautogui()

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except Exception as e:
    print(f"Warning: win32com.client import failed: {e}")
    WIN32COM_AVAILABLE = False
    # 创建一个模拟的win32com对象
    class MockWin32com:
        class client:
            @staticmethod
            def Dispatch(app_name):
                class MockApp:
                    def Quit(self): pass
                return MockApp()
    win32com = MockWin32com()

from tkinter import Tk, filedialog, Label, Button, Frame, Entry, OptionMenu, StringVar, IntVar, DoubleVar, messagebox, Checkbutton, TclError
import logging


class ActivityTracker:
    def __init__(self, root):
        try:
            self.root = root
            self.root.title("WorkTrace Mocker")
            self.root.geometry("900x600")  # 调整窗口高度
            self.root.resizable(True, True)

            # 初始化随机数种子，确保真正的随机性
            random.seed()

            # 配置文件路径 - 增强错误处理
            try:
                if getattr(sys, 'frozen', False):
                    self.config_path = os.path.join(os.path.dirname(sys.executable), "config.json")
                else:
                    self.config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
            except Exception as e:
                # 如果配置路径获取失败，使用当前目录
                self.config_path = "config.json"
                print(f"Warning: Could not determine config path, using default: {e}")

            # 首先初始化所有基本属性
            try:
                self.initialize_basic_attributes()
            except Exception as e:
                print(f"Error initializing basic attributes: {e}")
                raise

            # 自动创建配置文件
            try:
                self.ensure_config_exists()
            except Exception as e:
                print(f"Error ensuring config exists: {e}")
                # 继续执行，使用默认配置

            # 加载配置文件
            try:
                self.load_config()
            except Exception as e:
                print(f"Error loading config: {e}")
                # 继续执行，使用默认配置

            # 初始化日志系统
            try:
                self.setup_logging()

                # 记录程序启动信息（仅在日志启用时）
                if self.logger:
                    import platform
                    self.log_info("程序启动", f"WorkTrace Mocker 启动")
                    self.log_info("运行环境", f"Python版本: {sys.version}")
                    self.log_info("运行环境", f"操作系统: {platform.system()} {platform.release()}")
                    self.log_info("运行环境", f"是否为exe: {getattr(sys, 'frozen', False)}")
                    self.log_info("运行环境", f"pyautogui可用: {PYAUTOGUI_AVAILABLE}")
                    self.log_info("运行环境", f"win32com可用: {WIN32COM_AVAILABLE}")
                    self.log_info("配置文件", f"配置路径: {self.config_path}")

            except Exception as e:
                print(f"Error setting up logging: {e}")
                # 继续执行，不使用日志

            # 创建界面
            try:
                self.create_widgets()
            except Exception as e:
                print(f"Error creating widgets: {e}")
                raise

        except Exception as e:
            print(f"Critical error in ActivityTracker initialization: {e}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            raise

        # 在界面创建完成后立即更新时间显示
        self.root.after(100, self.update_save_time)  # 延迟100毫秒确保界面完全创建

        # 启动心跳监控（开发环境和exe环境均启用，便于测试）
        self.start_heartbeat_monitor()

    def initialize_basic_attributes(self):
        """初始化所有基本属性"""
        # 项目文件夹路径变量列表
        self.project_folders = []  # 存储多个项目文件夹路径
        self.folder_vars = []  # 初始化文件夹变量列表
        self.folder_widgets = []  # 初始化文件夹组件列表

        # 工作时间设置变量
        self.work_start_hour = IntVar(value=9)    # 开始工作时间（小时）
        self.work_start_minute = IntVar(value=0)   # 开始工作时间（分钟）
        self.work_end_hour = IntVar(value=18)      # 结束工作时间（小时）
        self.work_end_minute = IntVar(value=00)    # 结束工作时间（分钟）

        # 午休时间设置变量（仅在配置文件中配置，不在界面显示）
        self.lunch_break_enabled = True          # 是否启用午休功能
        self.lunch_start_hour = 12               # 午休开始时间（小时）
        self.lunch_start_minute = 0              # 午休开始时间（分钟）
        self.lunch_end_hour = 13                 # 午休结束时间（小时）
        self.lunch_end_minute = 30               # 午休结束时间（分钟）
        self.lunch_time_random_range = 5        # 午休时间随机区间（分钟）

        # 工作时间随机区间设置变量（分钟）
        self.work_time_random_range = 20  # 工作时间随机区间（分钟）

        # 保存延迟设置变量
        self.save_delay_min = DoubleVar(value=20)      # 最小保存延迟（分钟）
        self.save_delay_max = DoubleVar(value=60)       # 最大保存延迟（分钟）

        # 文件打开间隔设置变量
        self.file_interval_min = DoubleVar(value=30.0)  # 最小间隔（分钟）
        self.file_interval_max = DoubleVar(value=60.0)  # 最大间隔（分钟）

        # 线程控制变量
        self.task_thread = None
        self.running = False
        self.cancel_event = threading.Event()

        # 存储当前周期的实际保存时间（包含随机延迟）
        self.actual_save_time = None

        # 当前打开的文件（用于跟踪只保存当前文件）
        self.current_opened_file = None
        self.current_opened_time = None

        # 当日工作时间缓存（确保每天的工作时间只计算一次）
        self.daily_work_times = {}  # 格式: {日期字符串: (工作开始时间, 工作结束时间)}

        # 预定的文件打开时间（确保预计时间与实际时间一致）
        self.next_file_open_time = None

        # 缓存的时间设置（用于检测时间设置变化）
        self.cached_time_settings = {}

        # 防抖计时器（避免频繁的时间设置变化回调）
        self.time_change_timer = None        # 跟踪打开的文件和对应的软件程序
        self.opened_files = []  # 存储已打开的文件路径
        self.opened_programs = set()  # 存储已启动的程序名称

        # 智能关闭功能配置变量
        self.auto_close_on_work_end = True
        self.use_alt_f4 = True
        self.use_ctrl_q = True
        self.use_ctrl_w = True
        self.close_timeout = 3.0

        # 文件跟踪配置变量
        self.file_tracking_enabled = True
        self.track_program_mapping = True
        self.clear_tracking_on_stop = True

        # 程序检测配置变量
        self.window_check_interval = 1.0
        self.activation_delay = 0.5
        self.close_verification_delay = 1.5

        # 用户界面配置变量
        self.show_close_progress = True
        self.show_detected_programs = True
        self.status_update_interval = 1.0

        # 工作日历配置变量
        self.skip_weekends = True
        self.work_dates = []  # 调休工作日期列表 (格式: "MM-DD"每年生效 或 "YY-MM-DD"指定年份)
        self.holiday_dates = []  # 节假日日期列表 (格式: "MM-DD"每年生效 或 "YY-MM-DD"指定年份)

        # 日志功能配置变量
        self.logging_enabled = False  # 默认关闭日志功能
        self.log_file_path = ""  # 日志文件路径
        self.log_level = "INFO"  # 日志级别
        self.log_max_size = 10  # 日志文件最大大小（MB）
        self.log_backup_count = 5  # 保留的日志备份数量
        self.logger = None  # 日志记录器

        # 文件后缀白名单配置变量
        self.allowed_file_extensions = [
            '.txt', '.docx', '.doc', '.pdf', '.wps',
            '.py', '.java', '.cpp', '.html', '.js',
            '.md', '.rtf', '.odt', '.xlsx', '.xls',
            '.pptx', '.ppt', '.css', '.json', '.xml',
            '.php', '.c', '.h', '.cs', '.go', '.rs'
        ]  # 允许打开的文件后缀白名单

        # 文件扫描配置变量
        self.scan_subfolders = False  # 是否递归扫描子文件夹，默认为False

    def setup_logging(self):
        """初始化日志系统"""
        if not self.logging_enabled:
            self.logger = None
            return

        try:
            # 设置日志文件路径
            if not self.log_file_path:
                if getattr(sys, 'frozen', False):
                    log_dir = os.path.dirname(sys.executable)
                else:
                    log_dir = os.path.dirname(os.path.abspath(__file__))
                self.log_file_path = os.path.join(log_dir, "activity_tracker.log")

            # 确保日志目录存在
            log_dir = os.path.dirname(self.log_file_path)
            os.makedirs(log_dir, exist_ok=True)

            # 创建日志记录器
            self.logger = logging.getLogger('ActivityTracker')
            self.logger.setLevel(getattr(logging, self.log_level.upper(), logging.INFO))

            # 清除现有的处理器
            self.logger.handlers.clear()

            # 创建文件处理器（带旋转功能）
            from logging.handlers import RotatingFileHandler
            file_handler = RotatingFileHandler(
                self.log_file_path,
                maxBytes=self.log_max_size * 1024 * 1024,  # 转换为字节
                backupCount=self.log_backup_count,
                encoding='utf-8'
            )

            # 设置日志格式
            formatter = logging.Formatter(
                '%(asctime)s - %(levelname)s - %(message)s',
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            file_handler.setFormatter(formatter)

            # 添加处理器到记录器
            self.logger.addHandler(file_handler)

            # 记录日志系统启动
            self.log_info("日志系统已启动", extra_info=f"日志文件: {self.log_file_path}")

        except Exception as e:
            print(f"日志系统初始化失败: {e}")
            self.logger = None

    def log_info(self, message, extra_info=None):
        """记录信息级别日志"""
        if self.logger:
            full_message = f"{message}"
            if extra_info:
                full_message += f" | {extra_info}"
            self.logger.info(full_message)

    def log_warning(self, message, extra_info=None):
        """记录警告级别日志"""
        if self.logger:
            full_message = f"{message}"
            if extra_info:
                full_message += f" | {extra_info}"
            self.logger.warning(full_message)

    def log_error(self, message, extra_info=None):
        """记录错误级别日志"""
        if self.logger:
            full_message = f"{message}"
            if extra_info:
                full_message += f" | {extra_info}"
            self.logger.error(full_message)

    def log_debug(self, message, extra_info=None):
        """记录调试级别日志"""
        if self.logger:
            full_message = f"{message}"
            if extra_info:
                full_message += f" | {extra_info}"
            self.logger.debug(full_message)

    def create_widgets(self):
        # 状态显示区域
        status_frame = Frame(self.root, padx=5, pady=5)
        status_frame.pack(fill="x")

        self.status_label = Label(status_frame, text="就绪", fg="blue")
        self.status_label.pack(anchor="w")

        # 下次执行时间显示
        time_info_frame = Frame(self.root, padx=5, pady=5)
        time_info_frame.pack(fill="x")

        self.save_time_label = Label(time_info_frame, text="实际保存时间: 未设置", fg="gray")
        self.save_time_label.pack(anchor="w")

        # 项目文件夹设置区域
        self.folder_frame = Frame(self.root, padx=5, pady=10)
        self.folder_frame.pack(fill="x")

        Label(self.folder_frame, text="项目文件夹设置:", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky="w", pady=5, columnspan=3)

        # 添加文件夹按钮
        Button(self.folder_frame, text="+ 添加文件夹", command=self.add_folder_row).grid(row=1, column=0, sticky="w", pady=3, padx=2)

        # 如果已有配置的文件夹，创建对应的输入框
        self.refresh_folder_widgets()

        # 工作时间设置区域
        time_frame = Frame(self.root, padx=5, pady=10)
        time_frame.pack(fill="x")

        Label(time_frame, text="工作开始时间 (时:分):").grid(row=0, column=0, sticky="w", pady=3, padx=2)
        Entry(time_frame, textvariable=self.work_start_hour, width=6).grid(row=0, column=1, sticky="w", pady=3, padx=2)
        Label(time_frame, text=":").grid(row=0, column=2, sticky="w", pady=3)
        Entry(time_frame, textvariable=self.work_start_minute, width=6).grid(row=0, column=3, sticky="w", pady=3, padx=2)

        Label(time_frame, text="工作结束时间 (时:分):").grid(row=1, column=0, sticky="w", pady=3, padx=2)
        Entry(time_frame, textvariable=self.work_end_hour, width=6).grid(row=1, column=1, sticky="w", pady=3, padx=2)
        Label(time_frame, text=":").grid(row=1, column=2, sticky="w", pady=3)
        Entry(time_frame, textvariable=self.work_end_minute, width=6).grid(row=1, column=3, sticky="w", pady=3, padx=2)

        Label(time_frame, text="保存延迟范围 (分钟):").grid(row=2, column=0, sticky="w", padx=2)
        Entry(time_frame, textvariable=self.save_delay_min, width=6).grid(row=2, column=1, padx=2)
        Label(time_frame, text="-").grid(row=2, column=2)
        Entry(time_frame, textvariable=self.save_delay_max, width=6).grid(row=2, column=3, padx=2)

        Label(time_frame, text="文件打开间隔 (分钟):").grid(row=3, column=0, sticky="w", padx=2)
        Entry(time_frame, textvariable=self.file_interval_min, width=6).grid(row=3, column=1, padx=2)
        Label(time_frame, text="-").grid(row=3, column=2)
        Entry(time_frame, textvariable=self.file_interval_max, width=6).grid(row=3, column=3, padx=2)

        # 为时间设置变量添加变化监听器
        self.work_start_hour.trace_add('write', self.on_time_setting_changed)
        self.work_start_minute.trace_add('write', self.on_time_setting_changed)
        self.work_end_hour.trace_add('write', self.on_time_setting_changed)
        self.work_end_minute.trace_add('write', self.on_time_setting_changed)

        # 控制按钮
        button_frame = Frame(self.root, padx=5, pady=10)
        button_frame.pack(fill="x")

        self.toggle_btn = Button(button_frame, text="启动自动任务", command=self.toggle_task,bg="#00D2AA", fg="white", font=('SimHei' ,16, "bold"),
                relief="raised", bd=3, padx=20, pady=10, width=16,
                activebackground="#00C5A3", cursor="hand2")
        self.toggle_btn.pack(side="left", padx=5)

        self.run_once_btn = Button(button_frame, text="立即执行一次", command=self.run_once,
            bg="#87CEEB", fg="white",
            relief="raised", bd=2, padx=12, pady=6, cursor="hand2",
            activebackground="#7FC7E8")
        self.run_once_btn.pack(side="left", padx=5)

        self.close_btn = Button(button_frame, text="关闭所有软件", command=self.close_all_programs,
            bg="#FFA07A", fg="white",
            relief="raised", bd=2, padx=12, pady=6, cursor="hand2",
            activebackground="#FF9370")
        self.close_btn.pack(side="left", padx=5)

    def on_time_setting_changed(self, *args):
        """时间设置变化的回调方法（带防抖机制）"""
        # 如果程序正在运行，不处理时间设置变化
        if self.running:
            return

        # 取消之前的计时器
        if self.time_change_timer:
            self.root.after_cancel(self.time_change_timer)

        # 设置新的计时器，500毫秒后执行检查（防抖）
        self.time_change_timer = self.root.after(500, self.delayed_time_setting_check)

    def delayed_time_setting_check(self):
        """延迟执行的时间设置检查"""
        try:
            if not self.running and self.check_time_settings_changed():
                # 只记录日志，不更新界面显示
                self.log_info("时间设置变化", "检测到时间设置变化，已清除工作时间缓存")
        except Exception as e:
            # 避免在UI回调中出现异常
            pass
        finally:
            self.time_change_timer = None

    def add_folder_row(self):
        """添加一个新的文件夹输入行"""
        row_index = len(self.folder_vars) + 2  # 从第2行开始（标题和添加按钮占用前2行）

        # 创建StringVar用于存储路径
        folder_var = StringVar()
        self.folder_vars.append(folder_var)

        # 创建界面元素
        label = Label(self.folder_frame, text=f"文件夹 {len(self.folder_vars)}:")
        entry = Entry(self.folder_frame, textvariable=folder_var, width=50)
        browse_btn = Button(self.folder_frame, text="浏览",
                           command=lambda var=folder_var: self.browse_folder(var))
        remove_btn = Button(self.folder_frame, text="删除",
                           command=lambda idx=len(self.folder_vars)-1: self.remove_folder_row(idx))

        # 布局
        label.grid(row=row_index, column=0, sticky="w", pady=3, padx=2)
        entry.grid(row=row_index, column=1, pady=3, padx=2)
        browse_btn.grid(row=row_index, column=2, padx=2)
        remove_btn.grid(row=row_index, column=3, padx=2)

        # 存储widget引用以便后续删除
        self.folder_widgets.append({
            'label': label,
            'entry': entry,
            'browse_btn': browse_btn,
            'remove_btn': remove_btn
        })

    def remove_folder_row(self, index):
        """删除指定索引的文件夹行"""
        if 0 <= index < len(self.folder_vars):
            # 销毁界面元素
            for widget in self.folder_widgets[index].values():
                widget.destroy()

            # 从列表中移除
            del self.folder_vars[index]
            del self.folder_widgets[index]

            # 重新排列剩余的界面元素
            self.refresh_folder_widgets()

    def refresh_folder_widgets(self):
        """刷新文件夹输入框布局"""
        # 确保folder_widgets已初始化
        if not hasattr(self, 'folder_widgets'):
            self.folder_widgets = []

        # 清除现有的widget
        for widget_dict in self.folder_widgets:
            for widget in widget_dict.values():
                widget.destroy()

        self.folder_widgets.clear()

        # 重新创建所有文件夹输入框
        for i, folder_var in enumerate(self.folder_vars):
            row_index = i + 2

            label = Label(self.folder_frame, text=f"文件夹 {i+1}:")
            entry = Entry(self.folder_frame, textvariable=folder_var, width=50)
            browse_btn = Button(self.folder_frame, text="浏览",
                               command=lambda var=folder_var: self.browse_folder(var))
            remove_btn = Button(self.folder_frame, text="删除",
                               command=lambda idx=i: self.remove_folder_row(idx))

            label.grid(row=row_index, column=0, sticky="w", pady=3, padx=2)
            entry.grid(row=row_index, column=1, pady=3, padx=2)
            browse_btn.grid(row=row_index, column=2, padx=2)
            remove_btn.grid(row=row_index, column=3, padx=2)

            self.folder_widgets.append({
                'label': label,
                'entry': entry,
                'browse_btn': browse_btn,
                'remove_btn': remove_btn
            })

        # 如果没有任何文件夹，添加第一个
        if not self.folder_vars:
            self.add_folder_row()

    def browse_folder(self, var):
        """浏览并选择文件夹"""
        folder_path = filedialog.askdirectory(title="选择项目文件夹")
        if folder_path:
            var.set(folder_path)

    def ensure_config_exists(self):
        """如果配置文件不存在，自动创建一个默认配置文件"""
        if not os.path.exists(self.config_path):
            try:
                default_config = {
                    # 基本配置
                    "project_folders": [],  # 新的多文件夹配置
                    "work_start_hour": 9,
                    "work_start_minute": 0,
                    "work_end_hour": 18,
                    "work_end_minute": 0,
                    "work_time_random_range": 20,  # 工作时间随机区间（分钟）

                    # 午休时间配置
                    "lunch_break": {
                        "enabled": True,           # 是否启用午休功能
                        "start_hour": 11,          # 午休开始时间（小时）
                        "start_minute": 30,         # 午休开始时间（分钟）
                        "end_hour": 13,            # 午休结束时间（小时）
                        "end_minute": 30,          # 午休结束时间（分钟）
                        "random_range": 5         # 午休时间随机区间（分钟）
                    },

                    "save_delay_min": 20,
                    "save_delay_max": 50,
                    "file_interval_min": 30.0,
                    "file_interval_max": 60.0,

                    # 工作日历配置
                    "work_calendar": {
                        "skip_weekends": True,
                        "work_dates": [],  # 调休工作日期 (格式: "MM-DD"每年生效 或 "YY-MM-DD"指定年份)
                        "holiday_dates": []  # 节假日日期 (格式: "MM-DD"每年生效 或 "YY-MM-DD"指定年份)
                    },

                    # 智能关闭功能配置
                    "auto_close_on_work_end": True,
                    "close_strategy": {
                        "use_alt_f4": True,
                        "use_ctrl_q": True,
                        "use_ctrl_w": True,
                        "close_timeout": 3.0
                    },

                    # 文件跟踪配置
                    "file_tracking": {
                        "enabled": True,
                        "track_program_mapping": True,
                        "clear_tracking_on_stop": True
                    },

                    # 程序检测配置
                    "program_detection": {
                        "window_check_interval": 1.0,
                        "activation_delay": 0.5,
                        "close_verification_delay": 1.5
                    },

                    # 用户界面配置
                    "ui_settings": {
                        "show_close_progress": True,
                        "show_detected_programs": True,
                        "status_update_interval": 1.0
                    },

                    # 日志功能配置
                    "logging": {
                        "enabled": False,  # 默认关闭日志功能
                        "log_file_path": "",  # 日志文件路径（空则使用默认路径）
                        "log_level": "INFO",  # 日志级别：DEBUG, INFO, WARNING, ERROR
                        "log_max_size": 10,  # 日志文件最大大小（MB）
                        "log_backup_count": 5  # 保留的日志备份数量
                    },

                    # 文件过滤配置
                    "file_filtering": {
                        "allowed_extensions": [
                            ".txt", ".docx", ".doc", ".pdf", ".wps",
                            ".py", ".java", ".cpp", ".html", ".js",
                            ".md", ".rtf", ".odt", ".xlsx", ".xls",
                            ".pptx", ".ppt", ".css", ".json", ".xml",
                            ".php", ".c", ".h", ".cs", ".go", ".vue", 'xmind'
                        ],  # 允许打开的文件后缀白名单
                        "scan_subfolders": False  # 是否递归扫描子文件夹，默认关闭
                    }
                }

                os.makedirs(os.path.dirname(self.config_path), exist_ok=True)

                with open(self.config_path, "w", encoding="utf-8") as f:
                    json.dump(default_config, f, ensure_ascii=False, indent=2)

                self.log_info("默认配置创建", f"已创建默认配置文件: {self.config_path}")
            except Exception as e:
                messagebox.showwarning("警告", f"创建默认配置文件失败: {str(e)}\n程序仍可运行，但配置不会被保存")

    def get_random_file_from_all_folders(self, recursive=True):
        """从所有配置的文件夹中随机选择一个文件

        Args:
            recursive (bool): 是否递归扫描子文件夹，默认为True
        """
        all_files = []

        # 收集所有文件夹中的文件
        for folder_var in self.folder_vars:
            folder_path = folder_var.get().strip()
            if not folder_path or not os.path.exists(folder_path):
                continue

            try:
                if recursive:
                    # 递归扫描所有子文件夹
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            self._process_file(file, root, all_files)
                else:
                    # 只扫描当前文件夹，不包含子文件夹
                    try:
                        files = os.listdir(folder_path)
                        for file in files:
                            file_path = os.path.join(folder_path, file)
                            # 确保是文件，不是文件夹
                            if os.path.isfile(file_path):
                                self._process_file(file, folder_path, all_files)
                    except (OSError, PermissionError):
                        continue
            except Exception as e:
                self.update_status(f"获取文件夹 {folder_path} 内文件失败: {str(e)}", "orange")
                continue

        if all_files:
            # 确保真正的随机选择
            random.shuffle(all_files)  # 先打乱文件列表
            selected_file = random.choice(all_files)

            # 记录选择过程（用于调试）
            self.log_info("文件随机选择", f"候选文件数量: {len(all_files)} | 已选择: {os.path.basename(selected_file)}")

            return selected_file
        else:
            self.log_warning("文件选择", "所有配置文件夹中未找到符合条件的文件")
            return None

    def _process_file(self, file, root, all_files):
        """处理单个文件的过滤逻辑"""
        # 构建完整文件路径
        full_file_path = os.path.join(root, file)

        # 多重验证确保这是一个有效文件
        try:
            # 检查1: 确保路径存在
            if not os.path.exists(full_file_path):
                return

            # 检查2: 确保这是一个文件而不是文件夹
            if not os.path.isfile(full_file_path):
                return

            # 检查3: 确保不是目录（双重保险）
            if os.path.isdir(full_file_path):
                return

            # 检查4: 尝试获取文件大小（这会进一步确认文件有效性）
            file_size = os.path.getsize(full_file_path)
            if file_size < 0:  # 理论上文件大小不应该为负
                return

        except (OSError, IOError, PermissionError):
            # 如果任何文件系统操作失败，跳过这个文件
            return

        # 跳过临时文件和隐藏文件
        if (file.startswith('~$') or     # Office/WPS临时文件
            file.startswith('.~') or      # 其他临时文件
            file.startswith('~') or       # 临时文件
            file.endswith('.tmp') or      # 临时文件
            file.endswith('.temp') or     # 临时文件
            file.startswith('.')):        # 隐藏文件
            return

        # 使用配置的文件后缀白名单过滤
        file_extension = os.path.splitext(file)[1].lower()
        if file_extension in [ext.lower() for ext in self.allowed_file_extensions]:
            all_files.append(full_file_path)

    def check_available_files(self):
        """检查所有配置文件夹中是否有可用的文件"""
        # 根据递归扫描配置决定扫描方式
        files_found = self.get_random_file_from_all_folders(recursive=self.scan_subfolders)

        return {
            'has_files_current_level': files_found is not None if not self.scan_subfolders else False,
            'has_files_recursive': files_found is not None if self.scan_subfolders else False,
            'total_files_current': self._count_files(recursive=False) if not self.scan_subfolders else 0,
            'total_files_recursive': self._count_files(recursive=True) if self.scan_subfolders else 0
        }

    def _count_files(self, recursive=True):
        """统计可用文件数量"""
        all_files = []

        for folder_var in self.folder_vars:
            folder_path = folder_var.get().strip()
            if not folder_path or not os.path.exists(folder_path):
                continue

            try:
                if recursive:
                    # 递归扫描所有子文件夹
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            self._process_file(file, root, all_files)
                else:
                    # 只扫描当前文件夹
                    try:
                        files = os.listdir(folder_path)
                        for file in files:
                            file_path = os.path.join(folder_path, file)
                            if os.path.isfile(file_path):
                                self._process_file(file, folder_path, all_files)
                    except (OSError, PermissionError):
                        continue
            except Exception:
                continue

        return len(all_files)

    def open_random_file(self):
        """打开一个随机文件"""
        # 直接使用配置的扫描方式
        random_file = self.get_random_file_from_all_folders(recursive=self.scan_subfolders)

        if random_file:
            try:
                # 再次验证文件是否存在且是有效文件
                if not os.path.exists(random_file):
                    self.log_warning("文件打开失败", f"文件不存在: {random_file}")
                    self.update_status("选中的文件不存在，请重试", "orange")
                    return False

                if not os.path.isfile(random_file):
                    self.log_warning("文件打开失败", f"路径不是文件: {random_file}")
                    self.update_status("选中的路径不是有效文件，请重试", "orange")
                    return False

                # 检查文件大小，避免打开过大的文件
                try:
                    file_size = os.path.getsize(random_file)
                    # 限制文件大小不超过100MB
                    if file_size > 100 * 1024 * 1024:
                        self.log_warning("文件打开跳过", f"文件过大: {random_file} ({file_size / 1024 / 1024:.1f}MB)")
                        self.update_status(f"文件过大，跳过打开: {os.path.basename(random_file)}", "orange")
                        return False
                except OSError:
                    self.log_warning("文件大小检查失败", f"无法获取文件大小: {random_file}")

                file_name = os.path.basename(random_file)
                folder_name = os.path.basename(os.path.dirname(random_file))
                self.update_status(f"正在打开文件: {file_name} (来自: {folder_name})")

                # 记录文件打开日志
                self.log_info("打开文件", f"文件: {file_name} | 路径: {random_file} | 来源文件夹: {folder_name}")

                # 使用系统默认程序打开文件
                if sys.platform.startswith('win32'):
                    os.startfile(random_file)
                elif sys.platform.startswith('darwin'):
                    subprocess.Popen(['open', random_file])
                else:
                    subprocess.Popen(['xdg-open', random_file])

                time.sleep(2)  # 等待文件打开

                # 设置当前打开的文件
                self.current_opened_file = random_file
                self.current_opened_time = datetime.datetime.now()

                # 根据配置决定是否跟踪打开的文件
                if self.file_tracking_enabled:
                    self.opened_files.append(random_file)

                    # 根据配置决定是否根据文件扩展名判断程序类型
                    if self.track_program_mapping:
                        file_ext = os.path.splitext(random_file)[1].lower()
                        self.track_program_by_file_extension(file_ext)

                return True
            except Exception as e:
                self.update_status(f"打开文件失败: {str(e)}", "orange")
                return False
        else:
            self.update_status("所有配置的文件夹中没有找到合适的文件", "orange")
            return False

    def toggle_task(self):
        """切换任务状态：启动或停止"""
        if self.running:
            self.stop_task()
        else:
            self.start_task()

    def start_task(self):
        # 首先验证所有输入
        if not self.validate_all_inputs():
            return

        # 检查是否至少配置了一个文件夹
        valid_folders = [var.get().strip() for var in self.folder_vars if var.get().strip()]
        if not valid_folders:
            messagebox.showwarning("警告", "请至少添加一个项目文件夹")
            return

        # 验证文件夹是否存在
        for folder_path in valid_folders:
            if not os.path.exists(folder_path):
                self.log_error("文件夹验证失败", f"文件夹不存在: {folder_path}")
                messagebox.showwarning("警告", f"文件夹不存在: {folder_path}")
                return

        # 在启动任务前检查是否有可用文件
        try:
            file_check = self.check_available_files()
        except Exception as e:
            self.log_error("文件检查失败", f"错误: {str(e)}")
            messagebox.showerror("文件检查失败", f"无法扫描文件夹: {str(e)}")
            return

        if not file_check['has_files_current_level'] and not file_check['has_files_recursive']:
            messagebox.showwarning(
                "无可用文件",
                "扫描所有配置的文件夹后未找到任何可打开的文件！\n\n"
                f"支持的文件格式：{', '.join(self.allowed_file_extensions)}\n\n"
                "请检查：\n"
                "1. 文件夹路径是否正确\n"
                "2. 文件夹中是否包含支持格式的文件\n"
                "3. 文件夹访问权限是否正常"
            )
            return

        # 系统兼容性检查
        try:
            self.check_system_compatibility()
        except Exception as e:
            self.log_error("系统兼容性检查失败", f"错误: {str(e)}")
            if messagebox.askyesno("兼容性警告",
                f"检测到系统兼容性问题：{str(e)}\n\n是否仍要继续启动？\n注意：程序可能无法正常工作"):
                pass
            else:
                return

        self.save_config()

        # 启动时强制清除工作时间缓存，确保重新计算
        old_cache_count = len(self.daily_work_times)
        self.daily_work_times.clear()

        # 设置强制重新计算标记，确保启动时重新生成今天的工作时间
        self.force_recalculate_today = True

        self.log_info("任务启动", f"已清除 {old_cache_count} 个缓存的工作时间，将重新计算今日工作时间")

        # 检查时间设置是否发生了变化，如果变化则清除缓存
        if self.check_time_settings_changed():
            self.log_info("任务启动", "检测到时间设置变化，将使用新的时间设置")

        self.running = True
        self.cancel_event.clear()
        self.status_label.config(text="自动任务已启动，将按计划执行", fg="green")

        # 记录启动日志
        self.log_info("自动任务已启动", f"工作时间: {self.work_start_hour.get():02d}:{self.work_start_minute.get():02d} - {self.work_end_hour.get():02d}:{self.work_end_minute.get():02d}")

        # 更新按钮状态 - 切换为停止状态
        self.toggle_btn.config(text="停止自动任务", bg="#FF6B6B", activebackground="#FF5252")

        # 显示可用文件统计信息
        current_count = file_check['total_files_current']
        recursive_count = file_check['total_files_recursive']

        if current_count > 0:
            self.log_info("文件扫描结果",
                f"当前目录级别：{current_count} 个文件，包含子文件夹：{recursive_count} 个文件")
        else:
            self.log_info("文件扫描结果",
                f"当前目录级别无可用文件，仅在子文件夹中找到：{recursive_count} 个文件")

        # 更新时间显示（移除"尚未启动"标识）
        self.update_save_time()

        # 启动任务线程
        try:
            self.task_thread = threading.Thread(target=self.task_loop, daemon=True)
            self.task_thread.start()
            self.log_info("任务线程启动", "任务线程已成功启动")
        except Exception as e:
            self.log_error("任务线程启动失败", f"错误: {str(e)}")
            self.running = False
            self.cancel_event.set()
            self.toggle_btn.config(text="启动自动任务", bg="#00D2AA", activebackground="#00C5A3")
            messagebox.showerror("启动失败", f"无法启动任务线程: {str(e)}")

    def check_system_compatibility(self):
        """检查系统兼容性，特别是exe环境"""
        issues = []

        # 检查pyautogui可用性
        if not PYAUTOGUI_AVAILABLE:
            issues.append("pyautogui模块不可用，无法执行自动操作")

        # 检查win32com可用性
        if not WIN32COM_AVAILABLE:
            issues.append("win32com模块不可用，部分功能可能受限")

        # 检查是否在exe环境中
        if getattr(sys, 'frozen', False):
            self.log_info("运行环境检测", "程序在exe环境中运行")

            # 检查exe路径权限
            try:
                exe_dir = os.path.dirname(sys.executable)
                test_file = os.path.join(exe_dir, "test_write.tmp")
                with open(test_file, "w") as f:
                    f.write("test")
                os.remove(test_file)
            except Exception as e:
                issues.append(f"exe目录无写入权限: {str(e)}")

        # 检查配置文件权限
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r") as f:
                    f.read()
            else:
                with open(self.config_path, "w") as f:
                    f.write("{}")
                os.remove(self.config_path)
        except Exception as e:
            issues.append(f"配置文件访问权限问题: {str(e)}")

        if issues:
            raise Exception("; ".join(issues))

        self.log_info("系统兼容性检查", "所有检查通过")

    def start_heartbeat_monitor(self):
        """启动心跳监控"""
        try:
            self.heartbeat_count = 0
            self.last_heartbeat_time = datetime.datetime.now()
            self.heartbeat_monitor()
            self.log_info("心跳监控启动", "exe环境心跳监控已启动")
        except Exception as e:
            self.log_error("心跳监控启动失败", f"错误: {str(e)}")

    def heartbeat_monitor(self):
        """心跳监控方法"""
        try:
            current_time = datetime.datetime.now()
            self.heartbeat_count += 1

            # 只在出现问题时记录日志，正常运行时不记录
            if self.heartbeat_count % 60 == 0:  # 每60秒检查一次
                # 检查程序是否卡死（仅在运行状态下检查，超过10分钟没有状态更新）
                if self.running and hasattr(self, 'last_status_update_time'):
                    time_since_last_update = (current_time - self.last_status_update_time).total_seconds()
                    if time_since_last_update > 600:  # 10分钟
                        self.log_warning("程序状态检查", f"程序可能卡死，距离上次状态更新已过 {time_since_last_update/60:.1f} 分钟")

            self.last_heartbeat_time = current_time

            # 安排下一次心跳检查（1秒后）
            if hasattr(self, 'root') and self.root:
                self.root.after(1000, self.heartbeat_monitor)

        except Exception as e:
            # 心跳监控本身出错，记录错误但不影响主程序
            self.log_error("心跳监控错误", f"错误: {str(e)}")
            # 如果出错，尝试重新启动心跳监控
            if hasattr(self, 'root') and self.root:
                self.root.after(5000, self.heartbeat_monitor)  # 5秒后重试

    def validate_all_inputs(self):
        """验证所有用户输入，包括类型转换错误处理"""
        try:
            self.log_info("输入验证", "开始验证所有输入参数")

            # 验证工作时间设置
            try:
                start_hour = self.work_start_hour.get()
                start_minute = self.work_start_minute.get()
                end_hour = self.work_end_hour.get()
                end_minute = self.work_end_minute.get()

                if not (0 <= start_hour <= 23 and 0 <= start_minute <= 59 and
                        0 <= end_hour <= 23 and 0 <= end_minute <= 59):
                    raise ValueError("工作时间格式不正确")

                # 检查工作时间范围是否合理
                start_minutes = start_hour * 60 + start_minute
                end_minutes = end_hour * 60 + end_minute
                if start_minutes >= end_minutes:
                    raise ValueError("工作结束时间必须晚于开始时间")

            except Exception as e:
                error_msg = f"工作时间设置错误: {str(e)}"
                self.log_error("输入验证失败", error_msg)
                messagebox.showerror("输入错误", f"{error_msg}\n请检查时间格式（0-23小时，0-59分钟）")
                return False

            # 验证保存延迟设置
            try:
                save_min = self.save_delay_min.get()
                save_max = self.save_delay_max.get()

                if save_min < 0 or save_max < 0:
                    raise ValueError("保存延迟时间不能为负数")

                if save_min > save_max:
                    raise ValueError("最小保存延迟不能大于最大延迟")

            except TclError:
                error_msg = "保存延迟输入格式错误，请输入有效的数字"
                self.log_error("输入验证失败", f"保存延迟格式错误")
                messagebox.showerror("输入错误", error_msg)
                return False
            except Exception as e:
                error_msg = f"保存延迟设置错误: {str(e)}"
                self.log_error("输入验证失败", error_msg)
                messagebox.showerror("输入错误", f"{error_msg}\n请输入有效的数字")
                return False

            # 验证文件打开间隔设置
            try:
                interval_min = self.file_interval_min.get()
                interval_max = self.file_interval_max.get()

                if interval_min < 0 or interval_max < 0:
                    raise ValueError("文件打开间隔不能为负数")

                if interval_min > interval_max:
                    raise ValueError("最小文件打开间隔不能大于最大间隔")

                if interval_min < 0.1:
                    raise ValueError("文件打开间隔不能小于0.1分钟（6秒）")

            except TclError:
                error_msg = "文件打开间隔输入格式错误，请输入有效的数字"
                self.log_error("输入验证失败", f"文件间隔格式错误")
                messagebox.showerror("输入错误", error_msg + "\n\n输入提示：\n• 支持小数，如：0.5、1.2、2.5\n• 请使用英文句号(.)\n• 最小值为0.1分钟")
                return False
            except Exception as e:
                error_msg = f"文件打开间隔设置错误: {str(e)}"
                self.log_error("输入验证失败", error_msg)
                messagebox.showerror("输入错误", f"{error_msg}\n\n输入提示：\n• 支持小数，如：0.5、1.2、2.5\n• 请使用英文句号(.)\n• 最小值为0.1分钟")
                return False

            self.log_info("输入验证", "所有输入参数验证通过")
            return True

        except Exception as e:
            error_msg = f"输入验证过程中发生异常: {str(e)}"
            self.log_error("输入验证异常", error_msg)
            messagebox.showerror("验证异常", error_msg)
            return False

    def stop_task(self):
        if self.running:
            self.running = False
            self.cancel_event.set()

            # 记录停止日志
            self.log_info("自动任务已停止", f"运行时长统计已记录")

            # 使用root.after确保UI更新在主线程中执行
            self.root.after(0, lambda: self.status_label.config(text="任务已停止", fg="blue"))
            self.root.after(0, lambda: self.toggle_btn.config(text="启动自动任务", bg="#00D2AA", activebackground="#00C5A3"))

            # 清理状态
            self.actual_save_time = None
            self.current_opened_file = None
            self.current_opened_time = None
            self.next_file_open_time = None  # 清除下次文件打开时间
            self.root.after(0, self.update_save_time)

            # 根据配置决定是否清空文件和程序跟踪列表
            # 注意：即使配置为清空，我们也会延迟清空，让用户有机会关闭软件
            if self.clear_tracking_on_stop:
                # 记录当前跟踪信息，便于用户查看
                if self.opened_files or self.opened_programs:
                    self.log_info("跟踪信息保留",
                        f"保留跟踪信息以便关闭软件 - 文件: {len(self.opened_files)}个, 程序: {len(self.opened_programs)}个")
                    self.update_status("任务已停止，跟踪信息已保留用于关闭软件", "blue")
                else:
                    self.update_status("任务已停止", "blue")
            else:
                self.update_status("任务已停止，跟踪信息已保留", "blue")

    def is_work_day(self, date):
        """判断指定日期是否为工作日

        支持两种日期格式：
        - "MM-DD": 每年生效的月日配置（如 "01-01" 表示每年1月1日）
        - "YY-MM-DD": 指定年份的年月日配置（如 "25-07-21" 表示2025年7月21日）
        """
        try:
            # 获取星期几 (0=周一, 6=周日)
            weekday = date.weekday()

            # 格式化日期为两种格式
            date_str_mmdd = date.strftime("%m-%d")  # MM-DD 格式
            date_str_yymmdd = date.strftime("%y-%m-%d")  # YY-MM-DD 格式

            # 检查是否为节假日（支持两种格式）
            for holiday_date in self.holiday_dates:
                if holiday_date == date_str_mmdd or holiday_date == date_str_yymmdd:
                    return False

            # 检查是否为调休工作日（支持两种格式）
            for work_date in self.work_dates:
                if work_date == date_str_mmdd or work_date == date_str_yymmdd:
                    return True

            # 如果配置了跳过周末，且当前是周末（周六=5，周日=6）
            if self.skip_weekends and weekday >= 5:
                return False

            # 其他情况都是工作日
            return True

        except Exception as e:
            self.log_error("工作日判断错误", f"错误: {e}")
            # 出错时默认为工作日，避免程序停止
            return True

    def get_random_lunch_times(self, base_date):
        """获取指定日期的随机午休时间（带缓存机制，确保每天只计算一次）"""
        if not self.lunch_break_enabled:
            return None, None

        # 获取日期字符串作为缓存键
        date_key = f"{base_date.strftime('%Y-%m-%d')}_lunch"

        # 如果已经计算过该日期的午休时间，直接返回缓存的结果
        if date_key in self.daily_work_times:
            return self.daily_work_times[date_key]

        try:
            # 基础午休时间
            base_start_hour = self.lunch_start_hour
            base_start_minute = self.lunch_start_minute
            base_end_hour = self.lunch_end_hour
            base_end_minute = self.lunch_end_minute

            # 获取随机区间（分钟）
            random_range = getattr(self, 'lunch_time_random_range', 10)

            # 计算基础时间（以分钟为单位）
            base_start_minutes = base_start_hour * 60 + base_start_minute
            base_end_minutes = base_end_hour * 60 + base_end_minute

            # 生成随机偏移量（-random_range 到 +random_range 分钟）
            start_offset = random.randint(-random_range, random_range)
            end_offset = random.randint(-random_range, random_range)

            # 计算实际午休时间
            actual_start_minutes = base_start_minutes + start_offset
            actual_end_minutes = base_end_minutes + end_offset

            # 确保时间在合理范围内（0-1439分钟，即0:00-23:59）
            actual_start_minutes = max(0, min(1439, actual_start_minutes))
            actual_end_minutes = max(0, min(1439, actual_end_minutes))

            # 确保结束时间晚于开始时间
            if actual_end_minutes <= actual_start_minutes:
                actual_end_minutes = actual_start_minutes + 30  # 至少午休30分钟
                if actual_end_minutes > 1439:  # 如果超过23:59
                    actual_start_minutes = 1439 - 30  # 调整开始时间
                    actual_end_minutes = 1439

            # 转换回小时和分钟
            start_hour = actual_start_minutes // 60
            start_minute = actual_start_minutes % 60
            end_hour = actual_end_minutes // 60
            end_minute = actual_end_minutes % 60

            # 创建datetime对象
            lunch_start_time = base_date.replace(
                hour=start_hour,
                minute=start_minute,
                second=0,
                microsecond=0
            )
            lunch_end_time = base_date.replace(
                hour=end_hour,
                minute=end_minute,
                second=0,
                microsecond=0
            )

            # 缓存结果
            self.daily_work_times[date_key] = (lunch_start_time, lunch_end_time)

            # 记录日志
            self.log_info("午休时间计算", f"日期: {base_date.strftime('%Y-%m-%d')} | 午休开始: {lunch_start_time.strftime('%H:%M')} | 午休结束: {lunch_end_time.strftime('%H:%M')}")

            return lunch_start_time, lunch_end_time

        except Exception as e:
            self.log_error("午休时间计算错误", f"错误: {e}")
            # 出错时返回None，表示不启用午休
            return None, None

    def is_lunch_time(self, current_time):
        """判断当前时间是否在午休时间内"""
        if not self.lunch_break_enabled:
            return False

        try:
            # 获取今天的午休时间
            lunch_start_time, lunch_end_time = self.get_random_lunch_times(current_time)

            if lunch_start_time is None or lunch_end_time is None:
                return False

            # 检查当前时间是否在午休时间范围内
            return lunch_start_time <= current_time < lunch_end_time

        except Exception as e:
            self.log_error("午休时间判断错误", f"错误: {str(e)}")
            return False

    def get_random_work_times(self, base_date):
        """获取指定日期的随机工作开始和结束时间（带缓存机制，确保每天只计算一次）"""
        # 获取日期字符串作为缓存键
        date_key = base_date.strftime("%Y-%m-%d")
        today_key = datetime.datetime.now().strftime("%Y-%m-%d")

        # 检查是否需要强制重新计算今天的时间（启动时）
        force_recalc = getattr(self, 'force_recalculate_today', False)
        if force_recalc and date_key == today_key:
            # 清除今天的缓存，强制重新计算
            if date_key in self.daily_work_times:
                self.log_info("工作时间重新计算", f"启动时强制重新计算今日 ({date_key}) 工作时间")
                del self.daily_work_times[date_key]
            # 重置标记
            self.force_recalculate_today = False

        # 如果已经计算过该日期的工作时间，直接返回缓存的结果
        if date_key in self.daily_work_times:
            return self.daily_work_times[date_key]

        try:
            # 基础工作时间
            base_start_hour = self.work_start_hour.get()
            base_start_minute = self.work_start_minute.get()
            base_end_hour = self.work_end_hour.get()
            base_end_minute = self.work_end_minute.get()

            # 获取随机区间（分钟）
            random_range = getattr(self, 'work_time_random_range', 20)

            # 计算基础时间（以分钟为单位）
            base_start_minutes = base_start_hour * 60 + base_start_minute
            base_end_minutes = base_end_hour * 60 + base_end_minute

            # 生成随机偏移量（-random_range 到 +random_range 分钟）
            start_offset = random.randint(-random_range, random_range)
            end_offset = random.randint(-random_range, random_range)

            # 计算实际工作时间
            actual_start_minutes = base_start_minutes + start_offset
            actual_end_minutes = base_end_minutes + end_offset

            # 确保时间在合理范围内（0-1439分钟，即0:00-23:59）
            actual_start_minutes = max(0, min(1439, actual_start_minutes))
            actual_end_minutes = max(0, min(1439, actual_end_minutes))

            # 确保结束时间晚于开始时间
            if actual_end_minutes <= actual_start_minutes:
                actual_end_minutes = actual_start_minutes + 60  # 至少工作1小时
                if actual_end_minutes > 1439:  # 如果超过23:59
                    actual_start_minutes = 1439 - 60  # 调整开始时间
                    actual_end_minutes = 1439

            # 转换回小时和分钟
            start_hour = actual_start_minutes // 60
            start_minute = actual_start_minutes % 60
            end_hour = actual_end_minutes // 60
            end_minute = actual_end_minutes % 60

            # 创建datetime对象
            work_start_time = base_date.replace(
                hour=start_hour,
                minute=start_minute,
                second=0,
                microsecond=0
            )
            work_end_time = base_date.replace(
                hour=end_hour,
                minute=end_minute,
                second=0,
                microsecond=0
            )

            # 缓存结果
            self.daily_work_times[date_key] = (work_start_time, work_end_time)

            # 记录日志
            self.log_info("工作时间计算", f"日期: {date_key} | 开始: {work_start_time.strftime('%H:%M')} | 结束: {work_end_time.strftime('%H:%M')}")

            return work_start_time, work_end_time

        except Exception as e:
            self.log_error("工作时间计算错误", f"错误: {e}")
            # 出错时返回基础时间
            work_start_time = base_date.replace(
                hour=self.work_start_hour.get(),
                minute=self.work_start_minute.get(),
                second=0,
                microsecond=0
            )
            work_end_time = base_date.replace(
                hour=self.work_end_hour.get(),
                minute=self.work_end_minute.get(),
                second=0,
                microsecond=0
            )

            # 即使出错也要缓存结果
            self.daily_work_times[date_key] = (work_start_time, work_end_time)

            return work_start_time, work_end_time

    def check_time_settings_changed(self):
        """检查时间设置是否发生了变化，如果变化则清除相关缓存"""
        current_settings = {
            'work_start_hour': self.work_start_hour.get(),
            'work_start_minute': self.work_start_minute.get(),
            'work_end_hour': self.work_end_hour.get(),
            'work_end_minute': self.work_end_minute.get(),
            'work_time_random_range': getattr(self, 'work_time_random_range', 20)
        }

        # 如果是第一次检查，保存当前设置
        if not self.cached_time_settings:
            self.cached_time_settings = current_settings.copy()
            return False

        # 检查是否有任何设置发生了变化
        settings_changed = False
        for key, value in current_settings.items():
            if self.cached_time_settings.get(key) != value:
                settings_changed = True
                self.log_info("时间设置变化检测", f"设置项 {key} 从 {self.cached_time_settings.get(key)} 变更为 {value}")
                break

        if settings_changed:
            # 清除工作时间缓存
            old_cache_count = len(self.daily_work_times)
            self.daily_work_times.clear()
            self.log_info("工作时间缓存清除", f"由于时间设置变化，已清除 {old_cache_count} 个缓存的工作时间")

            # 更新缓存的设置
            self.cached_time_settings = current_settings.copy()

            return True

        return False

    def task_loop(self):
        """新的工作模式主循环 - 增强错误处理"""
        try:
            self.log_info("任务循环启动", "开始执行主任务循环")

            while self.running:
                try:
                    # 检查取消事件
                    if self.cancel_event.is_set():
                        self.update_status("任务已取消", "orange")
                        break

                    now = datetime.datetime.now()

                    # 获取今天的随机工作时间
                    try:
                        work_start_time, work_end_time = self.get_random_work_times(now)
                        self.log_info("工作时间获取", f"今日工作时间: {work_start_time.strftime('%H:%M')} - {work_end_time.strftime('%H:%M')}")
                    except Exception as e:
                        self.log_error("工作时间计算失败", f"错误: {str(e)}")
                        # 使用默认时间
                        work_start_time = now.replace(hour=9, minute=0, second=0, microsecond=0)
                        work_end_time = now.replace(hour=18, minute=0, second=0, microsecond=0)
                        self.log_warning("工作时间回退", f"使用默认工作时间: {work_start_time.strftime('%H:%M')} - {work_end_time.strftime('%H:%M')}")

                    # 如果当前时间已经过了今天的结束时间，寻找下一个工作日
                    if now >= work_end_time:
                        # 寻找下一个工作日
                        next_work_day = now.date() + datetime.timedelta(days=1)
                        while not self.is_work_day(next_work_day):
                            next_work_day += datetime.timedelta(days=1)

                        # 设置下一个工作日的随机工作时间
                        next_day_datetime = datetime.datetime.combine(next_work_day, datetime.time(9, 0))
                        try:
                            work_start_time, work_end_time = self.get_random_work_times(next_day_datetime)
                        except Exception as e:
                            self.log_error("下一工作日时间计算失败", f"错误: {str(e)}")
                            work_start_time = next_day_datetime
                            work_end_time = next_day_datetime.replace(hour=18)

                    # 如果当前时间在工作时间范围内，但今天不是工作日
                    elif work_start_time <= now < work_end_time:
                        if not self.is_work_day(now.date()):
                            # 今天不是工作日，寻找下一个工作日
                            next_work_day = now.date() + datetime.timedelta(days=1)
                            while not self.is_work_day(next_work_day):
                                next_work_day += datetime.timedelta(days=1)

                            next_day_datetime = datetime.datetime.combine(next_work_day, datetime.time(9, 0))
                            try:
                                work_start_time, work_end_time = self.get_random_work_times(next_day_datetime)
                            except Exception as e:
                                self.log_error("下一工作日时间计算失败", f"错误: {str(e)}")
                                work_start_time = next_day_datetime
                                work_end_time = next_day_datetime.replace(hour=18)
                        else:
                            # 今天是工作日且在工作时间内
                            self.update_status("当前在工作时间内，开始工作模式...")
                            try:
                                self.work_mode(work_end_time)
                            except Exception as e:
                                self.log_error("工作模式执行失败", f"错误: {str(e)}")
                                self.update_status(f"工作模式执行失败: {str(e)}", "red")
                            continue
                    # 如果当前时间在今天工作开始时间之前
                    else:
                        if not self.is_work_day(now.date()):
                            # 今天不是工作日，寻找下一个工作日
                            next_work_day = now.date() + datetime.timedelta(days=1)
                            while not self.is_work_day(next_work_day):
                                next_work_day += datetime.timedelta(days=1)

                            next_day_datetime = datetime.datetime.combine(next_work_day, datetime.time(9, 0))
                            try:
                                work_start_time, work_end_time = self.get_random_work_times(next_day_datetime)
                            except Exception as e:
                                self.log_error("下一工作日时间计算失败", f"错误: {str(e)}")
                                work_start_time = next_day_datetime
                                work_end_time = next_day_datetime.replace(hour=18)

                    # 等待到工作开始时间
                    wait_seconds = (work_start_time - now).total_seconds()
                    if wait_seconds > 0:
                        # 检查今天是否为工作日（用于显示信息）
                        today_is_work_day = self.is_work_day(now.date())
                        target_is_work_day = self.is_work_day(work_start_time.date())

                        # 检查是否跨天了
                        if work_start_time.date() > now.date():
                            days_diff = (work_start_time.date() - now.date()).days
                            if days_diff == 1:
                                # 动态判断是否为明天（避免硬编码）
                                tomorrow_date = now.date() + datetime.timedelta(days=1)
                                if work_start_time.date() == tomorrow_date:
                                    day_text = "明天"
                                else:
                                    day_text = work_start_time.strftime('%m-%d')

                                # 区分当前状态
                                if today_is_work_day:
                                    self.update_status(f"目前非工作时间，等待到{day_text} {work_start_time.strftime('%H:%M')} 开始工作...")
                                else:
                                    self.update_status(f"今天非工作日，等待到{day_text} {work_start_time.strftime('%H:%M')} 开始工作...")
                            else:
                                self.update_status(f"等待到 {work_start_time.strftime('%m-%d %H:%M')} 开始工作（{days_diff}天后）...")
                        else:
                            self.update_status(f"等待到今天 {work_start_time.strftime('%H:%M')} 开始工作...")

                        self.actual_save_time = work_start_time
                        self.update_save_time()

                        cancelled = self.wait_with_cancel(wait_seconds, work_start_time)
                        if cancelled:
                            break

                    if self.running and not self.cancel_event.is_set():
                        # 再次确认当前时间确实在工作时间内
                        current_time = datetime.datetime.now()

                        # 使用同一天的工作时间进行验证，避免重新计算导致的时间不一致
                        if current_time.date() == work_start_time.date():
                            # 同一天，使用已计算的工作时间
                            if work_start_time <= current_time < work_end_time:
                                self.update_status("工作时间到，开始文档操作...")
                                self.log_info("工作模式启动", f"确认时间: {current_time.strftime('%H:%M:%S')} 在工作时间 {work_start_time.strftime('%H:%M')}-{work_end_time.strftime('%H:%M')} 内")
                            else:
                                self.log_warning("工作时间确认失败", f"当前时间 {current_time.strftime('%H:%M:%S')} 不在工作时间范围内，将重新计算")
                                continue
                        else:
                            # 跨天了，重新计算当前日期的工作时间
                            current_work_start, current_work_end = self.get_random_work_times(current_time)
                            if current_work_start <= current_time < current_work_end:
                                self.update_status("工作时间到，开始文档操作...")
                                self.log_info("工作模式启动", f"确认时间: {current_time.strftime('%H:%M:%S')} 在工作时间 {current_work_start.strftime('%H:%M')}-{current_work_end.strftime('%H:%M')} 内")
                                work_end_time = current_work_end  # 更新工作结束时间
                            else:
                                self.log_warning("工作时间确认失败", f"当前时间 {current_time.strftime('%H:%M:%S')} 不在工作时间范围内")
                                continue

                        # 进入工作模式
                        try:
                            self.work_mode(work_end_time)
                        except Exception as e:
                            self.log_error("工作模式执行失败", f"错误: {str(e)}")
                            self.update_status(f"工作模式执行失败: {str(e)}", "red")
                            # 等待一段时间后重试，避免快速循环
                            if self.wait_with_cancel(60):  # 等待60秒
                                break

                except Exception as e:
                    self.log_error("任务循环迭代失败", f"错误: {str(e)}")
                    self.update_status(f"任务循环出错: {str(e)}", "red")
                    # 等待一段时间后重试，避免快速循环
                    if self.wait_with_cancel(30):  # 等待30秒
                        break

        except Exception as e:
            self.log_error("任务循环严重错误", f"错误: {str(e)}")
            self.update_status(f"任务循环严重错误: {str(e)}", "red")
        finally:
            self.update_status("就绪", "blue")
            self.log_info("任务循环结束", "主任务循环已结束")

    def work_mode(self, work_end_time):
        """工作模式：在工作时间内不断打开和保存文件"""
        start_time = datetime.datetime.now()
        self.log_info("工作模式开始", f"开始时间: {start_time.strftime('%H:%M:%S')}, 预计结束时间: {work_end_time.strftime('%H:%M:%S')}")
        self.update_status("进入工作模式，开始文档操作...")

        # 首次打开文件（工作开始时）
        valid_folders = [var.get().strip() for var in self.folder_vars if var.get().strip()]
        if valid_folders:
            self.update_status("工作开始，打开第一个文档文件...")
            self.open_random_file()

            # 为第一个文件设置保存延迟
            self.schedule_save_for_current_file(work_end_time)

            # 计算并显示下一次文件打开时间
            self.calculate_next_file_open_time(work_end_time)

        last_file_open_time = datetime.datetime.now()

        while self.running:
            # 检查取消事件
            if self.cancel_event.is_set():
                self.update_status("任务已取消", "orange")
                break

            now = datetime.datetime.now()

            # 检查是否已到工作结束时间
            if now >= work_end_time:
                # 二次验证：重新获取今天的工作时间，确保时间一致性
                try:
                    today_work_start, today_work_end = self.get_random_work_times(now)

                    # 如果重新计算的结束时间与原来的不一致，使用最新的时间
                    if today_work_end != work_end_time:
                        self.log_warning("工作时间不一致",
                            f"原结束时间: {work_end_time.strftime('%H:%M:%S')}, "
                            f"重新计算: {today_work_end.strftime('%H:%M:%S')}")
                        work_end_time = today_work_end

                        # 如果当前时间还在新的工作时间内，继续工作
                        if now < work_end_time:
                            self.log_info("工作时间延续", f"使用更新的结束时间 {work_end_time.strftime('%H:%M')}")
                            continue

                    # 确认工作时间确实结束了
                    if now >= work_end_time:
                        self.log_info("工作时间确认结束", f"当前时间 {now.strftime('%H:%M:%S')} >= 结束时间 {work_end_time.strftime('%H:%M:%S')}")
                        self.update_status("工作时间结束，正在处理未保存的文件...")
                    else:
                        continue

                except Exception as e:
                    self.log_error("工作时间验证失败", f"错误: {str(e)}，使用原始结束时间")
                    self.update_status("工作时间结束，正在处理未保存的文件...")

                # 检查是否有当前打开的文件还未保存
                if self.current_opened_file and self.actual_save_time:
                    # 如果有未保存的文件，先保存它
                    if self.actual_save_time > now:
                        self.log_info("工作结束处理", f"检测到未保存的文件: {os.path.basename(self.current_opened_file)}")
                        self.update_status("工作结束前保存最后打开的文件...")

                        # 立即执行保存操作
                        try:
                            self.perform_save_only(work_end_time)
                            self.log_info("工作结束保存", f"成功保存最后打开的文件: {os.path.basename(self.current_opened_file)}")
                            self.update_status("最后文件已保存")
                        except Exception as e:
                            self.log_error("工作结束保存失败", f"保存文件时出错: {str(e)}")
                            self.update_status("最后文件保存失败，但继续结束工作")

                        # 清除保存时间，避免重复保存
                        self.actual_save_time = None
                        self.update_save_time()

                        # 等待一小段时间确保保存完成
                        time.sleep(2)

                self.update_status("工作时间结束")

                # 根据配置决定是否工作结束后关闭所有软件
                if self.auto_close_on_work_end:
                    self.close_all_programs()
                    self.update_status("工作结束，已关闭相关软件")
                else:
                    self.update_status("工作结束，根据配置未自动关闭软件")

                break

            # 检查是否在午休时间内
            if self.is_lunch_time(now):
                # 获取午休时间信息用于显示
                lunch_start_time, lunch_end_time = self.get_random_lunch_times(now)
                if lunch_start_time and lunch_end_time:
                    # ✨ 在进入午休时间时，如果有待保存的文件，先立即保存
                    if self.current_opened_file and self.actual_save_time:
                        self.update_status("进入午休时间前，先保存当前文件...")
                        self.log_info("午休时间处理", f"检测到有待保存文件 {os.path.basename(self.current_opened_file)}，立即保存")
                        self.perform_save_only()
                        # 清除保存时间，避免午休后重复保存
                        self.actual_save_time = None
                        self.update_save_time()

                    self.update_status(f"午休时间 ({lunch_start_time.strftime('%H:%M')} - {lunch_end_time.strftime('%H:%M')})，暂停文件操作...")

                    # 在午休时间内，暂停所有文件操作，但不关闭应用
                    self.log_info("午休时间", f"进入午休时间，暂停文件操作直到 {lunch_end_time.strftime('%H:%M')}")

                    # 等待午休结束
                    while self.running and now < lunch_end_time:
                        if self.cancel_event.is_set():
                            break

                        # 更新状态显示剩余午休时间
                        remaining_minutes = int((lunch_end_time - now).total_seconds() / 60)
                        if remaining_minutes > 0:
                            self.update_status(f"午休中，还有 {remaining_minutes} 分钟恢复工作...")
                        else:
                            self.update_status("午休即将结束...")

                        # 等待1分钟或直到午休结束
                        wait_seconds = min(60, (lunch_end_time - now).total_seconds())
                        if wait_seconds > 0:
                            if self.wait_with_cancel(wait_seconds):
                                break

                        now = datetime.datetime.now()

                    # 午休结束，恢复工作
                    if self.running and not self.cancel_event.is_set():
                        self.update_status("午休结束，恢复文件操作...")
                        self.log_info("午休时间", "午休结束，恢复正常工作模式")

                # 继续下一轮循环检查
                continue

            # 检查是否需要打开下一个文件
            if valid_folders and self.next_file_open_time:
                if now >= self.next_file_open_time:
                    # 检查距离工作结束时间是否足够进行一次完整的文件操作
                    time_until_end = (work_end_time - now).total_seconds() / 60  # 转换为分钟
                    min_save_delay = max(0.1, self.save_delay_min.get())  # 最小保存延迟

                    # 如果距离结束时间小于最小保存延迟时间 + 1分钟缓冲时间，不再打开新文件
                    if time_until_end > (min_save_delay + 1):
                        # 时间充足，可以打开下一个文件
                        self.open_random_file()
                        last_file_open_time = now

                        # 为新文件设置保存延迟
                        self.schedule_save_for_current_file(work_end_time)

                        # 计算并显示下一次文件打开时间
                        self.calculate_next_file_open_time(work_end_time)

                        self.log_info("文件打开", f"打开新文件，距离工作结束还有 {time_until_end:.1f} 分钟")
                    else:
                        # 时间不足，不再打开新文件
                        self.log_info("文件打开跳过", f"距离工作结束仅剩 {time_until_end:.1f} 分钟，停止打开新文件")
                        self.update_status(f"工作即将结束（{time_until_end:.1f}分钟），不再打开新文件")
                        self.next_file_open_time = None  # 清除下次打开时间

            # 检查是否有待执行的保存操作
            if self.actual_save_time and now >= self.actual_save_time:
                # 检查保存时间是否已超过工作结束时间
                if self.actual_save_time <= work_end_time:
                    self.update_status("执行计划的保存操作...")
                    self.perform_save_only(work_end_time)
                else:
                    self.log_warning("保存操作跳过", f"保存时间 {self.actual_save_time.strftime('%H:%M:%S')} 超过工作结束时间 {work_end_time.strftime('%H:%M:%S')}")
                    self.update_status("保存时间超过工作结束时间，跳过保存操作")

                self.actual_save_time = None
                self.update_save_time()

            # 使用wait_with_cancel代替简单的sleep，以便更快响应停止按钮
            if self.wait_with_cancel(1):  # 等待1秒，但可以被取消
                break

    def schedule_save_for_current_file(self, work_end_time=None):
        """为当前文件安排保存时间"""
        save_delay_min = max(0.1, self.save_delay_min.get())  # 最小0.1分钟（6秒）
        save_delay_max = max(save_delay_min, self.save_delay_max.get())

        # 生成随机延迟
        save_delay_minutes = random.uniform(save_delay_min, save_delay_max)
        save_delay_seconds = int(save_delay_minutes * 60)

        proposed_save_time = datetime.datetime.now() + datetime.timedelta(seconds=save_delay_seconds)

        # 如果提供了工作结束时间，确保保存时间不超过工作结束时间
        if work_end_time and proposed_save_time > work_end_time:
            # 调整保存时间为工作结束前30秒
            adjusted_save_time = work_end_time - datetime.timedelta(seconds=30)
            current_time = datetime.datetime.now()

            # 确保调整后的时间仍然在当前时间之后
            if adjusted_save_time > current_time:
                self.actual_save_time = adjusted_save_time
                self.log_info("保存时间调整", f"原计划: {proposed_save_time.strftime('%H:%M:%S')}, 调整为: {adjusted_save_time.strftime('%H:%M:%S')} (工作结束前)")
            else:
                # 如果调整后的时间已经过了，则立即设置为当前时间后5秒
                self.actual_save_time = current_time + datetime.timedelta(seconds=5)
                self.log_info("保存时间调整", f"工作即将结束，立即安排保存: {self.actual_save_time.strftime('%H:%M:%S')}")
        else:
            self.actual_save_time = proposed_save_time

        # 启动实时倒计时
        remaining_seconds = int((self.actual_save_time - datetime.datetime.now()).total_seconds())
        self.start_save_countdown(max(1, remaining_seconds))

        self.update_save_time()

    def start_save_countdown(self, total_seconds):
        """启动实时保存倒计时"""
        def countdown_update():
            if not self.running or self.cancel_event.is_set():
                return

            now = datetime.datetime.now()
            if self.actual_save_time and now < self.actual_save_time:
                remaining_seconds = int((self.actual_save_time - now).total_seconds())

                if remaining_seconds > 0:
                    if remaining_seconds >= 60:
                        mins = remaining_seconds // 60
                        secs = remaining_seconds % 60
                        self.update_status(f"文件已打开，{mins}分{secs}秒后保存")
                    else:
                        self.update_status(f"文件已打开，{remaining_seconds}秒后保存")

                    # 1秒后再次更新
                    self.root.after(1000, countdown_update)
                else:
                    # 倒计时结束，准备保存
                    self.update_status("保存延迟结束，准备保存...")

        # 开始倒计时
        countdown_update()

    def wait_with_cancel(self, seconds, work_start_time=None):
        """精确等待指定秒数，支持取消操作，可动态更新日期显示"""
        start_time = datetime.datetime.now()
        end_time = start_time + datetime.timedelta(seconds=seconds)

        while datetime.datetime.now() < end_time:
            if self.cancel_event.is_set() or not self.running:
                return True

            # 计算剩余时间
            remaining = (end_time - datetime.datetime.now()).total_seconds()
            if remaining <= 0:
                break

            # 每10秒更新一次保存时间显示和状态
            sleep_time = min(1, remaining)
            time.sleep(sleep_time)

            # 每30秒更新一次显示（减少频率）
            if int(remaining) % 30 == 0:
                if self.actual_save_time:
                    self.update_save_time()

                # 如果提供了工作开始时间，动态更新状态显示
                if work_start_time:
                    current_now = datetime.datetime.now()
                    # 重新计算日期显示
                    if work_start_time.date() > current_now.date():
                        days_diff = (work_start_time.date() - current_now.date()).days
                        if days_diff == 1:
                            # 检查是否还是明天
                            tomorrow_date = current_now.date() + datetime.timedelta(days=1)
                            if work_start_time.date() == tomorrow_date:
                                day_text = "明天"
                            else:
                                day_text = work_start_time.strftime('%m-%d')
                        else:
                            day_text = f"{days_diff}天后"

                        # 检查今天是否为工作日
                        today_is_work_day = self.is_work_day(current_now.date())
                        if today_is_work_day:
                            status_text = f"目前非工作时间，等待到{day_text} {work_start_time.strftime('%H:%M')} 开始工作..."
                        else:
                            status_text = f"今天非工作日，等待到{day_text} {work_start_time.strftime('%H:%M')} 开始工作..."
                    elif work_start_time.date() == current_now.date():
                        # 已经是同一天了
                        status_text = f"等待到今天 {work_start_time.strftime('%H:%M')} 开始工作..."
                    else:
                        # 工作时间已经过了（理论上不应该发生）
                        status_text = f"等待到 {work_start_time.strftime('%m-%d %H:%M')} 开始工作..."

                    self.update_status(status_text)
                else:
                    # 更新状态以保持心跳监控活跃
                    current_status = self.status_label.cget("text")
                    if "等待到" in current_status:
                        self.update_status(current_status, self.status_label.cget("fg"))

        return False

    def run_once(self):
        """立即执行一次 - 直接打开文件并保存"""
        # 验证输入
        if not self.validate_all_inputs():
            return

        # 检查是否至少配置了一个文件夹
        valid_folders = [var.get().strip() for var in self.folder_vars if var.get().strip()]
        if not valid_folders:
            messagebox.showwarning("警告", "请至少添加一个项目文件夹")
            return

        # 验证文件夹是否存在
        for folder_path in valid_folders:
            if not os.path.exists(folder_path):
                self.log_error("文件夹验证失败", f"文件夹不存在: {folder_path}")
                messagebox.showwarning("警告", f"文件夹不存在: {folder_path}")
                return

        self.running = True
        thread = threading.Thread(target=self.perform_full_operation, daemon=True)
        thread.start()

    def perform_full_operation(self):
        """立即执行一次完整操作（模拟工作流程）"""
        self.update_status("开始执行一次完整的工作流程...")
        self.log_info("立即执行模式", "开始执行一次完整操作流程")

        try:
            # 检查是否有有效的文件夹
            valid_folders = [var.get().strip() for var in self.folder_vars if var.get().strip()]
            if valid_folders:
                self.update_status("正在随机打开文档文件...")
                self.log_info("立即执行模式", "准备随机打开文档文件")

                if self.open_random_file():
                    # 记录文件打开成功
                    self.log_info("立即执行模式", f"成功打开文件: {self.current_opened_file}")

                    # 使用新的保存调度方法
                    self.schedule_save_for_current_file()

                    # 计算保存延迟时间用于显示
                    save_delay_min = max(0.1, self.save_delay_min.get())
                    save_delay_max = max(save_delay_min, self.save_delay_max.get())
                    save_delay_minutes = random.uniform(save_delay_min, save_delay_max)
                    save_delay_seconds = int(save_delay_minutes * 60)

                    self.log_info("立即执行模式", f"等待保存延迟: {save_delay_minutes:.2f}分钟 ({save_delay_seconds}秒)")
                    self.update_status(f"等待保存延迟: {save_delay_seconds}秒...")

                    # 等待保存时间到达
                    while self.actual_save_time and datetime.datetime.now() < self.actual_save_time:
                        if self.cancel_event.is_set():
                            self.update_status("延迟已取消", "orange")
                            self.log_warning("立即执行模式", "保存延迟被用户取消")
                            self.running = False
                            return
                        time.sleep(1)

                    self.update_status("保存延迟结束，准备执行保存...")
                    self.log_info("立即执行模式", "保存延迟时间到达，开始执行保存操作")
                else:
                    self.log_warning("立即执行模式", "未能成功打开任何文件")

            # 执行保存操作
            self.perform_save_only()

            # 延迟一秒后关闭软件
            self.update_status("保存完成，准备关闭软件...")
            self.log_info("立即执行模式", "保存操作完成，准备关闭相关软件")
            time.sleep(1)

            # 执行关闭软件操作
            self.close_opened_programs()

            self.update_status("一次完整操作执行完毕")
            self.log_info("立即执行模式", "完整操作流程执行完毕")
            self.root.after(0, lambda: messagebox.showinfo("完成", "一次完整操作已执行完毕\n已完成：文件打开 → 保存延迟 → 执行保存 → 关闭软件"))

        except Exception as e:
            self.update_status(f"操作失败: {str(e)}", "red")
            self.log_error("立即执行模式", f"操作执行失败: {str(e)}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"操作失败: {str(e)}"))
        finally:
            # 执行完毕后清理
            self.running = False
            self.actual_save_time = None
            self.current_opened_file = None
            self.current_opened_time = None
            self.update_save_time()

    def perform_save_only(self, work_end_time=None):
        """执行保存操作 - 只保存当前打开的文件"""
        try:
            saved_count = self.save_current_opened_file()

            # ✨ 保存完成后清除保存时间，避免重复保存
            self.actual_save_time = None
            self.update_save_time()

            if saved_count > 0:
                self.update_status(f"保存操作完毕，已保存 {saved_count} 个文档")
                # 保存完成后，显示下一次打开文件的时间
                self.show_next_file_open_time(work_end_time)
            else:
                self.update_status("没有需要保存的文档文件")
        except Exception as e:
            self.log_error("保存操作失败", f"错误: {str(e)}")
            self.update_status(f"保存操作失败: {str(e)}", "red")

    def save_current_opened_file(self):
        """只保存当前打开的文件"""
        if not self.current_opened_file:
            self.log_warning("保存操作", "没有当前打开的文件记录")
            return 0

        self.update_status("正在保存当前打开的文档文件...")
        saved_count = 0

        try:
            # 检查pyautogui是否可用
            if not PYAUTOGUI_AVAILABLE:
                self.log_warning("保存操作", "pyautogui不可用，无法执行保存操作")
                self.update_status("pyautogui不可用，跳过保存操作", "orange")
                return 0

            # 获取所有窗口 - 添加错误处理
            try:
                all_windows = pyautogui.getAllWindows()
            except Exception as e:
                self.log_error("获取窗口列表失败", f"错误: {str(e)}")
                self.update_status("无法获取窗口列表，跳过保存操作", "orange")
                return 0

            current_file_name = os.path.basename(self.current_opened_file)
            current_file_name_without_ext = os.path.splitext(current_file_name)[0]

            self.log_info("开始保存当前文件", f"目标文件: {current_file_name} | 文件路径: {self.current_opened_file}")

            # 寻找与当前文件相关的窗口
            for window in all_windows:
                try:
                    window_title = window.title

                    # 跳过空标题窗口
                    if not window_title.strip():
                        continue

                    # 检查窗口标题是否包含当前文件名
                    if (current_file_name.lower() in window_title.lower() or
                        current_file_name_without_ext.lower() in window_title.lower()):

                        try:
                            # 激活窗口并保存 - 添加错误处理
                            try:
                                window.activate()
                                time.sleep(0.5)
                            except Exception as e:
                                self.log_warning("窗口激活失败", f"窗口: {window_title[:50]} | 错误: {str(e)}")
                                continue

                            try:
                                pyautogui.hotkey('ctrl', 's')
                                saved_count += 1
                            except Exception as e:
                                self.log_warning("保存快捷键失败", f"窗口: {window_title[:50]} | 错误: {str(e)}")
                                continue

                            # 获取清理后的窗口标题（避免过长的标题）
                            clean_title = self.clean_window_title(window_title)

                            # 记录保存操作日志
                            self.log_info("保存文档", f"窗口: {clean_title} | 文件: {current_file_name} | 文档已保存")

                            self.update_status(f"已保存文档: {clean_title}")
                            time.sleep(1)

                        except Exception as e:
                            clean_title = self.clean_window_title(window_title)
                            self.log_warning("保存文档失败", f"窗口: {clean_title} | 文件: {current_file_name} | 错误: {str(e)}")
                            self.update_status(f"保存窗口 '{clean_title}' 失败: {str(e)}", "orange")

                except Exception as e:
                    self.log_warning("处理窗口失败", f"错误: {str(e)}")
                    continue

            if saved_count > 0:
                self.log_info("当前文件保存完成", f"成功保存 {saved_count} 个与文件 {current_file_name} 相关的窗口")
            else:
                self.log_warning("保存操作", f"未找到与文件 {current_file_name} 相关的窗口")

            return saved_count

        except Exception as e:
            self.log_error("保存当前文件失败", f"文件: {self.current_opened_file} | 错误: {str(e)}")
            raise Exception(f"保存当前文件失败: {str(e)}")

    def clean_window_title(self, title):
        """清理窗口标题，避免显示过长或包含无关信息的标题"""
        if not title:
            return "未知窗口"

        # 限制标题长度
        if len(title) > 50:
            title = title[:47] + "..."

        # 移除一些常见的无关后缀
        unwanted_suffixes = [
            " - Microsoft Word",
            " - Microsoft Excel",
            " - Microsoft PowerPoint",
            " - WPS Writer",
            " - WPS Spreadsheets",
            " - WPS Presentation"
        ]

        for suffix in unwanted_suffixes:
            if title.endswith(suffix):
                title = title[:-len(suffix)]
                break

        return title.strip()

    def show_next_file_open_time(self, work_end_time=None):
        """显示下一次打开文件的时间（使用已计算的时间）"""
        if not self.running or not self.next_file_open_time:
            return

        try:
            now = datetime.datetime.now()
            time_diff = (self.next_file_open_time - now).total_seconds() / 60  # 转换为分钟

            if time_diff > 0:
                self.update_status(f"下一次文件打开时间: {self.next_file_open_time.strftime('%H:%M:%S')}")
            else:
                self.update_status("即将打开下一个文件...")

        except Exception as e:
            self.log_warning("显示下次文件时间失败", f"错误: {str(e)}")

    def calculate_next_file_open_time(self, work_end_time=None):
        """计算并设置下一次文件打开时间"""
        if not self.running:
            return

        try:
            # 计算下一次文件打开的时间间隔
            file_interval_min = max(0.1, self.file_interval_min.get())
            file_interval_max = max(file_interval_min, self.file_interval_max.get())
            next_interval_minutes = random.uniform(file_interval_min, file_interval_max)

            next_file_time = datetime.datetime.now() + datetime.timedelta(minutes=next_interval_minutes)

            # 如果提供了工作结束时间，检查下一次打开时间是否会导致来不及保存
            if work_end_time:
                # 计算最小保存延迟时间 + 缓冲时间
                min_save_delay = max(0.1, self.save_delay_min.get())
                buffer_time = 1  # 1分钟缓冲时间
                required_time = min_save_delay + buffer_time

                # 检查下一次打开时间 + 必需的操作时间是否超过工作结束时间
                if next_file_time + datetime.timedelta(minutes=required_time) > work_end_time:
                    # 时间不够完成一次完整的文件操作，不设置下次打开时间
                    self.next_file_open_time = None
                    self.log_info("下次文件打开计划",
                        f"原计划时间: {next_file_time.strftime('%H:%M:%S')} + 必需操作时间 {required_time:.1f}分钟 "
                        f"将超过工作结束时间 {work_end_time.strftime('%H:%M:%S')}，已取消")
                    self.update_status("临近工作结束时间，不再计划新的文件打开")
                    return

            # 设置下一次打开时间
            self.next_file_open_time = next_file_time

            # 显示下一次打开时间
            self.update_status(f"下一次文件打开时间: {next_file_time.strftime('%H:%M:%S')} ({next_interval_minutes:.1f}分钟后)")
            self.log_info("下次文件打开计划", f"预计时间: {next_file_time.strftime('%H:%M:%S')} | 间隔: {next_interval_minutes:.1f}分钟")

        except Exception as e:
            self.log_warning("计算下次文件时间失败", f"错误: {str(e)}")
            self.next_file_open_time = None

    def save_documents_in_all_folders(self):
        """保存当前打开的所有文档文件"""
        self.update_status("正在保存打开的文档文件...")

        try:
            # 获取所有窗口
            all_windows = pyautogui.getAllWindows()
            saved_count = 0

            # 定义常见的文档编辑器和办公软件
            document_apps = [
                "notepad", "记事本", "wordpad", "写字板",
                "microsoft word", "word", "excel", "powerpoint",
                "wps", "金山", "sublime", "atom", "brackets",
                "adobe", "pdf", "foxit", "福昕", "visual studio code",
                "pycharm", "intellij", "eclipse", "dev-c++"
            ]

            for window in all_windows:
                window_title = window.title.lower()
                # 检查是否是文档编辑器窗口
                is_document_window = any(app in window_title for app in document_apps)

                # 或者检查窗口标题是否包含配置文件夹中的文件名
                is_folder_file = False
                for folder_var in self.folder_vars:
                    folder_path = folder_var.get().strip()
                    if not folder_path or not os.path.exists(folder_path):
                        continue
                    try:
                        for root, dirs, files in os.walk(folder_path):
                            for file in files:
                                file_name_without_ext = os.path.splitext(file)[0].lower()
                                if file_name_without_ext in window_title:
                                    is_folder_file = True
                                    break
                            if is_folder_file:
                                break
                    except:
                        pass
                    if is_folder_file:
                        break

                if is_document_window or is_folder_file:
                    try:
                        # 激活窗口并保存
                        window.activate()
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 's')
                        saved_count += 1

                        # 记录保存操作日志
                        self.log_info("保存文档", f"窗口: {window.title[:50]} | 文档已保存")

                        self.update_status(f"已保存文档: {window.title[:50]}...")
                        time.sleep(1)
                    except Exception as e:
                        self.log_warning("保存文档失败", f"窗口: {window.title[:30]} | 错误: {str(e)}")
                        self.update_status(f"保存窗口 '{window.title[:30]}...' 失败: {str(e)}", "orange")

            if saved_count > 0:
                self.log_info("批量保存完成", f"成功保存 {saved_count} 个文档文件")
                self.update_status(f"成功保存 {saved_count} 个文档文件")
            else:
                self.log_warning("保存操作", "未找到需要保存的文档文件")
                self.update_status("未找到需要保存的文档文件", "orange")

        except Exception as e:
            raise Exception(f"文档保存失败: {str(e)}")

    def update_status(self, text, color="black"):
        # 记录状态更新时间（用于心跳监控）
        self.last_status_update_time = datetime.datetime.now()
        self.root.after(0, lambda: self.status_label.config(text=text, fg=color))

        # 只记录真正重要的状态变化，避免日志污染
        important_keywords = ["工作时间到", "工作模式", "开始文档操作", "任务已停止", "任务已启动"]

        # 避免重复记录相同的等待状态
        if not hasattr(self, '_last_logged_status'):
            self._last_logged_status = ""

        # 过滤掉重复的等待状态日志
        should_log = False
        if any(keyword in text for keyword in important_keywords):
            should_log = True
        elif "等待到" in text and self._last_logged_status != text:
            # 只有当等待状态发生实际变化时才记录
            should_log = True

        if should_log:
            self.log_info("状态更新", text)
            self._last_logged_status = text

    def update_save_time(self):
        """更新时间显示"""
        if self.actual_save_time:
            save_str = self.actual_save_time.strftime("%Y-%m-%d %H:%M:%S")
            self.root.after(0, lambda: self.save_time_label.config(text=f"下次操作时间: {save_str}"))
        else:
            # 计算下一个工作开始时间
            now = datetime.datetime.now()

            # 获取今天的随机工作时间
            work_start_time, work_end_time = self.get_random_work_times(now)

            # 检查今天是否为工作日，以及当前时间状态
            today_is_work_day = self.is_work_day(now.date())

            # 检查是否已启动自动任务
            task_status_suffix = "" if self.running else " （尚未启动）"

            if now >= work_end_time or not today_is_work_day:
                # 如果已过工作结束时间，或今天不是工作日，寻找下一个工作日
                next_work_day = now.date() + datetime.timedelta(days=1)
                while not self.is_work_day(next_work_day):
                    next_work_day += datetime.timedelta(days=1)

                # 获取下一个工作日的随机工作时间
                next_day_datetime = datetime.datetime.combine(next_work_day, datetime.time(9, 0))
                next_work_start, next_work_end = self.get_random_work_times(next_day_datetime)

                # 计算天数差异以显示更友好的信息
                days_diff = (next_work_day - now.date()).days
                if days_diff == 1:
                    # 动态判断是否为明天
                    tomorrow_date = now.date() + datetime.timedelta(days=1)
                    if next_work_day == tomorrow_date:
                        day_text = "明天"
                    else:
                        day_text = next_work_start.strftime('%m-%d')
                    status_text = f"下次工作时间: {day_text} {next_work_start.strftime('%H:%M')} - {next_work_end.strftime('%H:%M')}{task_status_suffix}"
                elif days_diff <= 7:
                    weekday_names = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']
                    weekday_name = weekday_names[next_work_day.weekday()]
                    status_text = f"下次工作时间: {weekday_name} {next_work_start.strftime('%m-%d %H:%M')} - {next_work_end.strftime('%H:%M')}{task_status_suffix}"
                else:
                    status_text = f"下次工作时间: {next_work_start.strftime('%Y-%m-%d %H:%M')} - {next_work_end.strftime('%H:%M')}{task_status_suffix}"

            elif now < work_start_time and today_is_work_day:
                # 今天是工作日且还没到工作开始时间
                status_text = f"今日工作时间: {work_start_time.strftime('%H:%M')} - {work_end_time.strftime('%H:%M')}{task_status_suffix}"
            else:
                # 当前在工作时间内
                if self.running:
                    status_text = f"工作中，结束时间: {work_end_time.strftime('%H:%M')}"
                else:
                    status_text = f"工作时间内，结束时间: {work_end_time.strftime('%H:%M')}{task_status_suffix}"

            self.root.after(0, lambda: self.save_time_label.config(text=status_text))

    def save_config(self):
        try:
            # 收集当前的文件夹路径
            project_folders = [var.get().strip() for var in self.folder_vars if var.get().strip()]

            config = {
                # 基本配置
                "project_folders": project_folders,  # 新的多文件夹配置
                "work_start_hour": self.work_start_hour.get(),
                "work_start_minute": self.work_start_minute.get(),
                "work_end_hour": self.work_end_hour.get(),
                "work_end_minute": self.work_end_minute.get(),
                "work_time_random_range": getattr(self, 'work_time_random_range', 20),

                # 午休时间配置
                "lunch_break": {
                    "enabled": getattr(self, 'lunch_break_enabled', True),
                    "start_hour": getattr(self, 'lunch_start_hour', 12),
                    "start_minute": getattr(self, 'lunch_start_minute', 0),
                    "end_hour": getattr(self, 'lunch_end_hour', 13),
                    "end_minute": getattr(self, 'lunch_end_minute', 30),
                    "random_range": getattr(self, 'lunch_time_random_range', 5)
                },

                "save_delay_min": self.save_delay_min.get(),
                "save_delay_max": self.save_delay_max.get(),
                "file_interval_min": self.file_interval_min.get(),
                "file_interval_max": self.file_interval_max.get(),

                # 智能关闭功能配置
                "auto_close_on_work_end": getattr(self, 'auto_close_on_work_end', True),
                "close_strategy": {
                    "use_alt_f4": getattr(self, 'use_alt_f4', True),
                    "use_ctrl_q": getattr(self, 'use_ctrl_q', True),
                    "use_ctrl_w": getattr(self, 'use_ctrl_w', True),
                    "close_timeout": getattr(self, 'close_timeout', 3.0)
                },

                # 文件跟踪配置
                "file_tracking": {
                    "enabled": getattr(self, 'file_tracking_enabled', True),
                    "track_program_mapping": getattr(self, 'track_program_mapping', True),
                    "clear_tracking_on_stop": getattr(self, 'clear_tracking_on_stop', True)
                },

                # 程序检测配置
                "program_detection": {
                    "window_check_interval": getattr(self, 'window_check_interval', 1.0),
                    "activation_delay": getattr(self, 'activation_delay', 0.5),
                    "close_verification_delay": getattr(self, 'close_verification_delay', 1.5)
                },

                # 用户界面配置
                "ui_settings": {
                    "show_close_progress": getattr(self, 'show_close_progress', True),
                    "show_detected_programs": getattr(self, 'show_detected_programs', True),
                    "status_update_interval": getattr(self, 'status_update_interval', 1.0)
                },

                # 工作日历配置
                "work_calendar": {
                    "skip_weekends": getattr(self, 'skip_weekends', True),
                    "work_dates": getattr(self, 'work_dates', []),
                    "holiday_dates": getattr(self, 'holiday_dates', [])
                },

                # 日志功能配置
                "logging": {
                    "enabled": getattr(self, 'logging_enabled', False),
                    "log_file_path": getattr(self, 'log_file_path', ""),
                    "log_level": getattr(self, 'log_level', "INFO"),
                    "log_max_size": getattr(self, 'log_max_size', 10),
                    "log_backup_count": getattr(self, 'log_backup_count', 5)
                },

                # 文件过滤配置
                "file_filtering": {
                    "allowed_extensions": getattr(self, 'allowed_file_extensions', [
                        '.txt', '.docx', '.doc', '.pdf', '.wps',
                        '.py', '.java', '.cpp', '.html', '.js',
                        '.md', '.rtf', '.odt', '.xlsx', '.xls',
                        '.pptx', '.ppt', '.css', '.json', '.xml',
                        '.php', '.c', '.h', '.cs', '.go', '.vue', 'xmind'
                    ]),
                    "scan_subfolders": getattr(self, 'scan_subfolders', False)
                }
            }

            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)

            with open(self.config_path, "w", encoding="utf-8") as f:
                json.dump(config, f, ensure_ascii=False, indent=2)

            return True
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {str(e)}")
            return False

    def load_config(self):
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)

                # 加载项目文件夹配置
                project_folders = config.get("project_folders", [])

                # 确保folder_vars已初始化
                if not hasattr(self, 'folder_vars'):
                    self.folder_vars = []

                # 清空现有的文件夹变量
                self.folder_vars.clear()

                # 创建对应数量的StringVar并设置值
                for folder_path in project_folders:
                    folder_var = StringVar()
                    folder_var.set(folder_path)
                    self.folder_vars.append(folder_var)

                # 加载工作时间设置
                self.work_start_hour.set(config.get("work_start_hour", 9))
                self.work_start_minute.set(config.get("work_start_minute", 0))
                self.work_end_hour.set(config.get("work_end_hour", 18))
                self.work_end_minute.set(config.get("work_end_minute", 0))

                # 加载工作时间随机区间设置
                self.work_time_random_range = config.get("work_time_random_range", 20)

                # 加载午休时间配置
                lunch_break_config = config.get("lunch_break", {})
                self.lunch_break_enabled = lunch_break_config.get("enabled", True)
                self.lunch_start_hour = lunch_break_config.get("start_hour", 12)
                self.lunch_start_minute = lunch_break_config.get("start_minute", 0)
                self.lunch_end_hour = lunch_break_config.get("end_hour", 13)
                self.lunch_end_minute = lunch_break_config.get("end_minute", 30)
                self.lunch_time_random_range = lunch_break_config.get("random_range", 5)

                # 加载延迟设置
                self.save_delay_min.set(config.get("save_delay_min", 20))
                self.save_delay_max.set(config.get("save_delay_max", 50))

                self.file_interval_min.set(config.get("file_interval_min", 30))
                self.file_interval_max.set(config.get("file_interval_max", 60))

                # 加载智能关闭功能配置
                self.auto_close_on_work_end = config.get("auto_close_on_work_end", True)

                close_strategy = config.get("close_strategy", {})
                self.use_alt_f4 = close_strategy.get("use_alt_f4", True)
                self.use_ctrl_q = close_strategy.get("use_ctrl_q", True)
                self.use_ctrl_w = close_strategy.get("use_ctrl_w", True)
                self.close_timeout = close_strategy.get("close_timeout", 3.0)

                # 加载文件跟踪配置
                file_tracking = config.get("file_tracking", {})
                self.file_tracking_enabled = file_tracking.get("enabled", True)
                self.track_program_mapping = file_tracking.get("track_program_mapping", True)
                self.clear_tracking_on_stop = file_tracking.get("clear_tracking_on_stop", True)

                # 加载程序检测配置
                program_detection = config.get("program_detection", {})
                self.window_check_interval = program_detection.get("window_check_interval", 1.0)
                self.activation_delay = program_detection.get("activation_delay", 0.5)
                self.close_verification_delay = program_detection.get("close_verification_delay", 1.5)

                # 加载用户界面配置
                ui_settings = config.get("ui_settings", {})
                self.show_close_progress = ui_settings.get("show_close_progress", True)
                self.show_detected_programs = ui_settings.get("show_detected_programs", True)
                self.status_update_interval = ui_settings.get("status_update_interval", 1.0)

                # 加载工作日历配置
                work_calendar = config.get("work_calendar", {})
                self.skip_weekends = work_calendar.get("skip_weekends", True)
                self.work_dates = work_calendar.get("work_dates", [])
                self.holiday_dates = work_calendar.get("holiday_dates", [])

                # 加载日志功能配置
                logging_config = config.get("logging", {})
                self.logging_enabled = logging_config.get("enabled", False)
                self.log_file_path = logging_config.get("log_file_path", "")
                self.log_level = logging_config.get("log_level", "INFO")
                self.log_max_size = logging_config.get("log_max_size", 10)
                self.log_backup_count = logging_config.get("log_backup_count", 5)

                # 加载文件过滤配置
                file_filtering = config.get("file_filtering", {})
                self.allowed_file_extensions = file_filtering.get("allowed_extensions", [
                    '.txt', '.docx', '.doc', '.pdf', '.wps',
                    '.py', '.java', '.cpp', '.html', '.js',
                    '.md', '.rtf', '.odt', '.xlsx', '.xls',
                    '.pptx', '.ppt', '.css', '.json', '.xml',
                    '.php', '.c', '.h', '.cs', '.go', '.rs'
                ])

                # 加载文件扫描控制配置
                self.scan_subfolders = file_filtering.get("scan_subfolders", False)

                self.log_info("配置加载完成", f"项目文件夹数量: {len(self.folder_vars)}")
                self.log_info("工作日历配置", f"跳过周末: {self.skip_weekends}, 调休日期: {len(self.work_dates)}个, 节假日: {len(self.holiday_dates)}个")
                self.log_info("午休时间配置", f"启用: {self.lunch_break_enabled}, 时间: {self.lunch_start_hour:02d}:{self.lunch_start_minute:02d}-{self.lunch_end_hour:02d}:{self.lunch_end_minute:02d}, 随机区间: {self.lunch_time_random_range}分钟")
                self.log_info("日志功能配置", f"启用: {self.logging_enabled}, 级别: {self.log_level}")
                self.log_info("文件过滤配置", f"允许的文件后缀: {len(self.allowed_file_extensions)}个")
                self.log_info("文件扫描配置", f"递归扫描: {self.scan_subfolders}")

        except Exception as e:
            self.log_error("配置加载失败", f"错误: {str(e)}")
            # 确保即使配置加载失败，folder_vars也是初始化的
            if not hasattr(self, 'folder_vars'):
                self.folder_vars = []

            # 显示警告但不中断程序运行
            try:
                messagebox.showwarning("警告", f"加载配置失败: {str(e)}\n将使用默认设置")
            except:
                # 如果无法显示消息框，将错误记录到日志或输出到控制台（仅在调试时）
                pass

    def close_all_programs(self):
        """智能关闭所有实际打开的软件"""
        try:
            self.update_status("正在检查需要关闭的软件...")

            # 获取实际运行的程序列表
            running_programs = self.get_actually_running_programs()

            if not running_programs:
                self.log_info("关闭软件操作", "没有检测到需要关闭的软件")
                self.update_status("没有检测到需要关闭的软件")
                return

            self.log_info("开始关闭软件", f"检测到 {len(running_programs)} 个需要关闭的程序")

            closed_programs = []
            if self.show_detected_programs:
                self.update_status(f"检测到 {len(running_programs)} 个需要关闭的程序...")
            else:
                self.update_status("开始关闭检测到的程序...")

            # 逐个关闭检测到的程序
            for program_info in running_programs:
                program_name = program_info['name']
                window = program_info['window']

                try:
                    if self.show_close_progress:
                        self.update_status(f"正在关闭: {program_name}")

                    # 激活窗口
                    window.activate()
                    time.sleep(self.activation_delay)

                    # 根据配置使用不同的关闭策略
                    if self.use_alt_f4:
                        pyautogui.hotkey('alt', 'f4')
                        time.sleep(self.close_verification_delay)

                    # 检查窗口是否还存在
                    try:
                        # 检查原窗口是否还存在
                        if window in pyautogui.getAllWindows():
                            if self.use_ctrl_q:
                                # 如果Alt+F4无效，尝试Ctrl+Q
                                pyautogui.hotkey('ctrl', 'q')
                                time.sleep(self.close_timeout)

                            # 再次检查
                            if window in pyautogui.getAllWindows() and self.use_ctrl_w:
                                # 如果还是无效，尝试Ctrl+W关闭当前标签页/文档
                                pyautogui.hotkey('ctrl', 'w')
                                time.sleep(self.close_timeout)
                    except:
                        pass  # 窗口可能已经关闭，这是正常的

                    closed_programs.append(program_name)
                    self.log_info("成功关闭程序", f"程序: {program_name}")

                except Exception as e:
                    self.log_error("关闭程序失败", f"程序: {program_name} | 错误: {str(e)}")
                    self.update_status(f"关闭 {program_name} 时出错: {str(e)}", "orange")

            if closed_programs:
                self.log_info("批量关闭完成", f"成功关闭 {len(closed_programs)} 个软件: {', '.join(closed_programs)}")
                self.update_status(f"成功关闭 {len(closed_programs)} 个软件: {', '.join(closed_programs[:3])}" +
                                 (f" 等" if len(closed_programs) > 3 else ""))
            else:
                self.log_warning("关闭软件操作", "没有成功关闭任何软件")
                self.update_status("没有成功关闭任何软件")

            # 在关闭软件操作完成后，根据配置决定是否清空跟踪列表
            if self.clear_tracking_on_stop:
                if self.opened_files or self.opened_programs:
                    self.log_info("清理跟踪信息", f"清空文件跟踪列表: {len(self.opened_files)}个文件, {len(self.opened_programs)}个程序")
                self.opened_files.clear()
                self.opened_programs.clear()

        except Exception as e:
            self.log_error("关闭软件时出错", f"错误: {str(e)}")
            self.update_status(f"关闭软件时出错: {str(e)}", "orange")

    def close_opened_programs(self):
        """关闭当前打开的文件相关程序（立即执行模式专用）"""
        try:
            self.update_status("正在关闭当前打开的文件相关程序...")
            self.log_info("立即执行模式", "开始关闭当前打开的文件相关程序")

            if not self.current_opened_file:
                self.log_warning("立即执行模式", "没有检测到当前打开的文件")
                self.update_status("没有检测到当前打开的文件")
                return

            # 获取实际运行的程序列表
            running_programs = self.get_actually_running_programs()

            if not running_programs:
                self.log_info("立即执行模式", "没有检测到需要关闭的程序")
                self.update_status("没有检测到需要关闭的程序")
                return

            # 过滤出与当前文件相关的程序
            current_filename = os.path.splitext(os.path.basename(self.current_opened_file))[0]
            related_programs = []

            for program_info in running_programs:
                program_name = program_info['name']
                window_title = program_info['window'].title.lower()

                # 检查窗口标题是否包含当前文件名
                if current_filename.lower() in window_title:
                    related_programs.append(program_info)

            if not related_programs:
                self.log_warning("立即执行模式", f"没有找到与文件 {self.current_opened_file} 相关的程序窗口")
                self.update_status("没有找到相关的程序窗口")
                return

            self.log_info("立即执行模式", f"找到 {len(related_programs)} 个与文件 {current_filename} 相关的程序")

            closed_programs = []
            # 逐个关闭相关程序
            for program_info in related_programs:
                program_name = program_info['name']
                window = program_info['window']

                try:
                    self.update_status(f"正在关闭: {program_name}")
                    self.log_info("立即执行模式", f"正在关闭程序: {program_name}")

                    # 激活窗口
                    window.activate()
                    time.sleep(self.activation_delay)

                    # 使用关闭策略
                    if self.use_alt_f4:
                        pyautogui.hotkey('alt', 'f4')
                        time.sleep(self.close_verification_delay)

                    # 检查窗口是否还存在，如果存在尝试其他方法
                    try:
                        if window in pyautogui.getAllWindows():
                            if self.use_ctrl_q:
                                pyautogui.hotkey('ctrl', 'q')
                                time.sleep(self.close_timeout)

                            if window in pyautogui.getAllWindows() and self.use_ctrl_w:
                                pyautogui.hotkey('ctrl', 'w')
                                time.sleep(self.close_timeout)
                    except:
                        pass  # 窗口可能已经关闭

                    closed_programs.append(program_name)
                    self.log_info("立即执行模式", f"成功关闭程序: {program_name}")

                except Exception as e:
                    self.log_error("立即执行模式", f"关闭程序失败: {program_name} | 错误: {str(e)}")
                    self.update_status(f"关闭 {program_name} 时出错: {str(e)}", "orange")

            if closed_programs:
                self.log_info("立即执行模式", f"程序关闭完成，成功关闭 {len(closed_programs)} 个相关程序: {', '.join(closed_programs)}")
                self.update_status(f"成功关闭 {len(closed_programs)} 个相关程序")
            else:
                self.log_warning("立即执行模式", "没有成功关闭任何程序")
                self.update_status("没有成功关闭任何程序")

        except Exception as e:
            self.log_error("立即执行模式", f"关闭程序时出错: {str(e)}")
            self.update_status(f"关闭程序时出错: {str(e)}", "orange")

    def close_program(self, program_name):
        """关闭指定名称的程序窗口"""
        try:
            windows = pyautogui.getAllWindows()
            closed = False

            for window in windows:
                if program_name.lower() in window.title.lower():
                    try:
                        # 激活窗口
                        window.activate()
                        time.sleep(0.5)

                        # 尝试使用Alt+F4关闭窗口
                        pyautogui.hotkey('alt', 'f4')
                        time.sleep(1)

                        # 检查窗口是否还存在
                        if not any(program_name.lower() in w.title.lower() for w in pyautogui.getAllWindows()):
                            closed = True
                        else:
                            # 如果Alt+F4无效，尝试Ctrl+Q（某些程序）
                            pyautogui.hotkey('ctrl', 'q')
                            time.sleep(1)

                            if not any(program_name.lower() in w.title.lower() for w in pyautogui.getAllWindows()):
                                closed = True

                    except Exception as e:
                        self.update_status(f"关闭{program_name}窗口失败: {str(e)}", "orange")
                        continue

            return closed

        except Exception as e:
            self.update_status(f"查找{program_name}窗口失败: {str(e)}", "orange")
            return False

    def track_program_by_file_extension(self, file_ext):
        """根据文件扩展名跟踪可能使用的程序"""
        # 文档类型到程序的映射
        extension_to_programs = {
            '.txt': ['记事本', 'Notepad', 'Sublime Text', 'VS Code'],
            '.docx': ['Microsoft Word', 'WPS Writer', 'WPS 文字'],
            '.doc': ['Microsoft Word', 'WPS Writer', 'WPS 文字'],
            '.pdf': ['Adobe Acrobat', 'Foxit Reader', '福昕', 'Microsoft Edge'],
            '.wps': ['WPS Writer', 'WPS 文字'],
            '.py': ['VS Code', 'PyCharm', 'IDLE', 'Sublime Text'],
            '.java': ['IntelliJ IDEA', 'Eclipse', 'VS Code'],
            '.cpp': ['VS Code', 'Dev-C++', 'Code::Blocks'],
            '.html': ['VS Code', 'Sublime Text', 'Chrome', 'Edge'],
            '.js': ['VS Code', 'Sublime Text', 'WebStorm'],
            '.md': ['VS Code', 'Typora', 'MarkdownPad'],
            '.xlsx': ['Microsoft Excel', 'WPS 表格'],
            '.xls': ['Microsoft Excel', 'WPS 表格'],
            '.pptx': ['Microsoft PowerPoint', 'WPS 演示'],
            '.ppt': ['Microsoft PowerPoint', 'WPS 演示']
        }

        if file_ext in extension_to_programs:
            for program in extension_to_programs[file_ext]:
                self.opened_programs.add(program)

    def get_actually_running_programs(self):
        """获取实际正在运行的程序列表（按软件进程分组，避免重复关闭）"""
        try:
            all_windows = pyautogui.getAllWindows()
            software_groups = {}  # 按软件类型分组
            excluded_titles = set()  # 排除不应关闭的窗口

            # 获取当前运行的程序名称（避免关闭自身）
            current_process_name = os.path.basename(sys.executable).lower()
            current_script_name = os.path.basename(__file__).lower() if '__file__' in globals() else 'activity_tracker.py'

            # 定义软件类型识别模式
            software_patterns = {
                'word': ['microsoft word', 'word', 'winword'],
                'excel': ['microsoft excel', 'excel'],
                'powerpoint': ['microsoft powerpoint', 'powerpoint'],
                'wps_writer': ['wps writer', 'wps 文字', 'wps文字'],
                'wps_spreadsheet': ['wps spreadsheets', 'wps 表格', 'wps表格'],
                'wps_presentation': ['wps presentation', 'wps 演示', 'wps演示'],
                'notepad': ['notepad.exe', '记事本'],
                'notepadpp': ['notepad++'],
                'sublime': ['sublime text'],
                'vscode': ['visual studio code', 'vscode'],
                'pycharm': ['pycharm'],
                'adobe_acrobat': ['adobe acrobat', 'acrobat'],
                'foxit': ['foxit reader', '福昕'],
                'chrome': ['google chrome'],
                'edge': ['microsoft edge']
            }

            # 定义不应关闭的软件模式（避免关闭开发环境自身）
            excluded_patterns = [
                'activity_tracker',     # 避免关闭自身（开发环境）
                'worktrace mocker',     # 避免关闭打包后的程序
                'workTraceMocker',      # 避免关闭打包后的程序（驼峰命名）
                'WorkTrace Mocker',
                'python',               # 避免关闭Python相关程序
                'pythonw',              # 避免关闭Python窗口版本
                'cmd.exe',             # 避免关闭命令行
                'powershell',          # 避免关闭PowerShell
                'explorer.exe',        # 避免关闭文件资源管理器
                'taskmgr.exe',         # 避免关闭任务管理器
                'pyinstaller',         # 避免关闭PyInstaller相关进程
                '.exe - python',       # 避免关闭Python相关的exe进程
            ]

            self.log_info("程序检测开始", f"当前进程: {current_process_name}, 脚本: {current_script_name}")

            for window in all_windows:
                window_title = window.title
                window_title_lower = window_title.lower()

                if len(window_title.strip()) == 0:  # 跳过空标题窗口
                    continue

                # 检查是否为不应关闭的软件
                should_exclude = False

                # 首先检查基本排除模式
                for excluded_pattern in excluded_patterns:
                    if excluded_pattern in window_title_lower:
                        should_exclude = True
                        excluded_titles.add(window_title)
                        self.log_info("排除窗口", f"基本排除模式匹配: {window_title} (模式: {excluded_pattern})")
                        break

                # 额外的自身保护检查
                if not should_exclude:
                    # 检查是否包含当前可执行文件名（去除扩展名）
                    current_exe_name = os.path.splitext(current_process_name)[0]
                    if current_exe_name in window_title_lower:
                        should_exclude = True
                        excluded_titles.add(window_title)
                        self.log_info("排除窗口", f"可执行文件名匹配: {window_title} (文件: {current_exe_name})")

                    # 检查是否为包含特定关键词的可能的自身程序
                    self_protection_keywords = ['活动痕迹', '自动活动', 'activity tracker', 'trace mocker']
                    for keyword in self_protection_keywords:
                        if keyword in window_title_lower:
                            should_exclude = True
                            excluded_titles.add(window_title)
                            self.log_info("排除窗口", f"自身保护关键词匹配: {window_title} (关键词: {keyword})")
                            break

                if should_exclude:
                    continue

                # 确定软件类型
                software_type = None
                detected_pattern = None
                matched_file = None  # 初始化匹配的文件名

                for soft_type, patterns in software_patterns.items():
                    for pattern in patterns:
                        if pattern in window_title_lower:
                            software_type = soft_type
                            detected_pattern = pattern
                            break
                    if software_type:
                        break

                # 如果没有匹配到预定义软件类型，检查是否包含我们的文件
                if not software_type:
                    contains_tracked_file = False

                    # 检查是否包含跟踪的文件
                    if self.opened_files:
                        for opened_file in self.opened_files:
                            file_name = os.path.basename(opened_file)
                            file_name_no_ext = os.path.splitext(file_name)[0]
                            if (file_name.lower() in window_title_lower or
                                file_name_no_ext.lower() in window_title_lower):
                                contains_tracked_file = True
                                matched_file = file_name
                                break

                    # 检查是否包含项目文件夹中的文件
                    if not contains_tracked_file:
                        for folder_var in self.folder_vars:
                            folder_path = folder_var.get().strip()
                            if not folder_path or not os.path.exists(folder_path):
                                continue

                            try:
                                for root, dirs, files in os.walk(folder_path):
                                    for file in files:
                                        file_name = os.path.basename(file)
                                        file_name_no_ext = os.path.splitext(file_name)[0]
                                        if (file_name.lower() in window_title_lower or
                                            file_name_no_ext.lower() in window_title_lower):
                                            contains_tracked_file = True
                                            matched_file = file_name
                                            break
                                    if contains_tracked_file:
                                        break
                            except (OSError, PermissionError, UnicodeError):
                                continue

                            if contains_tracked_file:
                                break

                    if contains_tracked_file:
                        # 为包含我们文件的未知软件创建临时分组
                        software_type = f'unknown_editor_{hash(window_title_lower) % 1000}'

                # 如果确定了软件类型，加入分组
                if software_type:
                    if software_type not in software_groups:
                        software_groups[software_type] = {
                            'windows': [],
                            'display_name': self._get_software_display_name(software_type, detected_pattern, matched_file),
                            'primary_window': None
                        }

                    software_groups[software_type]['windows'].append(window)

                    # 选择主窗口（通常是第一个或者标题最简洁的）
                    if (software_groups[software_type]['primary_window'] is None or
                        len(window_title) < len(software_groups[software_type]['primary_window'].title)):
                        software_groups[software_type]['primary_window'] = window

            # 构建最终的程序列表（每个软件只保留一个主窗口）
            running_programs = []
            for software_type, group_info in software_groups.items():
                if group_info['primary_window']:
                    running_programs.append({
                        'name': group_info['display_name'],
                        'window': group_info['primary_window'],
                        'title': group_info['primary_window'].title,
                        'software_type': software_type,
                        'window_count': len(group_info['windows'])
                    })

            # 记录检测结果
            if running_programs:
                self.log_info("程序检测完成", f"检测到 {len(running_programs)} 个不同的软件需要关闭")
                for prog in running_programs:
                    self.log_info("检测到软件", f"软件: {prog['name']} | 窗口数: {prog['window_count']} | 主窗口: {prog['title'][:50]}...")
            else:
                self.log_info("程序检测完成", "未检测到需要关闭的软件")

            if excluded_titles:
                self.log_info("排除的窗口", f"已排除 {len(excluded_titles)} 个不应关闭的窗口")

            return running_programs

        except Exception as e:
            self.log_error("获取运行程序失败", f"错误: {str(e)}")
            self.update_status(f"获取运行程序失败: {str(e)}", "orange")
            return []

    def _get_software_display_name(self, software_type, detected_pattern, matched_file):
        """获取软件的显示名称"""
        display_names = {
            'word': 'Microsoft Word',
            'excel': 'Microsoft Excel',
            'powerpoint': 'Microsoft PowerPoint',
            'wps_writer': 'WPS 文字',
            'wps_spreadsheet': 'WPS 表格',
            'wps_presentation': 'WPS 演示',
            'notepad': '记事本',
            'notepadpp': 'Notepad++',
            'sublime': 'Sublime Text',
            'vscode': 'Visual Studio Code',
            'pycharm': 'PyCharm',
            'adobe_acrobat': 'Adobe Acrobat',
            'foxit': 'Foxit Reader',
            'chrome': 'Google Chrome',
            'edge': 'Microsoft Edge'
        }

        if software_type in display_names:
            return display_names[software_type]
        elif software_type.startswith('unknown_editor_'):
            if matched_file:
                return f'文档编辑器 ({matched_file})'
            else:
                return '未知文档编辑器'
        else:
            return detected_pattern.title() if detected_pattern else software_type

# 确保导入winreg（仅在Windows上可用）
try:
    import winreg
except ImportError:
    winreg = None

if __name__ == "__main__":
    try:
        # 初始化Tkinter根窗口
        try:
            root = Tk()
            print("Tkinter root window created successfully")
        except Exception as e:
            print(f"Failed to create Tkinter root window: {e}")
            sys.exit(1)

        # 初始化应用程序
        try:
            app = ActivityTracker(root)
            # 应用程序初始化完成后，记录启动信息到主日志系统
            if hasattr(app, 'log_info') and app.logger:
                app.log_info("程序启动", "ActivityTracker 应用程序初始化成功")
            print("ActivityTracker application initialized successfully")
        except Exception as e:
            print(f"Failed to initialize ActivityTracker application: {e}")
            import traceback
            print(f"Traceback: {traceback.format_exc()}")
            try:
                root.destroy()
            except:
                pass
            sys.exit(1)

        # 启动主事件循环
        try:
            # 记录主循环启动到日志
            if hasattr(app, 'log_info') and app.logger:
                app.log_info("程序运行", "主事件循环启动")
            print("Starting main event loop")

            root.mainloop()

            # 记录主循环正常结束
            if hasattr(app, 'log_info') and app.logger:
                app.log_info("程序结束", "主事件循环正常结束")
            print("Main event loop ended normally")
        except Exception as e:
            error_msg = f"Main event loop crashed: {e}"
            # 记录崩溃信息到日志
            if hasattr(app, 'log_error') and app.logger:
                app.log_error("程序崩溃", f"主事件循环崩溃: {e}")
            print(error_msg)

            import traceback
            traceback_msg = traceback.format_exc()
            if hasattr(app, 'log_error') and app.logger:
                app.log_error("程序崩溃详情", traceback_msg)
            print(f"Traceback: {traceback_msg}")
            sys.exit(1)

    except Exception as e:
        # 最后的异常处理
        error_msg = f"FATAL ERROR in main: {e}"
        # 尝试记录到主日志系统（如果可用）
        try:
            if 'app' in locals() and hasattr(app, 'log_error') and app.logger:
                app.log_error("程序致命错误", f"主程序发生致命错误: {e}")
            print(error_msg)
        except:
            print(error_msg)

        import traceback
        traceback_msg = traceback.format_exc()
        try:
            if 'app' in locals() and hasattr(app, 'log_error') and app.logger:
                app.log_error("程序致命错误详情", traceback_msg)
            print(f"Traceback: {traceback_msg}")
        except:
            print(f"Traceback: {traceback_msg}")

        sys.exit(1)
