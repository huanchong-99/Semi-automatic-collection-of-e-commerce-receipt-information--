#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
主应用程序入口

整合所有模块的主应用程序
"""

from utils import *
from ui_components import UIComponents
from browser_controller import BrowserController
from element_collector import ElementCollector
from data_processor import DataProcessor
from clipboard_manager import ClipboardManager
from captcha_detector import CaptchaDetector
from page_turner import PageTurner
from config_manager import ConfigManager
from retry_manager import RetryManager

class ShippingInfoCollector(
    UIComponents,
    BrowserController,
    ElementCollector,
    DataProcessor,
    ClipboardManager,
    CaptchaDetector,
    PageTurner,
    ConfigManager
):
    """收货信息采集工具主类"""
    
    def __init__(self, root):
        """初始化应用程序"""
        self.root = root
        
        # 初始化基本属性
        self._init_basic_attributes()
        
        # 按正确顺序初始化所有组件
        # 首先初始化UI组件（这会创建界面）
        UIComponents.__init__(self, root)
        
        # 然后初始化其他功能模块
        BrowserController.__init__(self)
        ElementCollector.__init__(self)
        DataProcessor.__init__(self)
        ClipboardManager.__init__(self)
        CaptchaDetector.__init__(self)
        PageTurner.__init__(self)
        ConfigManager.__init__(self)
        
        # 初始化重试管理器 - 新增功能
        self.retry_manager = RetryManager()
        
        # 设置窗口置顶状态
        self._update_always_on_top()
        
        # 绑定键盘事件
        self.root.bind("<Key>", self._handle_key_event)
        
        # 加载偏移量配置（按元素名称保存的WASD微调配置）
        self._load_offset_config()
        
        # 同时清空映射文件，确保不会重新加载旧数据
        self._save_clipboard_mappings()
        self._log_info("已清空所有历史映射数据和文件，准备收集全新信息", "green")
        
        # 启动时记录日志
        self._log_info("程序已启动，等待操作...", "blue")
        self._log_info(f"订单ID与收货信息映射字典初始化完成，包含 {len(self.order_clipboard_contents)} 个映射", "blue")
        
        # 程序启动时自动加载验证码模板
        self._auto_load_captcha_templates()
        
        # 程序启动时自动加载元素配置
        self._load_element_config()
        
        # 更新验证码状态显示
        self._update_captcha_status_display = UIComponents._update_captcha_status_display.__get__(self)
        self._update_captcha_status_display()

    def _init_basic_attributes(self):
        """初始化基本属性"""
        # 基本运行状态
        self.auto_action_interval = 1.0
        self.is_running = False
        self.is_paused = False
        self.force_stop_flag = False
        
        # 浏览器相关
        self.is_browser_connected = False
        self.browser_process = None
        self.ws = None
        self.session_id = None
        self.request_id = 1
        self.debug_port = 9222
        self.user_data_dir = os.path.join(tempfile.gettempdir(), "edge_user_data")
        self.driver = None
        
        # 数据收集相关
        self.collected_data = []
        self.operation_sequence = []
        self.element_offsets = {}
        self.current_order_id = None
        self.last_order_ids = []
        self.consecutive_same_order = 0
        self.scroll_distance_multiplier = 1.0
        
        # 采集模式相关
        self.collection_mode = tk.StringVar(value="正常模式")
        self.is_distance_learning = False
        self.distance_learning_step = 0
        self.first_element_position = None
        self.second_element_position = None
        self.scroll_distance = None
        
        # 剪贴板相关
        self.last_clipboard_content = ""
        self.order_clipboard_contents = {}
        self.orders_need_review = set()
        self.content_validation_results = {}
        self.last_captured_order_id = None
        self.clipboard_monitor_active = False
        self.last_known_clipboard = ""
        self.clipboard_monitor_thread = None
        
        # 重试机制相关 - 新增功能
        self.retry_current_order = False
        self.current_order_index = None
        self.retry_count = 0
        self.max_retry_attempts = 3
        
        # 验证码检测相关
        self.captcha_running = False
        self.captcha_detected = False
        self.last_detection_time = 0
        self.captcha_detection_thread = None
        self.template_images = []
        self.template_paths = []
        self.monitor_area = None
        self.detection_interval = 0.25
        self.similarity_threshold = 0.7
        self.consecutive_frames = 3
        self.consecutive_detections = 0
        self.consecutive_non_detections = 0
        self.use_mask_detection = True
        self.mask_threshold = 0.1
        self.temp_dir = tempfile.mkdtemp()
        self.screenshot_files = []
        self.target_window_handle = None
        self.target_window_title = ""
        self.use_window_capture = True
        self.captcha_force_stop = False
        
        # 翻页相关
        self.next_page_xpath = None
        self.next_page_collected = False
        self.page_turn_count = 0
        self.target_page_count = 20
        self.screenshot_dir = "page_screenshots"
        self.collecting_page_turn = False
        
        # UI相关
        self.always_on_top = tk.BooleanVar(value=True)
        self.confirm_click = tk.BooleanVar(value=False)
        
        # 智能循环相关
        self.ref1_xpath = None
        self.ref2_xpath = None
        self.scroll_container_xpath = None
        self.scroll_step = None
        
        # 创建截图目录
        if not os.path.exists(self.screenshot_dir):
            os.makedirs(self.screenshot_dir)
            
        # 程序启动时自动加载验证码模板（必须在template_images初始化之后）
        # 这个调用会在CaptchaDetector.__init__()中执行
        
        # 每次启动时清空所有历史映射数据，确保收集全新信息
        self.order_clipboard_contents = {}
        
        # 启动时记录日志
        print(f"DEBUG-INIT: 初始化订单ID与收货信息映射字典，包含 {len(self.order_clipboard_contents)} 个映射")

def main():
    """主函数"""
    # 检查依赖是否完整
    if not dependencies_ok:
        print("\n程序无法启动，请先安装缺失的依赖包。")
        input("按回车键退出...")
        return
    
    # 设置高DPI感知（Windows系统）
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception as e:
        logging.warning(f"无法设置DPI感知: {str(e)}")
    
    # 创建Tk根窗口
    try:
        root = tk.Tk()
    except Exception as e:
        logging.error(f"Tk初始化失败: {str(e)}")
        print(f"GUI初始化失败: {str(e)}")
        print("请确认已正确安装Python和Tkinter，或尝试重新安装Python。")
        print(f"Python版本: {sys.version}")
        print(f"TCL_LIBRARY环境变量: {os.environ.get('TCL_LIBRARY', '未设置')}")
        print(f"TK_LIBRARY环境变量: {os.environ.get('TK_LIBRARY', '未设置')}")
        input("按回车键退出...")
        return  # 退出程序
    
    # 创建应用实例
    app = ShippingInfoCollector(root)
    
    # 启动应用主循环
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # 捕获并记录未处理的异常
        error_msg = f"程序发生未处理的异常: {str(e)}\n{traceback.format_exc()}"
        logging.error(error_msg)
        try:
            messagebox.showerror("错误", f"程序发生错误:\n{str(e)}\n\n详细信息已记录到日志文件中。")
        except Exception:
            # 如果messagebox也失败，至少确保错误被记录
            print(f"程序发生严重错误: {str(e)}\n详细信息已记录到日志文件中。")