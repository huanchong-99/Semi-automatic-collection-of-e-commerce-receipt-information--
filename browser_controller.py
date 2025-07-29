#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
浏览器控制相关

从原始main.py文件中拆分出来的模块
"""

from utils import *
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service


class BrowserController:
    """浏览器控制相关"""
    
    def __init__(self):
        """初始化浏览器控制器"""
        pass

    def _start_browser(self):
        """启动浏览器并通过远程调试端口连接"""
        # 检查Selenium依赖
        if webdriver is None:
            self._log_info("Selenium模块未正确导入，请安装: pip install selenium", "red")
            messagebox.showerror("依赖缺失", "Selenium模块未正确导入，请安装: pip install selenium")
            return
            
        self._log_info("正在启动浏览器...", "blue")
        self.connection_progress["value"] = 25

        def connect_browser():
            try:
                BROWSER_PATH = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
                DEBUG_PORT = str(self.debug_port)
                USER_DATA_DIR = self.user_data_dir
                if not os.path.exists(USER_DATA_DIR):
                    os.makedirs(USER_DATA_DIR)
                cmd = [BROWSER_PATH, f'--remote-debugging-port={DEBUG_PORT}', f'--user-data-dir={USER_DATA_DIR}']
                cmd_str = ' '.join(cmd)
                self._log_info(f"正在启动Edge浏览器: {cmd_str}", "blue")
                subprocess.Popen(cmd, creationflags=0x08000000)
                self._log_info('等待浏览器启动...')
                time.sleep(2)
                
                max_retries = 15
                driver = None
                for i in range(max_retries):
                    self._log_info(f'尝试连接到浏览器 (第 {i + 1}/{max_retries} 次)...')
                    try:
                        options = Options()
                        options.add_experimental_option('debuggerAddress', f'localhost:{DEBUG_PORT}')
                        
                        # 使用Service类指定驱动路径（相对路径）
                        driver_path = "msedgedriver.exe"
                        
                        if os.path.exists(driver_path):
                            self._log_info(f"使用本地驱动: {driver_path}", "blue")
                            service = Service(executable_path=driver_path)
                            driver = webdriver.Edge(service=service, options=options)
                        else:
                            self._log_info("本地驱动不存在，使用系统默认驱动", "orange")
                            driver = webdriver.Edge(options=options)
                        
                        if driver.window_handles:
                            driver.switch_to.window(driver.window_handles[-1])
                            self.driver = driver
                            break
                    except Exception as e:
                        self._log_info(f"连接失败: {e}", "orange")
                        driver = None
                    time.sleep(1)
                    
                if not driver:
                    raise ConnectionError('无法连接到浏览器。')
                    
                self._log_info('成功连接到Edge浏览器。', "green")
                self.is_browser_connected = True
                self.browser_status.config(text="已连接", foreground="green")
                self.root.after(0, lambda: self.start_button.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.configure_button.config(state=tk.NORMAL))
                # 启用智能循环与滚动设置区域的所有按钮
                self.root.after(0, lambda: self.collect_ref1_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.collect_ref2_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.reload_ref_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.collect_page_btn.config(state=tk.NORMAL))
                self.root.after(0, lambda: self.reload_page_btn.config(state=tk.NORMAL))
                self._inject_hover_listener()
            except Exception as e:
                self.driver = None
                self._log_info(f'连接浏览器失败: {e}', "red")
                self.browser_status.config(text="连接失败", foreground="red")
                self.connection_progress["value"] = 0

        threading.Thread(target=connect_browser, daemon=True).start()

    def _inject_hover_listener(self):
        """递归注入悬停监听脚本到所有frame，便于采集鼠标悬停元素，对齐代码逻辑.md"""
        if not self.driver:
            return
        self._log_info('准备向所有框架（包括嵌套框架）注入悬停监听脚本...')
        try:
            self.driver.switch_to.default_content()
            self._inject_listener_recursive()
            self._log_info('悬停监听脚本注入完成。')
        except Exception as e:
            self._log_info(f'注入悬停监听脚本时发生主错误: {e}', 'red')
        finally:
            if self.driver:
                self.driver.switch_to.default_content()

    def _inject_listener_recursive(self):
        if not self.driver:
            self._log_info('无法注入监听脚本：浏览器未连接', 'red')
            return
        setup_script = '''
        if (window.pddToolListenerInjected) { return; }
        window.pddToolListenerInjected = true;
        window.lastHoveredElement = null;
        document.addEventListener('mouseover', function(e) {
            window.lastHoveredElement = e.target;
        }, true);
        '''
        self.driver.execute_script(setup_script)
        iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
        for i in range(len(iframes)):
            try:
                self.driver.switch_to.frame(i)
                self._inject_listener_recursive()
            except Exception as e:
                self._log_info(f'递归注入时无法切换到某个iframe: {e}', 'orange')
            finally:
                self.driver.switch_to.parent_frame()
    
    def _close_browser(self):
        """关闭浏览器连接和进程"""
        if self.driver:
            self.driver.quit()
            self.driver = None
        if hasattr(self, 'ws') and self.ws:
            self.ws.close()
        if hasattr(self, 'browser_process') and self.browser_process:
            self.browser_process.terminate()
        self.is_browser_connected = False
        self.browser_status.config(text="未连接", foreground="red")
        # 禁用智能循环与滚动设置区域的所有按钮
        if hasattr(self, 'collect_ref1_btn'):
            self.collect_ref1_btn.config(state=tk.DISABLED)
        if hasattr(self, 'collect_ref2_btn'):
            self.collect_ref2_btn.config(state=tk.DISABLED)
        if hasattr(self, 'reload_ref_btn'):
            self.reload_ref_btn.config(state=tk.DISABLED)
        if hasattr(self, 'collect_page_btn'):
            self.collect_page_btn.config(state=tk.DISABLED)
        if hasattr(self, 'reload_page_btn'):
            self.reload_page_btn.config(state=tk.DISABLED)
        self._log_info("浏览器已关闭", "blue")

    def _switch_focus_to_browser(self):
        """
        主动将焦点切换到浏览器窗口
        """
        if not self.driver:
            self._log_info("浏览器未连接，无法切换焦点", "orange")
            return False
        
        try:
            # 尝试执行一个简单的JavaScript来激活浏览器窗口
            self.driver.execute_script("window.focus()")
            
            # 获取浏览器窗口标题
            title = self.driver.title
            
            # 尝试通过标题查找窗口并激活
            if title:
                hwnd = ctypes.windll.user32.FindWindowW(None, title)
                if hwnd:
                    ctypes.windll.user32.SetForegroundWindow(hwnd)
                    self._log_info(f"已将焦点切换到浏览器: {title}", "blue")
                    return True
                
            self._log_info("无法通过标题切换焦点到浏览器", "orange")
            return False
        except Exception as e:
            self._log_info(f"切换焦点到浏览器失败: {str(e)}", "red")
            return False
            

    def _switch_focus_between_browser_and_tool(self, action):
        """
        在浏览器和采集工具之间切换焦点
        
        参数:
        - action: 要执行的操作，如"click"、"getText"等
        """
        if action in ["click", "clickAndGetClipboard"]:
            # 点击前确保浏览器有焦点
            self._switch_focus_to_browser()
            time.sleep(0.2)
        else:
            # 其他操作确保采集工具有焦点
            self._manage_focus()
            time.sleep(0.2)
    

    def _manage_focus(self):
        """
        焦点管理器，确保在关键操作后焦点回到采集工具
        """
        try:
            # 获取主窗口句柄
            hwnd = ctypes.windll.user32.FindWindowW(None, "收货信息自动采集工具")
            if hwnd:
                # 尝试将焦点设置回采集工具窗口
                ctypes.windll.user32.SetForegroundWindow(hwnd)
                self._log_info("已将焦点返回到采集工具", "blue")
                return True
            else:
                self._log_info("未找到采集工具窗口，无法重置焦点", "orange")
                return False
        except Exception as e:
            self._log_info(f"焦点管理失败: {str(e)}", "red")
            return False
    

    def _ensure_focus_for_clipboard(self, max_attempts=3, delay=0.5):
        """
        确保在获取剪贴板内容前采集工具有焦点
        
        参数:
        - max_attempts: 最大尝试次数
        - delay: 每次尝试后的延时(秒)
        
        返回:
        - 是否成功将焦点设置到采集工具
        """
        for attempt in range(max_attempts):
            # 先尝试恢复焦点
            self._manage_focus()
            
            # 等待焦点切换完成
            time.sleep(delay)
            
            try:
                # 检查当前活动窗口是否为采集工具
                foreground_window = ctypes.windll.user32.GetForegroundWindow()
                window_text_length = ctypes.windll.user32.GetWindowTextLengthW(foreground_window)
                buffer = ctypes.create_unicode_buffer(window_text_length + 1)
                ctypes.windll.user32.GetWindowTextW(foreground_window, buffer, window_text_length + 1)
                
                if "收货信息自动采集工具" in buffer.value:
                    self._log_info(f"已确认焦点在采集工具上 (尝试 {attempt+1}/{max_attempts})", "green")
                    return True
                    
                self._log_info(f"焦点仍在其他窗口: {buffer.value} (尝试 {attempt+1}/{max_attempts})", "orange")
            except Exception as e:
                self._log_info(f"检查焦点失败: {str(e)} (尝试 {attempt+1}/{max_attempts})", "red")
        
        self._log_info("无法确保焦点回到采集工具，将尝试继续操作", "red")
        return False
    
