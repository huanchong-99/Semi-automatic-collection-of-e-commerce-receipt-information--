#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验证码检测相关

从原始main.py文件中拆分出来的模块
"""

from utils import *
import glob

class CaptchaDetector:
    """验证码检测相关"""
    
    def __init__(self):
        pass

    def _show_captcha_manager(self):
        """显示验证码管理窗口"""
        manager_window = tk.Toplevel(self.root)
        manager_window.title("验证码检测管理")
        manager_window.geometry("600x500")
        manager_window.transient(self.root)
        manager_window.grab_set()
        
        # 模板管理区域
        template_frame = ttk.LabelFrame(manager_window, text="模板管理")
        template_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 模板列表
        ttk.Label(template_frame, text="当前模板:").pack(anchor="w", pady=(10, 5))
        self.template_listbox = tk.Listbox(template_frame, height=5)
        self.template_listbox.pack(fill="x", padx=10, pady=5)
        
        # 更新模板列表
        self._update_template_list()
        
        # 模板操作按钮
        template_btn_frame = ttk.Frame(template_frame)
        template_btn_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Button(template_btn_frame, text="添加模板", command=self._add_captcha_template).pack(side="left", padx=5)
        ttk.Button(template_btn_frame, text="清除所有模板", command=self._clear_captcha_templates).pack(side="left", padx=5)
        ttk.Button(template_btn_frame, text="自动加载模板", command=self._auto_load_captcha_templates).pack(side="left", padx=5)
        ttk.Button(template_btn_frame, text="截图保存为模板", command=self._capture_window_as_template).pack(side="left", padx=5)
        
        # 高级设置区域
        settings_frame = ttk.LabelFrame(manager_window, text="高级设置")
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        # 检测间隔设置
        interval_frame = ttk.Frame(settings_frame)
        interval_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(interval_frame, text="检测间隔(秒):").pack(side="left")
        self.detection_interval_var = tk.DoubleVar(value=self.detection_interval)
        interval_spinbox = ttk.Spinbox(interval_frame, from_=0.1, to=5.0, increment=0.1, 
                                     textvariable=self.detection_interval_var, width=10)
        interval_spinbox.pack(side="left", padx=5)
        
        # 相似度阈值设置
        similarity_frame = ttk.Frame(settings_frame)
        similarity_frame.pack(fill="x", padx=10, pady=5)
        ttk.Label(similarity_frame, text="相似度阈值:").pack(side="left")
        self.similarity_threshold_var = tk.DoubleVar(value=self.similarity_threshold)
        similarity_scale = ttk.Scale(similarity_frame, from_=0.1, to=1.0, 
                                   variable=self.similarity_threshold_var, orient="horizontal")
        similarity_scale.pack(side="left", fill="x", expand=True, padx=5)
        self.similarity_label = ttk.Label(similarity_frame, text=f"{self.similarity_threshold:.2f}")
        self.similarity_label.pack(side="left")
        
        def update_similarity_label(*args):
            self.similarity_label.config(text=f"{self.similarity_threshold_var.get():.2f}")
        
        self.similarity_threshold_var.trace('w', update_similarity_label)
        
        # 遮罩层检测设置
        mask_frame = ttk.Frame(settings_frame)
        mask_frame.pack(fill="x", padx=10, pady=5)
        self.use_mask_detection_var = tk.BooleanVar(value=self.use_mask_detection)
        ttk.Checkbutton(mask_frame, text="启用遮罩层检测", variable=self.use_mask_detection_var).pack(side="left")
        
        # 控制按钮
        control_frame = ttk.Frame(manager_window)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        def apply_settings():
            self.detection_interval = self.detection_interval_var.get()
            self.similarity_threshold = self.similarity_threshold_var.get()
            self.use_mask_detection = self.use_mask_detection_var.get()
            self._log_info("验证码检测设置已更新", "green")
            manager_window.destroy()
        
        ttk.Button(control_frame, text="应用设置", command=apply_settings).pack(side="right", padx=5)
        ttk.Button(control_frame, text="取消", command=manager_window.destroy).pack(side="right", padx=5)
    

    def _update_template_list(self):
        """更新模板列表显示"""
        if hasattr(self, 'template_listbox'):
            self.template_listbox.delete(0, tk.END)
            for i, template in enumerate(self.template_images):
                self.template_listbox.insert(tk.END, f"模板 {i+1} ({template['name']})")
    

    def _select_captcha_target_window(self):
        """选择验证码检测的目标窗口"""
        try:
            import win32gui
            
            def enum_windows_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
                    windows.append((hwnd, win32gui.GetWindowText(hwnd)))
                return True
            
            windows = []
            win32gui.EnumWindows(enum_windows_callback, windows)
            
            # 创建窗口选择对话框
            select_window = tk.Toplevel(self.root)
            select_window.title("选择目标窗口")
            select_window.geometry("400x300")
            select_window.transient(self.root)
            select_window.grab_set()
            
            ttk.Label(select_window, text="请选择要监控的窗口:").pack(pady=10)
            
            window_listbox = tk.Listbox(select_window)
            window_listbox.pack(fill="both", expand=True, padx=10, pady=5)
            
            # 配置高亮颜色
            window_listbox.configure(selectbackground="#0078d4", selectforeground="white")
            
            # 添加窗口项并高亮包含Microsoft Edge的项
            edge_indices = []  # 记录包含Microsoft Edge的项的索引
            for i, (hwnd, title) in enumerate(windows):
                display_text = f"{title} (句柄: {hwnd})"
                window_listbox.insert(tk.END, display_text)
                
                # 检查是否包含Microsoft Edge
                if "Microsoft Edge" in title or "Microsoft​ Edge" in title:
                    edge_indices.append(i)
            
            # 为包含Microsoft Edge的项设置特殊背景色
            for index in edge_indices:
                window_listbox.itemconfig(index, {'bg': '#ffeb3b', 'fg': '#000000'})  # 黄色背景，黑色文字
            
            # 如果找到Microsoft Edge窗口，自动选中第一个
            if edge_indices:
                window_listbox.selection_set(edge_indices[0])  # 选中第一个Edge窗口
                window_listbox.activate(edge_indices[0])  # 激活该项
                window_listbox.see(edge_indices[0])  # 确保该项可见
            
            def on_select():
                selection = window_listbox.curselection()
                if selection:
                    selected_window = windows[selection[0]]
                    self.target_window_handle = selected_window[0]
                    self.target_window_title = selected_window[1]
                    self._update_captcha_status_display()
                    self._log_info(f"已选择目标窗口: {self.target_window_title}", "green")
                    select_window.destroy()
            
            ttk.Button(select_window, text="确定", command=on_select).pack(pady=10)
            
        except Exception as e:
            self._log_info(f"选择目标窗口失败: {e}", "error")
    

    def _add_captcha_template(self):
        """添加验证码模板"""
        try:
            from tkinter import filedialog
            file_path = filedialog.askopenfilename(
                title="选择验证码模板图片",
                filetypes=[("图片文件", "*.jpg *.jpeg *.png *.bmp")]
            )
            
            if file_path:
                # 使用cv2.imdecode处理中文路径
                try:
                    with open(file_path, 'rb') as f:
                        image_data = f.read()
                    image_array = np.frombuffer(image_data, np.uint8)
                    template_image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
                except Exception as e:
                    template_image = None
                    self._log_info(f"读取文件失败: {e}", "error")
                
                if template_image is not None:
                    template_name = os.path.basename(file_path)
                    self.template_images.append({
                        'name': template_name,
                        'image': template_image,
                        'path': file_path
                    })
                    self._update_template_list()
                    self._update_captcha_status_display()
                    self._log_info(f"已添加模板: {template_name}", "green")
                else:
                    self._log_info("无法读取图片文件", "error")
        except Exception as e:
            self._log_info(f"添加模板失败: {e}", "error")
    

    def _clear_captcha_templates(self):
        """清除所有验证码模板"""
        self.template_images.clear()
        self._update_template_list()
        self._update_captcha_status_display()
        self._log_info("已清除所有模板", "orange")
    

    def _auto_load_captcha_templates(self):
        """自动加载验证码模板"""
        try:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            templates_dir = os.path.join(current_dir, "captcha_templates")
            
            # 确保目录存在
            if not os.path.exists(templates_dir):
                self._log_info(f"模板目录不存在: {templates_dir}", "orange")
                return
            
            # 使用os.listdir遍历文件，避免glob的编码问题
            template_files = []
            try:
                for filename in os.listdir(templates_dir):
                    if filename.startswith("模板") and filename.endswith(".jpg"):
                        full_path = os.path.join(templates_dir, filename)
                        template_files.append(full_path)
            except UnicodeDecodeError as e:
                self._log_info(f"读取模板目录时遇到编码问题: {e}", "error")
                # 尝试使用不同的编码
                try:
                    for filename in os.listdir(templates_dir.encode('utf-8').decode('utf-8')):
                        if "模板" in filename and filename.endswith(".jpg"):
                            full_path = os.path.join(templates_dir, filename)
                            template_files.append(full_path)
                except Exception as e2:
                    self._log_info(f"使用UTF-8编码也失败: {e2}", "error")
                    return
            
            loaded_count = 0
            for file_path in template_files:
                # 使用cv2.imdecode处理中文路径
                try:
                    with open(file_path, 'rb') as f:
                        image_data = f.read()
                    image_array = np.frombuffer(image_data, np.uint8)
                    template_image = cv2.imdecode(image_array, cv2.IMREAD_COLOR)
                except Exception as e:
                    template_image = None
                    self._log_info(f"读取模板文件失败 {os.path.basename(file_path)}: {e}", "error")
                
                if template_image is not None:
                    template_name = os.path.basename(file_path)
                    # 检查是否已存在同名模板
                    if not any(t['name'] == template_name for t in self.template_images):
                        self.template_images.append({
                            'name': template_name,
                            'image': template_image,
                            'path': file_path
                        })
                        loaded_count += 1
            
            self._update_template_list()
            self._update_captcha_status_display()
            
            if loaded_count > 0:
                self._log_info(f"自动加载了 {loaded_count} 个模板", "green")
            else:
                self._log_info("未找到新的模板文件", "orange")
                
        except Exception as e:
            self._log_info(f"自动加载模板失败: {e}", "error")
    

    def _capture_window_as_template(self):
        """截图目标窗口并保存为模板"""
        try:
            if not self.target_window_handle:
                self._log_info("请先选择目标窗口", "error")
                return
            
            # 截取目标窗口
            screenshot = self._capture_target_window()
            if screenshot is None:
                self._log_info("截取窗口失败", "error")
                return
            
            # 转换为PIL图像
            screenshot_pil = Image.fromarray(cv2.cvtColor(screenshot, cv2.COLOR_BGR2RGB))
            
            # 获取程序所在目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            templates_dir = os.path.join(current_dir, "captcha_templates")
            
            # 确保captcha_templates文件夹存在
            if not os.path.exists(templates_dir):
                os.makedirs(templates_dir)
            
            # 查找现有模板文件，确定下一个编号
            template_pattern = os.path.join(templates_dir, "模板*.jpg")
            existing_templates = glob.glob(template_pattern)
            
            # 确定新模板的编号
            if not existing_templates:
                template_name = "模板.jpg"
            else:
                # 提取现有模板的编号
                numbers = []
                for template_path in existing_templates:
                    filename = os.path.basename(template_path)
                    if filename == "模板.jpg":
                        numbers.append(1)
                    elif filename.startswith("模板") and filename.endswith(".jpg"):
                        try:
                            num_str = filename[2:-4]  # 去掉"模板"和".jpg"
                            if num_str.isdigit():
                                numbers.append(int(num_str))
                        except:
                            pass
                
                # 找到下一个可用编号
                if numbers:
                    next_num = max(numbers) + 1
                    if next_num == 2:
                        template_name = "模板2.jpg"
                    else:
                        template_name = f"模板{next_num}.jpg"
                else:
                    template_name = "模板.jpg"
            
            # 保存模板文件
            template_path = os.path.join(templates_dir, template_name)
            screenshot_pil.save(template_path, "JPEG", quality=95)
            
            # 加载到模板列表中
            template_image = cv2.cvtColor(screenshot, cv2.COLOR_BGR2RGB)
            template_image = cv2.cvtColor(template_image, cv2.COLOR_RGB2BGR)
            
            self.template_images.append({
                'name': template_name,
                'image': template_image,
                'path': template_path
            })
            
            self._update_template_list()
            self._update_captcha_status_display()
            self._log_info(f"已截图并保存为模板: {template_name}", "green")
            
        except Exception as e:
            self._log_info(f"截图保存为模板失败: {e}", "error")
    

    def _start_captcha_detection(self):
        """启动验证码检测"""
        if self.captcha_running:
            return
        
        if not self.template_images and not self.use_mask_detection:
            self._log_info("请先添加模板或启用遮罩层检测", "error")
            return
        
        if not self.target_window_handle:
            self._log_info("请先选择目标窗口", "error")
            return
        
        self.captcha_running = True
        self.captcha_detected = False
        self._update_captcha_status_display()
        
        # 启动检测线程
        self.captcha_thread = threading.Thread(target=self._captcha_detection_loop, daemon=True)
        self.captcha_thread.start()
        
        self._log_info("验证码检测已启动", "green")
    

    def _stop_captcha_detection(self):
        """停止验证码检测"""
        self.captcha_running = False
        self.captcha_detected = False
        self._update_captcha_status_display()
        self._log_info("验证码检测已停止", "orange")
    

    def _captcha_detection_loop(self):
        """验证码检测循环"""
        frame_count = 0  # 帧计数器，用于跳过初始几帧
        while self.captcha_running:
            try:
                # 截取目标窗口
                screenshot = self._capture_target_window()
                if screenshot is None:
                    time.sleep(self.detection_interval)
                    continue
                
                # 跳过初始几帧以避免误报
                if frame_count < 3:  # 跳过前3帧
                    frame_count += 1
                    time.sleep(self.detection_interval)
                    continue
                
                detected = False
                
                # 模板匹配检测
                if self.template_images:
                    detected = self._template_match_detection(screenshot)
                
                # 遮罩层检测
                if not detected and self.use_mask_detection:
                    detected = self._mask_layer_detection(screenshot)
                
                # 处理检测结果
                if detected and not self.captcha_detected:
                    self.captcha_detected = True
                    self.root.after(0, self._on_captcha_detected)
                elif not detected and self.captcha_detected:
                    self.captcha_detected = False
                    self.root.after(0, self._on_captcha_disappeared)
                
                # 更新状态显示
                self.root.after(0, self._update_captcha_status_display)
                
            except Exception as e:
                self._log_info(f"验证码检测出错: {e}", "error")
            
            time.sleep(self.detection_interval)
    

    def _capture_target_window(self):
        """截取目标窗口"""
        try:
            import win32gui, win32ui, win32con
            from PIL import Image
            
            # 检查窗口是否存在
            if not win32gui.IsWindow(self.target_window_handle):
                return None
            
            # 获取窗口客户区矩形（不包括标题栏和边框）
            client_rect = win32gui.GetClientRect(self.target_window_handle)
            client_width = client_rect[2]
            client_height = client_rect[3]
            
            if client_width <= 0 or client_height <= 0:
                return None
            
            # 获取窗口设备上下文
            hwnd_dc = win32gui.GetWindowDC(self.target_window_handle)
            mfc_dc = win32ui.CreateDCFromHandle(hwnd_dc)
            save_dc = mfc_dc.CreateCompatibleDC()
            
            # 创建位图对象
            save_bitmap = win32ui.CreateBitmap()
            save_bitmap.CreateCompatibleBitmap(mfc_dc, client_width, client_height)
            save_dc.SelectObject(save_bitmap)
            
            # 复制窗口内容到位图
            # 使用PrintWindow API，这对于某些应用程序更可靠
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            result = user32.PrintWindow(self.target_window_handle, save_dc.GetSafeHdc(), 3)  # PW_RENDERFULLCONTENT = 3
            
            if result == 0:
                # PrintWindow失败，尝试BitBlt方法
                result = save_dc.BitBlt((0, 0), (client_width, client_height), mfc_dc, (0, 0), win32con.SRCCOPY)
            
            if result == 0:
                # 两种方法都失败
                return None
            
            # 获取位图数据
            bmp_info = save_bitmap.GetInfo()
            bmp_str = save_bitmap.GetBitmapBits(True)
            
            # 转换为PIL图像
            im = Image.frombuffer(
                'RGB',
                (bmp_info['bmWidth'], bmp_info['bmHeight']),
                bmp_str, 'raw', 'BGRX', 0, 1
            )
            
            # 转换为OpenCV格式
            screenshot = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
            
            # 清理资源
            win32gui.DeleteObject(save_bitmap.GetHandle())
            save_dc.DeleteDC()
            mfc_dc.DeleteDC()
            win32gui.ReleaseDC(self.target_window_handle, hwnd_dc)
            
            return screenshot
            
        except Exception as e:
            self._log_info(f"截取窗口失败: {e}", "error")
            return None
    

    def _template_match_detection(self, screenshot):
        """模板匹配检测"""
        try:
            for template_info in self.template_images:
                template = template_info['image']
                result = cv2.matchTemplate(screenshot, template, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, _ = cv2.minMaxLoc(result)
                
                if max_val >= self.similarity_threshold:
                    return True
            return False
        except Exception as e:
            self._log_info(f"模板匹配检测出错: {e}", "error")
            return False
    

    def _mask_layer_detection(self, screenshot):
        """遮罩层检测"""
        try:
            # 转换为灰度图
            gray = cv2.cvtColor(screenshot, cv2.COLOR_BGR2GRAY)
            
            # 计算图像的平均亮度
            mean_brightness = np.mean(gray)
            
            # 如果平均亮度过低，可能存在遮罩层
            if mean_brightness < 50:  # 可调阈值
                return True
            
            # 检测大面积的暗色区域
            _, binary = cv2.threshold(gray, 30, 255, cv2.THRESH_BINARY)
            dark_pixels = np.sum(binary == 0)
            total_pixels = gray.shape[0] * gray.shape[1]
            dark_ratio = dark_pixels / total_pixels
            
            # 如果暗色像素比例过高，可能存在遮罩层
            if dark_ratio > 0.7:  # 可调阈值
                return True
            
            return False
        except Exception as e:
            self._log_info(f"遮罩层检测出错: {e}", "error")
            return False
    

    # _update_captcha_status_display 方法已移至 UIComponents 类中
    # 避免方法重复定义导致的调用混乱
    

    def _on_captcha_detected(self):
        """检测到验证码时的处理"""
        try:
            # 停止当前操作
            if self.is_running and not self.is_paused:
                self.is_paused = True
                self.force_stop_flag = True
                self._log_info("检测到验证码，自动暂停操作", "red")
            
            # 播放警报音
            try:
                winsound.Beep(1000, 500)  # 1000Hz, 500ms
            except:
                pass
            
            # 显示警告框
            from tkinter import messagebox
            messagebox.showwarning("验证码检测", "检测到验证码！\n\n操作已自动暂停，请手动处理验证码后继续。")
            
            self._update_captcha_status_display()
            
        except Exception as e:
            self._log_info(f"处理验证码检测事件出错: {e}", "error")
    

    def _on_captcha_disappeared(self):
        """验证码消失时的处理 - 阶段3增强：集成重试管理器"""
        try:
            # 清除强制停止标志和暂停状态
            if hasattr(self, 'force_stop_flag'):
                self.force_stop_flag = False
            if hasattr(self, 'is_paused'):
                self.is_paused = False
            
            self._log_info("验证码已消失，恢复操作执行", "green")
            self._update_captcha_status_display()
            
            # 阶段3增强：使用重试管理器记录验证码消失事件
            if hasattr(self, 'retry_manager') and self.retry_manager.is_retry_logging_enabled():
                from utils import create_retry_log_entry, write_retry_log
                log_entry = create_retry_log_entry("captcha_disappeared", "system", {
                    "timestamp": datetime.now().isoformat(),
                    "current_order_index": getattr(self, 'current_order_index', None),
                    "retry_triggered": True
                })
                write_retry_log(log_entry)
            
            # 设置重试标志 - 新增功能
            self.retry_current_order = True
            
            # 阶段3增强：重置重试计数器（新的验证码处理周期）
            if hasattr(self, 'retry_manager') and hasattr(self, 'current_order_index'):
                self.retry_manager.reset_retry_attempts("order_processing", self.current_order_index)
                self._log_info(f"已重置订单 {self.current_order_index} 的重试计数器", "blue")
            
            self.root.after(0, self._trigger_order_retry_enhanced)
            
            # 如果之前因验证码暂停，现在可以准备重新执行当前订单
            if hasattr(self, 'current_order_index'):
                self._log_info("验证码已处理，准备重试当前订单", "green")
            
        except Exception as e:
            self._log_info(f"处理验证码消失事件出错: {e}", "error")
    
    def _trigger_order_retry(self):
        """触发订单重试 - 新增方法"""
        try:
            if hasattr(self, 'retry_current_order') and self.retry_current_order:
                self._log_info("触发当前订单重试机制", "blue")
                # 重试标志将在 data_processor.py 中被检查和处理
        except Exception as e:
            self._log_info(f"触发订单重试失败: {e}", "error")
    
    def _trigger_order_retry_enhanced(self):
        """触发订单重试 - 阶段3增强方法"""
        try:
            if hasattr(self, 'retry_current_order') and self.retry_current_order:
                self._log_info("触发当前订单重试机制", "blue")
                
                # 阶段3增强：使用重试管理器获取重试配置
                if hasattr(self, 'retry_manager'):
                    retry_delay = self.retry_manager.get_retry_delay()
                    retry_strategies = self.retry_manager.get_retry_strategies()
                    
                    self._log_info(f"重试配置 - 延迟: {retry_delay}秒, 策略: {', '.join(retry_strategies)}", "cyan")
                    
                    # 记录重试触发事件
                    if self.retry_manager.is_retry_logging_enabled():
                        from utils import create_retry_log_entry, write_retry_log
                        log_entry = create_retry_log_entry("retry_triggered", "captcha_recovery", {
                            "retry_delay": retry_delay,
                            "retry_strategies": retry_strategies,
                            "current_order_index": getattr(self, 'current_order_index', None)
                        })
                        write_retry_log(log_entry)
                
                # 重试标志将在 data_processor.py 中被检查和处理
        except Exception as e:
            self._log_info(f"触发增强订单重试失败: {e}", "error")
    
