#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
UI组件和界面相关

从原始main.py文件中拆分出来的模块
"""

import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from utils import logger
import logging
import os

class UIComponents:
    """UI组件和界面相关"""

    def __init__(self, root):
        """初始化UI组件和界面相关"""
        self.root = root
        self._create_gui()

    def _log_info(self, message, color=None):
        print(f"LOG: {message}")  # 强制输出到控制台，便于调试
        logger.info(message)
        # 添加到UI文本框
        self.status_text.config(state=tk.NORMAL)  # 临时设置为可写
        self.status_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {message}\n")
        # 应用颜色标签
        if color:
            last_line_start = self.status_text.index(f"end-2l")
            self.status_text.tag_add(f"color_{color}", last_line_start, "end-1c")
        self.status_text.see(tk.END)  # 滚动到最新内容
        self.status_text.config(state=tk.DISABLED)  # 恢复只读
    

    def _create_gui(self):
        """创建GUI界面"""
        # 设置窗口标题和大小
        self.root.title("收货信息自动采集工具")
        self.root.geometry("850x600")
        
        # 创建主框架并使用网格布局
        self.main_frame = ttk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # ====== 顶部区域：浏览器连接状态与模式切换 ======
        self.top_frame = ttk.Frame(self.main_frame)
        self.top_frame.pack(fill=tk.X, pady=(0, 5))
        
        # 浏览器连接按钮和状态
        self.browser_frame = ttk.Frame(self.top_frame)
        self.browser_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.browser_button = ttk.Button(
            self.browser_frame, 
            text="打开浏览器", 
            command=self._start_browser
        )
        self.browser_button.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Label(self.browser_frame, text="浏览器状态:").pack(side=tk.LEFT)
        self.browser_status = ttk.Label(
            self.browser_frame, 
            text="未连接",
            foreground="red"
        )
        self.browser_status.pack(side=tk.LEFT, padx=5)
        
        # 浏览器连接进度条
        self.connection_progress = ttk.Progressbar(
            self.browser_frame, 
            orient=tk.HORIZONTAL, 
            length=200, 
            mode='determinate'
        )
        self.connection_progress.pack(side=tk.LEFT, padx=10)
        
        # 模式切换下拉框
        self.mode_frame = ttk.Frame(self.top_frame)
        self.mode_frame.pack(side=tk.RIGHT)
        
        ttk.Label(self.mode_frame, text="操作模式:").pack(side=tk.LEFT)
        self.mode_combobox = ttk.Combobox(
            self.mode_frame, 
            textvariable=self.collection_mode,
            values=["正常模式", "采集模式"],
            state="readonly",
            width=10
        )
        self.mode_combobox.pack(side=tk.LEFT, padx=5)
        self.mode_combobox.bind("<<ComboboxSelected>>", self._mode_changed)
        
        # ====== 中间区域：控制按钮和状态显示 ======
        self.middle_frame = ttk.Frame(self.main_frame)
        self.middle_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # 控制按钮区
        self.control_frame = ttk.Frame(self.middle_frame)
        self.control_frame.pack(fill=tk.X)
        
        self.start_button = ttk.Button(
            self.control_frame, 
            text="开始", 
            command=self._start_collection,
            state=tk.DISABLED  # 初始禁用，直到浏览器连接
        )
        self.start_button.pack(side=tk.LEFT, padx=5)
        
        self.pause_button = ttk.Button(
            self.control_frame, 
            text="暂停", 
            command=self._pause_collection,
            state=tk.DISABLED
        )
        self.pause_button.pack(side=tk.LEFT, padx=5)
        
        self.continue_button = ttk.Button(
            self.control_frame, 
            text="继续", 
            command=self._continue_collection,
            state=tk.DISABLED
        )
        self.continue_button.pack(side=tk.LEFT, padx=5)
        
        self.stop_button = ttk.Button(
            self.control_frame, 
            text="终止", 
            command=self._stop_collection,
            state=tk.DISABLED
        )
        self.stop_button.pack(side=tk.LEFT, padx=5)
        
        # 添加配置操作按钮
        self.configure_button = ttk.Button(
            self.control_frame, 
            text="配置操作", 
            command=self._configure_operations,
            state=tk.DISABLED  # 初始禁用，直到浏览器连接
        )
        self.configure_button.pack(side=tk.LEFT, padx=5)
        
        # 添加验证码选择目标窗口按钮
        self.captcha_window_button = ttk.Button(
            self.control_frame,
            text="选择目标窗口",
            command=self._select_captcha_target_window
        )
        self.captcha_window_button.pack(side=tk.LEFT, padx=5)
        
        # 添加验证码管理按钮
        self.captcha_manage_button = ttk.Button(
            self.control_frame,
            text="验证码配置",
            command=self._show_captcha_manager
        )
        self.captcha_manage_button.pack(side=tk.LEFT, padx=5)
        
        # 移除重复的采集翻页元素按钮，该按钮已在智能循环与滚动设置区域中
        
        # 翻页页数选择下拉框
        ttk.Label(self.control_frame, text="翻页页数:").pack(side=tk.LEFT, padx=(10, 0))
        self.page_count_var = tk.StringVar(value="20")
        self.page_count_combo = ttk.Combobox(
            self.control_frame,
            textvariable=self.page_count_var,
            values=["5", "10", "20", "50"],
            state="readonly",
            width=5
        )
        self.page_count_combo.pack(side=tk.LEFT, padx=5)
        self.page_count_combo.current(2)  # 默认选择20
        
        # 绑定翻页页数变化事件
        self.page_count_var.trace('w', self.on_page_count_changed)
        
        # 进度显示区
        self.progress_frame = ttk.Frame(self.middle_frame)
        self.progress_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(self.progress_frame, text="进度:").pack(side=tk.LEFT)
        self.progress_label = ttk.Label(self.progress_frame, text="0/0")
        self.progress_label.pack(side=tk.LEFT, padx=5)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            orient=tk.HORIZONTAL, 
            length=400, 
            mode='determinate'
        )
        self.progress_bar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)
        
        # 状态文本显示区（带颜色）
        self.status_frame = ttk.LabelFrame(self.middle_frame, text="状态信息")
        self.status_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        from tkinter.scrolledtext import ScrolledText
        self.status_text = ScrolledText(
            self.status_frame, 
            wrap=tk.WORD,
            height=10
        )
        self.status_text.pack(fill=tk.BOTH, expand=True)
        self.status_text.config(state=tk.DISABLED)  # 初始设置为只读
        
        # 配置状态文本的颜色标签
        self.status_text.tag_config("color_red", foreground="red")
        self.status_text.tag_config("color_green", foreground="green")
        self.status_text.tag_config("color_blue", foreground="blue")
        self.status_text.tag_config("color_orange", foreground="orange")
        
        # ====== 统一功能区域：整合系统状态、数据导出、智能循环设置 ======
        self.unified_function_frame = ttk.LabelFrame(self.middle_frame, text="功能控制面板")
        self.unified_function_frame.pack(fill=tk.X, pady=5)
        
        # 创建主容器
        main_container = ttk.Frame(self.unified_function_frame)
        main_container.pack(fill=tk.X, padx=8, pady=3)
        
        # === 第一部分：左右两个LabelFrame并排布局 ===
        top_section = ttk.Frame(main_container)
        top_section.pack(fill=tk.X, pady=(0, 4))
        
        # 左侧：默认就行，别乱动（可折叠）
        default_section = ttk.Frame(top_section)
        default_section.pack(side=tk.LEFT, padx=(0, 15))
        
        # 右侧：主要功能区域LabelFrame
        main_function_frame = ttk.LabelFrame(top_section, text="功能区域")
        main_function_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 主要功能区域内容容器
        main_function_content = ttk.Frame(main_function_frame)
        main_function_content.pack(fill=tk.BOTH, expand=True, padx=8, pady=3)
        
        # 第一行：系统状态和数据导出
        first_row = ttk.Frame(main_function_content)
        first_row.pack(fill=tk.X, pady=(0, 4))
        
        # 系统状态显示
        status_section = ttk.Frame(first_row)
        status_section.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 状态标题
        status_title = ttk.Label(status_section, text="系统状态", font=("Arial", 10, "bold"))
        status_title.pack(anchor=tk.W)
        
        # 第一行：验证码检测状态
        captcha_row = ttk.Frame(status_section)
        captcha_row.pack(fill=tk.X, pady=(2, 0))
        
        self.captcha_status_label = ttk.Label(
            captcha_row, 
            text="验证码检测: 未启动", 
            font=("Arial", 9)
        )
        self.captcha_status_label.pack(side=tk.LEFT)
        
        self.captcha_status_indicator = tk.Canvas(
            captcha_row, 
            width=14, 
            height=14, 
            bg=self.root.cget('bg'), 
            highlightthickness=0
        )
        self.captcha_status_indicator.pack(side=tk.LEFT, padx=(8, 15))
        self.captcha_indicator_circle = self.captcha_status_indicator.create_oval(2, 2, 12, 12, fill="gray", outline="")
        
        self.target_window_label = ttk.Label(
            captcha_row,
            text="目标窗口: 未设置",
            font=("Arial", 9)
        )
        self.target_window_label.pack(side=tk.LEFT, padx=(0, 15))
        
        self.template_count_label = ttk.Label(
            captcha_row,
            text="模板数量: 0",
            font=("Arial", 9)
        )
        self.template_count_label.pack(side=tk.LEFT)
        
        # 第二行：模块化翻页状态
        paging_row = ttk.Frame(status_section)
        paging_row.pack(fill=tk.X, pady=(2, 0))
        
        self.modular_paging_status_label = ttk.Label(
            paging_row,
            text="模块化翻页: 已启用",
            font=("Arial", 9),
            foreground="green"
        )
        self.modular_paging_status_label.pack(side=tk.LEFT)
        
        self.paging_status_indicator = tk.Canvas(
            paging_row, 
            width=14, 
            height=14, 
            bg=self.root.cget('bg'), 
            highlightthickness=0
        )
        self.paging_status_indicator.pack(side=tk.LEFT, padx=(8, 15))
        self.paging_indicator_circle = self.paging_status_indicator.create_oval(2, 2, 12, 12, fill="green", outline="")
        
        self.modular_paging_info_label = ttk.Label(
            paging_row,
            text="每页作为独立模块处理，解决跨页元素定位问题",
            font=("Arial", 9),
            foreground="gray"
        )
        self.modular_paging_info_label.pack(side=tk.LEFT)
        
        # === 左侧：默认就行，别乱动（可折叠界面）===
        self.default_labelframe = ttk.LabelFrame(default_section, text="默认就行，别乱动 ▼", padding=5)
        self.default_labelframe.pack(fill=tk.BOTH, expand=True)
        
        # 创建内容框架（可以隐藏/显示）
        self.default_content = ttk.Frame(self.default_labelframe)
        self.default_content.pack(fill=tk.BOTH, expand=True)
        
        # 第一行：操作间隔设置
        default_row1 = ttk.Frame(self.default_content)
        default_row1.pack(fill=tk.X, pady=(0, 3))
        
        ttk.Label(default_row1, text="操作间隔(秒):", font=("Arial", 9)).pack(side=tk.LEFT)
        self.interval_var = tk.StringVar(value=str(self.auto_action_interval))
        self.interval_entry = ttk.Spinbox(
            default_row1, 
            from_=0.1, 
            to=10.0, 
            increment=0.1, 
            width=5, 
            textvariable=self.interval_var,
            command=self._update_interval
        )
        self.interval_entry.pack(side=tk.LEFT, padx=(3, 0))
        self.interval_entry.bind('<Return>', self._update_interval)
        self.interval_entry.bind('<FocusOut>', self._update_interval)
        
        # 第二行：元素偏移量管理
        default_row2 = ttk.Frame(self.default_content)
        default_row2.pack(fill=tk.X, pady=(0, 3))
        
        ttk.Label(default_row2, text="元素偏移量", font=("Arial", 9)).pack(side=tk.LEFT)
        self.offset_manage_button = ttk.Button(
            default_row2,
            text="管理偏移量",
            command=self._show_offset_manager
        )
        self.offset_manage_button.pack(side=tk.LEFT, padx=(3, 0))
        
        # 第三行：剪贴板管理
        default_row3 = ttk.Frame(self.default_content)
        default_row3.pack(fill=tk.X, pady=(0, 3))
        
        ttk.Button(
            default_row3,
            text="保存剪贴板映射",
            command=self._save_clipboard_mappings
        ).pack(fill=tk.X, pady=(0, 2))
        
        # 第四行：手动关联
        default_row4 = ttk.Frame(self.default_content)
        default_row4.pack(fill=tk.X)
        
        ttk.Button(
            default_row4,
            text="手动关联订单ID",
            command=lambda: self._manually_associate_clipboard_with_order_id()
        ).pack(fill=tk.X)
        
        # 添加折叠/展开功能
        self.default_expanded = True
        self.default_labelframe.bind('<Button-1>', self._toggle_default_section)
        
        # 数据导出功能（移动到第一行右侧，竖直排列）
        export_section = ttk.Frame(first_row)
        export_section.pack(side=tk.RIGHT, padx=(10, 65))
        
        # 导出标题
        export_title = ttk.Label(export_section, text="数据导出", font=("Arial", 10, "bold"))
        export_title.pack(anchor=tk.W)
        
        # 正常模式下使用的按钮（竖直排列）
        self.excel_button = ttk.Button(
            export_section,
            text="导出Excel",
            command=self._export_excel,
            state=tk.DISABLED,  # 初始禁用，直到有数据
            width=12
        )
        self.excel_button.pack(pady=2)
        
        self.word_button = ttk.Button(
            export_section,
            text="导出Word",
            command=self._export_word,
            state=tk.DISABLED,  # 初始禁用，直到有数据
            width=12
        )
        self.word_button.pack(pady=2)
        
        # 采集模式下使用的按钮
        self.json_button = ttk.Button(
            export_section,
            text="导出JSON",
            command=self._export_json,
            state=tk.DISABLED,  # 初始禁用
            width=8
        )
        # 初始不显示，根据模式切换显示
        
        # 导出按钮只保留导出相关功能
        # 剪贴板管理功能已移动到"默认就行，别乱动"分类
        
        # ====== 底部区域：设置与导出 ======
        self.bottom_frame = ttk.Frame(self.main_frame)
        self.bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        # 设置区
        self.settings_frame = ttk.Frame(self.top_frame)
        self.settings_frame.pack(side=tk.RIGHT, padx=5)
        
        self.always_on_top_check = ttk.Checkbutton(
            self.settings_frame,
            text="窗口置顶",
            variable=self.always_on_top,
            command=self._update_always_on_top
        )
        self.always_on_top_check.pack(side=tk.LEFT, padx=10)
        
        self.confirm_click_check = ttk.Checkbutton(
            self.settings_frame,
            text="点击前确认",
            variable=self.confirm_click
        )
        self.confirm_click_check.pack(side=tk.LEFT, padx=10)
        
        # 操作间隔和偏移量管理功能已移动到"默认就行，别乱动"分类

        # 第二行：智能循环与滚动设置
        smart_loop_section = ttk.Frame(main_function_content)
        smart_loop_section.pack(fill=tk.X, pady=(0, 4))
        
        # 智能循环标题
        smart_loop_title = ttk.Label(smart_loop_section, text="智能循环与滚动设置", font=("Arial", 10, "bold"))
        smart_loop_title.pack(anchor=tk.W)
        
        # 初始化相关变量
        self.ref1_xpath = None
        self.ref2_xpath = None
        self.scroll_container_xpath = None  # 兼容性保留，但不再采集
        self.scroll_step = None
        self.next_page_xpath = None  # 翻页元素XPath
        self.next_page_collected = False  # 翻页元素采集标志
        
        # 第一行：采集按钮和重新加载按钮
        first_row = ttk.Frame(smart_loop_section)
        first_row.pack(fill=tk.X, pady=(2, 0))
        
        self.collect_ref1_btn = ttk.Button(first_row, text="采集第1个订单参照点", command=self._collect_ref1_xpath, state=tk.DISABLED)
        self.collect_ref1_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        self.collect_ref2_btn = ttk.Button(first_row, text="采集第2个订单参照点", command=self._collect_ref2_xpath, state=tk.DISABLED)
        self.collect_ref2_btn.pack(side=tk.LEFT, padx=5)
        
        self.reload_ref_btn = ttk.Button(first_row, text="重新加载参照点", command=self._reload_reference_points, state=tk.DISABLED)
        self.reload_ref_btn.pack(side=tk.LEFT, padx=5)
        
        self.collect_page_btn = ttk.Button(first_row, text="采集翻页元素", command=self.collect_page_turn_element, state=tk.DISABLED)
        self.collect_page_btn.pack(side=tk.LEFT, padx=5)
        
        self.reload_page_btn = ttk.Button(first_row, text="重新加载翻页元素", command=self._reload_page_turn_element, state=tk.DISABLED)
        self.reload_page_btn.pack(side=tk.LEFT, padx=5)
        
        # 第二行：状态显示
        second_row = ttk.Frame(smart_loop_section)
        second_row.pack(fill=tk.X, pady=(2, 0))
        
        self.ref1_label = ttk.Label(second_row, text="第1个订单参照点: 未设置", font=("Arial", 9))
        self.ref1_label.pack(side=tk.LEFT, padx=(0, 15))
        self.ref2_label = ttk.Label(second_row, text="第2个订单参照点: 未设置", font=("Arial", 9))
        self.ref2_label.pack(side=tk.LEFT, padx=(0, 15))
        self.page_turn_label = ttk.Label(second_row, text="翻页元素: 未设置", font=("Arial", 9))
        self.page_turn_label.pack(side=tk.LEFT)
        
        # 模块化翻页开关（隐藏但默认启用）
        # 注意：模块化翻页功能默认启用，不再显示UI控件
        self.use_modular_paging = tk.BooleanVar(value=True)  # 默认启用
        # 不再创建UI控件，但保留变量以确保功能正常工作
        
        # 自动调用一次回调函数以初始化状态
        self._on_modular_paging_changed()

    def _toggle_default_section(self, event=None):
        """切换默认设置区域的折叠/展开状态"""
        if self.default_expanded:
            # 折叠：隐藏内容
            self.default_content.pack_forget()
            self.default_labelframe.config(text="默认就行，别乱动 ▶")
            self.default_expanded = False
        else:
            # 展开：显示内容
            self.default_content.pack(fill=tk.BOTH, expand=True)
            self.default_labelframe.config(text="默认就行，别乱动 ▼")
            self.default_expanded = True
    
    def _update_always_on_top(self):
        """更新窗口置顶状态"""
        if self.always_on_top.get():
            self.root.attributes('-topmost', True)
        else:
            self.root.attributes('-topmost', False)
            
    def _update_interval(self, event=None):
        """更新操作间隔设置"""
        try:
            value = float(self.interval_var.get())
            if value < 0.1:
                value = 0.1
            elif value > 10.0:
                value = 10.0
            self.auto_action_interval = value
            self.interval_var.set(str(value))
            self._log_info(f"已更新操作间隔为 {value} 秒", "blue")
        except ValueError:
            self.interval_var.set(str(self.auto_action_interval))
            self._log_info("操作间隔必须是有效的数字", "red")
    
    def _mode_changed(self, event=None):
        """处理模式切换"""
        mode = self.collection_mode.get()
        self._log_info(f"已切换到{mode}", "blue")
        
        # 根据不同模式更新UI和行为
        if mode == "正常模式":
            # 正常模式下的UI调整
            self.start_button.config(text="开始")
            
            # 显示Excel和Word导出按钮，隐藏JSON导出按钮
            self.excel_button.pack(side=tk.LEFT, padx=5, pady=5)
            self.word_button.pack(side=tk.LEFT, padx=5, pady=5)
            self.json_button.pack_forget()
            
        else:  # 采集模式
            # 采集模式下的UI调整
            self.start_button.config(text="开始采集")
            
            # 隐藏Excel和Word导出按钮，显示JSON导出按钮
            self.excel_button.pack_forget()
            self.word_button.pack_forget()
            self.json_button.pack(side=tk.LEFT, padx=5, pady=5)
    
    def _handle_key_event(self, event):
        """处理键盘事件"""
        key = event.char
        keysym = event.keysym
        if hasattr(self, '_pending_collect') and self._pending_collect:
            if not self.driver:
                self._log_info("浏览器未连接，无法采集", "red")
                messagebox.showerror("错误", "浏览器未连接，无法采集")
                self._pending_collect = None
                return
            try:
                self._inject_hover_listener()
                self.driver.switch_to.default_content()
                xpath = self._get_hovered_xpath_recursive()
                if not xpath:
                    self._log_info("未检测到悬停元素，请确保鼠标已悬停在目标元素上，且页面未刷新。", "orange")
                    messagebox.showerror("采集失败", "未能获取到元素信息，请确保鼠标悬停在浏览器页面内的有效元素上。")
                    self._pending_collect = None
                    return
                if self._pending_collect == 'ref1':
                    self.ref1_xpath = xpath
                    self._log_info(f"已采集第1个订单参照点XPath: {xpath}", "green")
                    messagebox.showinfo("采集成功", "第1个订单参照点采集成功！")
                elif self._pending_collect == 'ref2':
                    self.ref2_xpath = xpath
                    self._log_info(f"已采集第2个订单参照点XPath: {xpath}", "green")
                    messagebox.showinfo("采集成功", "第2个订单参照点采集成功！")
                
                # 更新状态显示并保存配置
                self._update_element_status_display()
                self._save_element_config()
            except Exception as e:
                self._log_info(f"采集参照点异常: {str(e)}", "red")
                import traceback
                self._log_info(traceback.format_exc(), "red")
                messagebox.showerror("采集失败", f"采集过程中发生异常：{str(e)}")
            finally:
                self._pending_collect = None
            return
            
        # 处理.键：采集模式下采集，正常模式下暂停，翻页采集模式下采集翻页元素
        if key == '.':
            # 如果正在采集翻页元素
            if hasattr(self, 'collecting_page_turn') and self.collecting_page_turn:
                self._log_info("检测到.键，开始采集翻页元素...", "blue")
                self._handle_page_turn_collection()
                return
            elif self.collection_mode.get() == "采集模式":
                if self.is_running and not self.is_paused:
                    self._log_info("检测到.键，开始采集元素...", "blue")
                    self._collect_element()
                    return
            else:  # 正常模式
                if self.is_running and not self.is_paused:
                    self._log_info("检测到.键，暂停操作...", "orange")
                    self._pause_collection()
                    return
        
        # 处理-键：继续操作
        if key == '-':
            if self.is_running and self.is_paused:
                self._log_info("检测到-键，继续操作...", "green")
                self._continue_collection()
                return
        
        # 处理*键：终止操作并立即启用数据导出
        if key == '*':
            if self.is_running:
                self._log_info("检测到*键，终止操作...", "red")
                self._stop_collection()
                # 立即启用数据导出功能
                if hasattr(self, 'collected_data') and len(self.collected_data) > 0:
                    self._log_info("操作已终止，数据导出功能已启用", "green")
                    messagebox.showinfo("操作终止", "操作已终止，现在可以进行数据导出。")
                return
    
    def _load_element_config(self):
        """加载元素配置文件"""
        try:
            import json
            config_path = "element_config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 加载参照点
                ref_points = config.get('reference_points', {})
                self.ref1_xpath = ref_points.get('ref1_xpath')
                self.ref2_xpath = ref_points.get('ref2_xpath')
                
                # 加载翻页元素
                page_turn = config.get('page_turn', {})
                self.next_page_xpath = page_turn.get('next_page_xpath')
                # 如果存在翻页元素XPath，设置采集标志为True
                if self.next_page_xpath:
                    self.next_page_collected = True
                else:
                    self.next_page_collected = False
                
                # 更新UI显示
                self._update_element_status_display()
                self._log_info("已自动加载元素配置", "green")
            else:
                self._log_info("元素配置文件不存在，使用默认设置", "orange")
        except Exception as e:
            self._log_info(f"加载元素配置失败: {e}", "red")
    
    def _save_element_config(self):
        """保存元素配置文件"""
        try:
            import json
            config = {
                "reference_points": {
                    "ref1_xpath": self.ref1_xpath,
                    "ref2_xpath": self.ref2_xpath
                },
                "page_turn": {
                    "next_page_xpath": self.next_page_xpath
                }
            }
            
            config_path = "element_config.json"
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            self._log_info("已保存元素配置", "green")
        except Exception as e:
            self._log_info(f"保存元素配置失败: {e}", "red")
    
    def _update_element_status_display(self):
        """更新元素状态显示"""
        # 更新参照点状态
        if self.ref1_xpath:
            self.ref1_label.config(text=f"第1个订单参照点: 已设置")
        else:
            self.ref1_label.config(text="第1个订单参照点: 未设置")
        
        if self.ref2_xpath:
            self.ref2_label.config(text=f"第2个订单参照点: 已设置")
        else:
            self.ref2_label.config(text="第2个订单参照点: 未设置")
        
        # 更新翻页元素状态
        if self.next_page_xpath:
            self.page_turn_label.config(text="翻页元素: 已设置")
        else:
            self.page_turn_label.config(text="翻页元素: 未设置")
    
    def _reload_reference_points(self):
        """重新加载参照点配置"""
        try:
            import json
            config_path = "element_config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                ref_points = config.get('reference_points', {})
                self.ref1_xpath = ref_points.get('ref1_xpath')
                self.ref2_xpath = ref_points.get('ref2_xpath')
                
                self._update_element_status_display()
                self._log_info("已重新加载参照点配置", "green")
            else:
                self._log_info("配置文件不存在", "orange")
        except Exception as e:
            self._log_info(f"重新加载参照点配置失败: {e}", "red")
    
    def _reload_page_turn_element(self):
        """重新加载翻页元素配置"""
        try:
            import json
            config_path = "element_config.json"
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                page_turn = config.get('page_turn', {})
                self.next_page_xpath = page_turn.get('next_page_xpath')
                # 如果存在翻页元素XPath，设置采集标志为True
                if self.next_page_xpath:
                    self.next_page_collected = True
                    self._log_info("已重新加载翻页元素配置", "green")
                else:
                    self.next_page_collected = False
                    self._log_info("翻页元素配置为空", "orange")
                
                self._update_element_status_display()
            else:
                self._log_info("配置文件不存在", "orange")
        except Exception as e:
            self._log_info(f"重新加载翻页元素配置失败: {e}", "red")
    
    def _update_captcha_status_display(self):
        """更新验证码检测状态显示"""
        if hasattr(self, 'captcha_running') and self.captcha_running:
            if hasattr(self, 'captcha_detected') and self.captcha_detected:
                # 检测到验证码 - 红色
                self.captcha_status_label.config(text="验证码检测: 检测到验证码")
                self.captcha_status_indicator.itemconfig(self.captcha_indicator_circle, fill="red")
            else:
                # 未检测到验证码 - 绿色
                self.captcha_status_label.config(text="验证码检测: 运行中")
                self.captcha_status_indicator.itemconfig(self.captcha_indicator_circle, fill="green")
        else:
            # 未启动 - 灰色
            self.captcha_status_label.config(text="验证码检测: 未启动")
            self.captcha_status_indicator.itemconfig(self.captcha_indicator_circle, fill="gray")
        
        # 更新目标窗口显示
        if hasattr(self, 'target_window_title') and self.target_window_title:
            self.target_window_label.config(text=f"目标窗口: {self.target_window_title}")
        else:
            self.target_window_label.config(text="目标窗口: 未设置")
        
        # 更新模板数量显示
        if hasattr(self, 'template_images'):
            template_count = len(self.template_images)
            self.template_count_label.config(text=f"模板数量: {template_count}")
        else:
            self.template_count_label.config(text="模板数量: 0")
    
    def _on_modular_paging_changed(self):
        """模块化翻页开关变化回调"""
        try:
            # 同步到data_processor的use_modular_paging属性
            if hasattr(self, 'use_modular_paging'):
                use_modular = self.use_modular_paging.get()
                self.use_modular_paging_value = use_modular  # 保存到实例属性
                
                # 更新UI状态显示
                if hasattr(self, 'modular_paging_status_label') and hasattr(self, 'paging_status_indicator'):
                    if use_modular:
                        self.modular_paging_status_label.config(
                            text="模块化翻页: 已启用",
                            foreground="green"
                        )
                        self.paging_status_indicator.itemconfig(self.paging_indicator_circle, fill="green")
                    else:
                        self.modular_paging_status_label.config(
                            text="模块化翻页: 已禁用",
                            foreground="red"
                        )
                        self.paging_status_indicator.itemconfig(self.paging_indicator_circle, fill="red")
                
                if use_modular:
                    self._log_info("已启用模块化翻页模式", "green")
                    self._log_info("模块化翻页将每页作为独立模块处理，解决跨页元素定位问题", "blue")
                else:
                    self._log_info("已禁用模块化翻页模式，使用传统翻页方式", "orange")
                    
        except Exception as e:
            self._log_info(f"模块化翻页开关设置失败: {e}", "red")
    
    def _shrink_window_to_corner(self):
        """将窗口缩小到左上角，然后移动到右上角"""
        try:
            # 设置缩小后的窗口大小和位置
            compact_width = 388  # X轴再缩小2个像素点
            compact_height = 450  # 增加高度以显示完整的系统状态
            x_position = 10  # 距离左边缘10像素
            y_position = 10  # 距离顶部10像素
            
            # 应用新的窗口几何设置（先移动到左上角）
            self.root.geometry(f"{compact_width}x{compact_height}+{x_position}+{y_position}")
            
            # 设置窗口置顶
            self.root.attributes('-topmost', True)
            
            self._log_info(f"窗口已缩小到左上角 ({compact_width}x{compact_height})", "blue")
            
            # 延时后移动到右上角
            self.root.after(1000, self._move_to_top_right)
            
        except Exception as e:
            self._log_info(f"窗口缩小失败: {str(e)}", "red")
            import traceback
            self.logger.error(f"窗口缩小异常: {traceback.format_exc()}")
    
    def _move_to_top_right(self):
        """将窗口移动到屏幕右上角"""
        try:
            # 获取屏幕宽度
            screen_width = self.root.winfo_screenwidth()
            
            # 计算右上角位置
            compact_width = 388  # X轴再缩小2个像素点
            compact_height = 450
            x_position = screen_width - compact_width - 10  # 距离右边缘10像素
            y_position = 10  # 距离顶部10像素
            
            # 移动到右上角
            self.root.geometry(f"{compact_width}x{compact_height}+{x_position}+{y_position}")
            
            self._log_info("窗口已移动到右上角", "green")
            
        except Exception as e:
            self._log_info(f"窗口移动失败: {str(e)}", "red")
            import traceback
            self.logger.error(f"窗口移动异常: {traceback.format_exc()}")
    
