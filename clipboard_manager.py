#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
剪贴板管理相关

从原始main.py文件中拆分出来的模块
"""

from utils import *

class ClipboardManager:
    """剪贴板管理相关"""
    
    def __init__(self):
        """初始化剪贴板管理器"""
        pass

    def _get_clipboard_content(self):
        """
        获取剪贴板内容，添加焦点保护
        """
        # 检查pyperclip依赖
        if pyperclip is None:
            self._log_info("pyperclip模块未正确导入，请安装: pip install pyperclip", "red")
            return ""
            
        try:
            # 尝试将焦点返回到采集工具窗口
            self._ensure_focus_for_clipboard()
            
            # 增加短暂延时，确保剪贴板内容已更新
            time.sleep(0.3)
            
            # 获取剪贴板内容
            clipboard_content = pyperclip.paste()
            content_length = len(clipboard_content) if clipboard_content else 0
            
            self._log_info(f"获取剪贴板内容，长度: {content_length}", "blue")
            return clipboard_content
        except Exception as e:
            self._log_info(f"获取剪贴板内容失败: {str(e)}", "red")
            return ""
    

    def _wait_for_clipboard_content(self, timeout=10.0, check_interval=0.2, min_length=10):
        """
        等待剪贴板内容更新，带超时机制，增加初始内容过滤
        
        参数:
        - timeout: 最大等待时间（秒）
        - check_interval: 检查间隔（秒）
        - min_length: 有效内容最小长度
        
        返回:
        - 更新的剪贴板内容或None
        """
        # 检查pyperclip依赖
        if pyperclip is None:
            self._log_info("pyperclip模块未正确导入，请安装: pip install pyperclip", "red")
            return None
            
        start_time = time.time()
        last_content = pyperclip.paste()
        initial_length = len(last_content) if last_content else 0
        
        # 获取初始剪贴板状态进行比较
        initial_content = getattr(self, 'initial_clipboard_content', "")
        
        self._log_info(f"开始等待剪贴板更新，初始内容长度: {initial_length}", "blue")
        
        # 清空剪贴板
        pyperclip.copy("")
        time.sleep(0.2)
        
        while time.time() - start_time < timeout:
            # 检查是否已终止操作
            if not self.is_running:
                self._log_info("操作已终止，停止等待剪贴板更新", "orange")
                return None
                
            # 确保焦点回到采集工具
            if (time.time() - start_time) > 2.0:  # 2秒后开始尝试恢复焦点
                self._manage_focus()
                
            current_content = pyperclip.paste()
            current_length = len(current_content) if current_content else 0
            
            # 检查内容是否等于初始剪贴板内容
            is_initial_content = current_content == initial_content
            if is_initial_content and current_length > 0:
                self._log_info(f"检测到剪贴板内容与脚本启动前相同，继续等待...", "orange")
                time.sleep(check_interval)
                continue
                
            # 检查内容长度是否超过1000字符，如果是则很可能是错误内容
            if current_length > 1000:
                self._log_info(f"检测到异常长度的剪贴板内容({current_length}字符)，继续等待...", "red")
                time.sleep(check_interval)
                continue
                
            # 内容发生变化且达到最小长度要求且不等于初始内容
            if current_content != last_content and current_length >= min_length and not is_initial_content:
                self._log_info(f"检测到有效剪贴板内容更新，长度: {current_length}", "green")
                return current_content
                
            last_content = current_content
            time.sleep(check_interval)
        
        self._log_info(f"等待剪贴板更新超时，最终内容长度: {len(last_content) if last_content else 0}", "orange")
        
        # 确保不返回初始内容
        if last_content == initial_content:
            self._log_info("最终内容与脚本启动前相同，返回空内容", "red")
            return None
            
        return last_content if len(last_content) >= min_length else None
    

    def _store_clipboard_content(self, element_name, content, order_id=None):
        if not content or not isinstance(content, str) or not content.strip():
            self._log_info(f"剪贴板内容为空，未存储", "orange")
            return False
        # 必须有明确订单ID，禁止 None、空字符串、temp_id 写入
        if not order_id or not str(order_id).strip() or str(order_id).startswith("temp_order"):
            self._log_info("未提供有效订单ID，拒绝写入映射", "red")
            return False
        clean_order_id = order_id.replace("订单编号：", "") if isinstance(order_id, str) else str(order_id)
        if clean_order_id in self.order_clipboard_contents:
            existing_content = self.order_clipboard_contents[clean_order_id]
            if len(existing_content) > len(content):
                self._log_info(f"保留现有更长的收货信息映射 (现有:{len(existing_content)}字符 vs 新内容:{len(content)}字符)", "orange")
                print(f"DEBUG-STORE-SKIP: 订单ID '{clean_order_id}' 已有更长的收货信息 ({len(existing_content)} > {len(content)})")
                return False
        self.order_clipboard_contents[clean_order_id] = content
        print(f"DEBUG-STORE-MAP: 订单ID '{clean_order_id}' -> 收货信息长度: {len(content)}")
        self._log_info(f"已建立映射: 订单ID {clean_order_id} <-> 收货信息", "green")
        return True


    def _save_clipboard_mappings(self):
        """
        保存订单ID与剪贴板内容的映射到JSON文件
        """
        try:
            if not hasattr(self, 'order_clipboard_contents') or not self.order_clipboard_contents:
                self._log_info("没有剪贴板映射数据可保存", "orange")
                return
                
            # 准备要保存的数据
            save_data = {
                "updated_at": datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                "order_clipboard_mappings": self.order_clipboard_contents
            }
            
            # 保存到JSON文件
            with open('clipboard_mappings.json', 'w', encoding='utf-8') as f:
                json.dump(save_data, f, ensure_ascii=False, indent=2)
            
            self._log_info(f"已保存 {len(self.order_clipboard_contents)} 个订单ID与剪贴板内容的映射", "green")
        except Exception as e:
            self._log_info(f"保存剪贴板映射失败: {str(e)}", "red")
    

    def _load_clipboard_mappings(self):
        """
        从JSON文件加载订单ID与剪贴板内容的映射
        """
        # 确保字典已初始化
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
        
        try:
            if os.path.exists('clipboard_mappings.json'):
                with open('clipboard_mappings.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                if 'order_clipboard_mappings' in data:
                    # 将已加载的映射合并到现有字典
                    loaded_mappings = data['order_clipboard_mappings']
                    self.order_clipboard_contents.update(loaded_mappings)
                    self._log_info(f"已加载 {len(loaded_mappings)} 个订单ID与剪贴板内容的映射", "green")
                    return True
            
            self._log_info("未找到剪贴板映射文件或文件格式不正确", "orange")
            return False
        except Exception as e:
            self._log_info(f"加载剪贴板映射失败: {str(e)}", "red")
            return False


    def _clean_existing_clipboard_mappings(self):
        """清理现有的剪贴板映射，移除异常长度或无效的内容"""
        if not hasattr(self, 'order_clipboard_contents') or not self.order_clipboard_contents:
            return 0
            
        cleaned_count = 0
        for order_id, content in list(self.order_clipboard_contents.items()):
            # 检查内容长度
            if len(content) > 1000:
                self._log_info(f"清理异常长度的收货信息: 订单ID {order_id}, 长度 {len(content)}字符", "orange")
                del self.order_clipboard_contents[order_id]
                cleaned_count += 1
                continue
                
            # 验证内容有效性
            is_valid, confidence, reason = self._is_valid_shipping_info(content)
            if not is_valid:
                self._log_info(f"清理无效的收货信息: 订单ID {order_id}, 原因: {reason}", "orange")
                del self.order_clipboard_contents[order_id]
                cleaned_count += 1
                
        if cleaned_count > 0:
            self._log_info(f"已清理 {cleaned_count} 条异常收货信息映射", "green")
            # 保存清理后的映射
            self._save_clipboard_mappings()
            
        return cleaned_count
            

    def _start_clipboard_monitor(self):
        """
        启动剪贴板内容监听线程
        """
        if self.clipboard_monitor_active:
            return
            
        self.clipboard_monitor_active = True
        self._log_info("启动剪贴板监听...", "blue")
        
        def monitor_clipboard():
            last_content = ""
            while self.clipboard_monitor_active:
                try:
                    current_content = pyperclip.paste()
                    if current_content != last_content and current_content.strip():
                        last_content = current_content
                        self.last_known_clipboard = current_content
                        
                        # 尝试提取当前订单ID
                        current_order_id = self.current_order_id or self._extract_current_order_id()
                        
                        # 记录日志但不打印剪贴板内容以避免日志过大
                        content_length = len(current_content)
                        self._log_info(f"[监听器] 检测到剪贴板变化，长度: {content_length}，当前订单ID: {current_order_id}", "blue")
                        
                        # 如果有订单ID，则建立映射
                        if current_order_id and content_length > 10:
                            self._store_clipboard_content("监听器捕获", current_content, current_order_id)
                except:
                    pass
                    
                # 减少CPU使用
                time.sleep(0.5)
        
        # 启动监听线程
        self.clipboard_monitor_thread = threading.Thread(target=monitor_clipboard, daemon=True)
        self.clipboard_monitor_thread.start()
    

    def _stop_clipboard_monitor(self):
        """
        停止剪贴板监听线程
        """
        self.clipboard_monitor_active = False
        if self.clipboard_monitor_thread:
            # 等待线程结束
            self.clipboard_monitor_thread.join(1.0)
            self.clipboard_monitor_thread = None
            self._log_info("停止剪贴板监听", "blue")
            

    def _manually_associate_clipboard_with_order_id(self, clipboard_content=None):
        """
        手动关联剪贴板内容与订单ID的对话框
        
        参数:
        - clipboard_content: 可选的剪贴板内容，如果不提供将自动获取
        
        返回:
        - 是否成功关联
        """
        if not clipboard_content:
            clipboard_content = self._get_clipboard_content()
        
        if not clipboard_content or not clipboard_content.strip():
            messagebox.showerror("错误", "剪贴板内容为空，无法关联")
            return False
            
        # 创建对话框
        dialog = tk.Toplevel(self.root)
        dialog.title("手动关联订单ID")
        dialog.geometry("500x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 设置为模态窗口
        dialog.focus_set()
        
        # 预览剪贴板内容
        ttk.Label(dialog, text="剪贴板内容预览:").pack(pady=5)
        preview_frame = ttk.Frame(dialog)
        preview_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        preview_text = ScrolledText(preview_frame, height=10)
        preview_text.pack(fill=tk.BOTH, expand=True)
        preview_text.insert(tk.END, clipboard_content[:500] + ("..." if len(clipboard_content) > 500 else ""))
        preview_text.config(state=tk.DISABLED)
        
        # 订单ID输入
        id_frame = ttk.Frame(dialog)
        id_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(id_frame, text="输入订单ID:").pack(side=tk.LEFT)
        order_id_var = tk.StringVar()
        order_id_entry = ttk.Entry(id_frame, textvariable=order_id_var, width=30)
        order_id_entry.pack(side=tk.LEFT, padx=5)
        
        # 最近使用的订单ID
        if self.last_order_ids:
            recent_frame = ttk.LabelFrame(dialog, text="最近使用的订单ID")
            recent_frame.pack(fill=tk.X, padx=10, pady=5)
            
            for order_id in reversed(self.last_order_ids):
                def set_id(oid=order_id):
                    order_id_var.set(oid)
                
                ttk.Button(
                    recent_frame,
                    text=order_id,
                    command=set_id
                ).pack(side=tk.LEFT, padx=5, pady=5)
        
        # 按钮区域
        btn_frame = ttk.Frame(dialog)
        btn_frame.pack(fill=tk.X, pady=10)
        
        result = {"confirmed": False, "order_id": ""}
        
        def on_confirm():
            input_id = order_id_var.get().strip()
            if not input_id:
                messagebox.showerror("错误", "请输入订单ID")
                return
                
            result["confirmed"] = True
            result["order_id"] = input_id
            dialog.destroy()
            
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(btn_frame, text="确认关联", command=on_confirm).pack(side=tk.RIGHT, padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.RIGHT, padx=10)
        
        # 等待对话框关闭
        dialog.wait_window()
        
        # 处理结果
        if result["confirmed"] and result["order_id"]:
            self._store_clipboard_content("手动关联", clipboard_content, result["order_id"])
            
            # 将ID添加到最近使用列表
            if result["order_id"] not in self.last_order_ids:
                self.last_order_ids.append(result["order_id"])
                # 保留最近的5个ID
                if len(self.last_order_ids) > 5:
                    self.last_order_ids.pop(0)
            
            # 更新当前订单ID
            self.current_order_id = result["order_id"]
            
            self._log_info(f"已手动关联订单ID: {result['order_id']} 与剪贴板内容", "green")
            self._save_clipboard_mappings()
            return True
            
        return False


    def _is_valid_shipping_info(self, content):
        """
        验证内容是否为有效的收货信息，采用更灵活的规则
        返回 (is_valid, confidence_score, reason)
        """
        import re
        
        if not content or not isinstance(content, str):
            return False, 0, "内容为空或类型错误"
        
        content = content.strip()
        if len(content) < 5:  # 内容过短，肯定不是有效收货信息
            return False, 0, "内容过短"
            
        # 添加最大长度限制 - 正常收货信息不应超过1000字符
        if len(content) > 1000:
            return False, 0, f"内容过长 ({len(content)}字符)，超出正常收货信息范围"
        
        # 1. 首先检查是否包含明显的错误日志内容
        error_indicators = [
            r"\[\d{2}:\d{2}:\d{2}\].*错误",
            r"\[\d{2}:\d{2}:\d{2}\].*异常",
            r"\[\d{2}:\d{2}:\d{2}\].*Exception",
            r"\[\d{2}:\d{2}:\d{2}\].*Error",
            r"Traceback \(most recent call last\)",
            r"Exception in thread",
            r"cannot access local variable",
            r"ImportError:",
            r"AttributeError:",
            r"TypeError:",
            r"ValueError:"
        ]
        
        # 增加代码特征检测
        code_indicators = [
            r"def\s+\w+\(.*\):", 
            r"class\s+\w+\(.*\):",
            r"import\s+\w+",
            r"from\s+\w+\s+import",
            r"return\s+",
            r"if\s+.*:",
            r"else:",
            r"elif\s+.*:",
            r"for\s+.*\s+in\s+.*:",
            r"while\s+.*:",
            r"try:",
            r"except",
            r"finally:",
            r"```python",
            r"```"
        ]
        
        # 增加JSON格式检测 - 防止将元素配置数据误认为收货信息
        json_indicators = [
            r'"full_text"\s*:',
            r'"chinese_only"\s*:',
            r'"xpath"\s*:',
            r'"action"\s*:',
            r'"custom_name"\s*:',
            r'"element_type"\s*:',
            r'"captured_at"\s*:',
            r'\[\s*\{.*"full_text"',
            r'\{.*"xpath".*\}',
            r'/html/body/div\[\d+\]'
        ]
        
        for pattern in error_indicators:
            if re.search(pattern, content):
                return False, 0, f"内容包含错误日志特征: {pattern}"
        
        # 检查代码特征
        for pattern in code_indicators:
            if re.search(pattern, content):
                return False, 0, f"内容包含代码特征: {pattern}"
        
        # 检查JSON格式特征 - 防止将元素配置数据误认为收货信息
        for pattern in json_indicators:
            if re.search(pattern, content):
                return False, 0, f"内容包含JSON/元素配置特征: {pattern}"
        
        # 2. 检查是否包含手机号码 (这是收货信息的强特征)
        # 支持手机号和座机号两种格式
        phone_patterns = [
            r"1[3-9]\d{9}",  # 手机号
            r"\d{3,4}-\d{7,8}"  # 座机号 如 021-53395199
        ]
        
        has_phone = any(re.search(pattern, content) for pattern in phone_patterns)
        
        # 3. 检查是否包含地址特征 (省/市/区/县等)
        address_patterns = [
            r"(省|市|区|县|镇|乡|村|路|街|号楼|单元|室)",
            r"[东南西北中]门",
            r"[A-Za-z0-9#@\-]+号?库"  # 匹配仓库编号，如"2号库@DX-5E74D2M6D-F#"
        ]
        
        has_address = any(re.search(pattern, content) for pattern in address_patterns)
        
        # 4. 检查行数 - 有效的收货信息通常至少有2-3行
        lines = [line for line in content.split('\n') if line.strip()]
        has_multiple_lines = len(lines) >= 2
        
        # 5. 检查是否包含典型的收货信息格式特征
        # - 第一行通常是姓名（不做严格验证，因为可能是任何字符）
        # - 第二行通常是电话号码
        # - 后面几行是地址
        
        format_confidence = 0
        if has_multiple_lines:
            # 检查第二行是否包含电话号码
            if len(lines) >= 2 and any(re.search(pattern, lines[1]) for pattern in phone_patterns):
                format_confidence += 30
            
            # 检查第三行开始是否包含地址特征
            if len(lines) >= 3 and any(re.search(pattern, '\n'.join(lines[2:])) for pattern in address_patterns):
                format_confidence += 30
        
        # 计算总体可信度
        confidence = 0
        
        # 手机号是强特征
        if has_phone:
            confidence += 40
        
        # 地址特征也很重要
        if has_address:
            confidence += 30
        
        # 多行格式增加可信度
        if has_multiple_lines:
            confidence += 10
        
        # 典型格式特征
        confidence += format_confidence * 0.2  # 权重较低
        
        # 检查是否包含错误日志常见词汇（降低可信度）
        suspicious_terms = [
            "error", "exception", "failed", "undefined", "null", 
            "错误", "异常", "失败", "未定义", "空值"
        ]
        
        for term in suspicious_terms:
            if term.lower() in content.lower():
                confidence -= 15
                break
        
        # 最终判断
        is_valid = confidence >= 50  # 降低阈值，更宽松的判断
        
        reason = (
            f"电话:{has_phone}, 地址:{has_address}, "
            f"多行格式:{has_multiple_lines}, 格式特征:{format_confidence}, "
            f"可信度:{confidence}"
        )
        
        return is_valid, confidence, reason
    

    def _batch_review_shipping_info(self, orders_to_review):
        """提供批量审核收货信息的界面"""
        from tkinter import Toplevel, Label, Button, Text, Scrollbar, Frame, StringVar, Entry
        import tkinter as tk
        
        if not orders_to_review:
            return
        
        # 创建审核窗口
        review_window = Toplevel(self.root)
        review_window.title("收货信息审核")
        review_window.geometry("800x600")
        review_window.transient(self.root)
        review_window.grab_set()
        
        # 创建说明标签
        Label(review_window, text="以下订单的收货信息可能存在问题，请逐一审核:", font=("Arial", 12)).pack(pady=10)
        
        # 创建滚动区域
        main_frame = Frame(review_window)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(main_frame)
        scrollbar = Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 存储修改后的内容
        modified_contents = {}
        
        # 为每个需要审核的订单创建编辑区域
        for i, (order_id, content, confidence, reason, flag) in enumerate(orders_to_review):
            frame = Frame(scrollable_frame, bd=2, relief="groove")
            frame.pack(fill="x", expand=True, pady=5)
            
            # 订单信息
            Label(frame, text=f"订单 {i+1}/{len(orders_to_review)}: {order_id}", font=("Arial", 11, "bold")).pack(anchor="w")
            Label(frame, text=f"可信度: {confidence}分 | 原因: {reason}", fg="orange").pack(anchor="w")
            Label(frame, text=f"标记: {flag}", fg="red").pack(anchor="w")
            
            # 内容编辑区
            Label(frame, text="收货信息:").pack(anchor="w")
            text_widget = Text(frame, height=5, width=80)
            text_widget.pack(fill="x", padx=5, pady=5)
            text_widget.insert("1.0", content)
            
            # 保存修改内容的回调
            def save_content(order_id=order_id, text_widget=text_widget):
                modified_contents[order_id] = text_widget.get("1.0", "end-1c")
            
            # 操作按钮
            btn_frame = Frame(frame)
            btn_frame.pack(fill="x", pady=5)
            
            Button(btn_frame, text="确认无误", command=lambda oid=order_id: self._confirm_valid_content(oid, review_window)).pack(side="left", padx=5)
            Button(btn_frame, text="保存修改", command=lambda: save_content()).pack(side="left", padx=5)
            Button(btn_frame, text="删除此映射", command=lambda oid=order_id: self._delete_mapping(oid, review_window)).pack(side="left", padx=5)
        
        # 底部按钮
        bottom_frame = Frame(review_window)
        bottom_frame.pack(fill="x", pady=10)
        
        def on_finish():
            # 应用所有修改
            for order_id, new_content in modified_contents.items():
                if order_id in self.order_clipboard_contents and new_content.strip():
                    self.order_clipboard_contents[order_id] = new_content
                    self._log_info(f"已更新订单 {order_id} 的收货信息", "green")
            
            # 保存修改
            self._save_clipboard_mappings()
            
            # 清除需要审核的标记
            if hasattr(self, 'orders_need_review'):
                self.orders_need_review.clear()
            
            review_window.destroy()
        
        Button(bottom_frame, text="完成审核", command=on_finish, font=("Arial", 12)).pack(side="right", padx=10)
        Button(bottom_frame, text="取消", command=review_window.destroy, font=("Arial", 12)).pack(side="right", padx=10)
        
        # 等待窗口关闭
        self.root.wait_window(review_window)


    def _confirm_valid_content(self, order_id, parent_window=None):
        """确认订单内容有效，从需要审核的列表中移除"""
        if hasattr(self, 'orders_need_review') and order_id in self.orders_need_review:
            self.orders_need_review.remove(order_id)
            self._log_info(f"已确认订单 {order_id} 的收货信息有效", "green")
        
        # 更新界面反馈
        if parent_window:
            from tkinter import Label
            feedback = Label(parent_window, text=f"已确认订单 {order_id} 有效", fg="green")
            feedback.pack()
            parent_window.after(2000, feedback.destroy)


    def _delete_mapping(self, order_id, parent_window=None):
        """删除指定订单的映射"""
        if order_id in self.order_clipboard_contents:
            del self.order_clipboard_contents[order_id]
            self._log_info(f"已删除订单 {order_id} 的收货信息映射", "orange")
            
            # 从需要审核的列表中移除
            if hasattr(self, 'orders_need_review') and order_id in self.orders_need_review:
                self.orders_need_review.remove(order_id)
            
            # 保存修改
            self._save_clipboard_mappings()
        
        # 更新界面反馈
        if parent_window:
            from tkinter import Label
            feedback = Label(parent_window, text=f"已删除订单 {order_id} 的映射", fg="orange")
            feedback.pack()
            parent_window.after(2000, feedback.destroy)
            
