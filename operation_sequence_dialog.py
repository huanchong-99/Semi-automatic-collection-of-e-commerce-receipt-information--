import tkinter as tk
from tkinter import ttk
from tkinter.simpledialog import askinteger
from tkinter import messagebox


def askinteger_topmost(parent, title, prompt, **kwargs):
    """创建一个置顶的整数输入对话框"""
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.geometry("300x150")
    dialog.transient(parent)
    dialog.grab_set()
    dialog.attributes('-topmost', True)
    
    # 居中显示
    dialog.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() // 2) - (300 // 2)
    y = parent.winfo_rooty() + (parent.winfo_height() // 2) - (150 // 2)
    dialog.geometry(f"+{x}+{y}")
    
    result = [None]
    
    # 创建输入框
    ttk.Label(dialog, text=prompt).pack(pady=10)
    
    entry_var = tk.StringVar()
    if 'initialvalue' in kwargs:
        entry_var.set(str(kwargs['initialvalue']))
    
    entry = ttk.Entry(dialog, textvariable=entry_var)
    entry.pack(pady=5)
    entry.focus()
    
    def on_ok():
        try:
            value = int(entry_var.get())
            if 'minvalue' in kwargs and value < kwargs['minvalue']:
                messagebox.showerror("错误", f"值不能小于 {kwargs['minvalue']}")
                return
            if 'maxvalue' in kwargs and value > kwargs['maxvalue']:
                messagebox.showerror("错误", f"值不能大于 {kwargs['maxvalue']}")
                return
            result[0] = value
            dialog.destroy()
        except ValueError:
            messagebox.showerror("错误", "请输入有效的整数")
    
    def on_cancel():
        dialog.destroy()
    
    # 按钮
    btn_frame = ttk.Frame(dialog)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.LEFT, padx=5)
    ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side=tk.LEFT, padx=5)
    
    # 绑定回车键
    entry.bind('<Return>', lambda e: on_ok())
    
    dialog.wait_window()
    return result[0]


def messagebox_topmost(parent, msg_type, title, message):
    """创建一个置顶的消息框"""
    if msg_type == "warning":
        messagebox.showwarning(title, message, parent=parent)
    elif msg_type == "error":
        messagebox.showerror(title, message, parent=parent)
    elif msg_type == "info":
        messagebox.showinfo(title, message, parent=parent)


class OperationSequenceDialog(tk.Toplevel):
    """操作选择与排序对话框，用于配置要执行的操作及其顺序"""
    
    def __init__(self, parent, elements_data=None):
        """
        初始化对话框
        parent: 父窗口
        elements_data: 从JSON文件加载的元素数据列表
        """
        super().__init__(parent)
        self.parent = parent
        self.elements_data = elements_data or []
        self.result = None
        
        # 设置窗口属性
        self.title("操作选择与顺序配置")
        self.geometry("800x500")  # 较大的窗口以显示表格和控件
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        
        # 居中显示
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"+{x}+{y}")
        
        # 创建主框架
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 创建操作表格
        self._create_operations_table(main_frame)
        
        # 指定一个元素作为订单数量元素的框架
        count_frame = ttk.LabelFrame(main_frame, text="订单数量元素选择")
        count_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(count_frame, text="选择一个包含订单数量的元素(可选):").pack(side=tk.LEFT, padx=5)
        self.count_element_var = tk.StringVar(value="")
        self.count_element_combo = ttk.Combobox(
            count_frame, 
            textvariable=self.count_element_var,
            state="readonly",
            width=30
        )
        # 填充下拉框选项
        count_options = ["(不使用自动检测)"] + [elem["name"] for elem in self.elements_data]
        self.count_element_combo["values"] = count_options
        self.count_element_combo.current(0)  # 默认选择"不使用自动检测"
        self.count_element_combo.pack(side=tk.LEFT, padx=5)
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        ttk.Button(btn_frame, text="全部选中", command=self._select_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="全部取消", command=self._deselect_all).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="按名称排序", command=self._sort_by_name).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(btn_frame, text="确定", command=self._on_ok).pack(side=tk.RIGHT, padx=5)
        ttk.Button(btn_frame, text="取消", command=self._on_cancel).pack(side=tk.RIGHT, padx=5)
        
        # 创建元素变量
        self._create_element_variables()
        
        # 填充表格
        self._populate_table()
        
        # 等待窗口关闭
        self.wait_window(self)
    
    def _create_operations_table(self, parent):
        """创建操作表格"""
        # 创建表格框架
        table_frame = ttk.LabelFrame(parent, text="操作配置")
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建表格视图
        columns = ("enabled", "order", "name", "action", "loop_mode", "preview")  # 添加loop_mode列
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="browse")
        
        # 设置列标题
        self.tree.heading("enabled", text="启用")
        self.tree.heading("order", text="顺序")
        self.tree.heading("name", text="元素名称")
        self.tree.heading("action", text="操作类型")
        self.tree.heading("loop_mode", text="循环模式")  # 新增列标题
        self.tree.heading("preview", text="预览")
        
        # 设置列宽度
        self.tree.column("enabled", width=50, anchor=tk.CENTER)
        self.tree.column("order", width=50, anchor=tk.CENTER)
        self.tree.column("name", width=200, anchor=tk.W)
        self.tree.column("action", width=150, anchor=tk.CENTER)
        self.tree.column("loop_mode", width=100, anchor=tk.CENTER)  # 设置新列宽度
        self.tree.column("preview", width=80, anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局表格和滚动条
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 绑定事件处理
        self.tree.bind("<ButtonRelease-1>", self._on_tree_click)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
    
    def _create_element_variables(self):
        """为每个元素创建变量"""
        # 存储每个元素的启用状态、顺序、操作类型和循环模式
        self.element_vars = []
        
        for i, elem in enumerate(self.elements_data):
            enabled_var = tk.BooleanVar(value=elem.get("enabled", True))
            order_var = tk.IntVar(value=elem.get("order", i+1))
            action_var = tk.StringVar(value=elem.get("action", "getText"))
            loop_mode_var = tk.StringVar(value=elem.get("loop_mode", "always"))  # 默认为"始终循环"
            
            self.element_vars.append({
                "element_id": elem["element_id"],
                "enabled_var": enabled_var,
                "order_var": order_var,
                "action_var": action_var,
                "loop_mode_var": loop_mode_var  # 添加循环模式变量
            })
    
    def _populate_table(self):
        """填充表格数据"""
        # 先清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 按顺序排序元素
        sorted_elements = sorted(self.elements_data, key=lambda x: x.get("order", 0))
        
        # 添加元素到表格
        for elem in sorted_elements:
            # 获取对应的变量
            vars_dict = next((v for v in self.element_vars if v["element_id"] == elem["element_id"]), None)
            if vars_dict:
                # 创建表格项
                enabled_text = "✓" if vars_dict["enabled_var"].get() else ""
                order = vars_dict["order_var"].get()
                name = elem.get("name", "未命名")
                action = vars_dict["action_var"].get()
                action_text = {
                    "getText": "获取文本",
                    "click": "点击元素",
                    "clickAndGetClipboard": "点击并获取剪贴板"
                }.get(action, action)
                
                # 获取循环模式文本
                loop_mode = vars_dict["loop_mode_var"].get()
                loop_mode_text = "单次循环" if loop_mode == "once" else "始终循环"
                
                item_values = (enabled_text, order, name, action_text, loop_mode_text, "预览")
                item_id = self.tree.insert("", "end", values=item_values)
                
                # 存储元素ID与表格项的映射
                self.tree.item(item_id, tags=(str(elem["element_id"]),))
    
    def _on_tree_click(self, event):
        """处理表格点击事件"""
        # 获取点击的列
        region = self.tree.identify_region(event.x, event.y)
        column = self.tree.identify_column(event.x)
        
        if region == "cell":
            item_id = self.tree.identify_row(event.y)
            if not item_id:
                return
                
            # 获取元素ID
            elem_id = int(self.tree.item(item_id)["tags"][0])
            
            # 根据点击的列执行不同操作
            if column == "#1":  # 启用列
                self._toggle_element_enabled(elem_id, item_id)
            elif column == "#2":  # 顺序列
                self._edit_element_order(elem_id, item_id)
            elif column == "#4":  # 操作类型列
                self._edit_element_action(elem_id, item_id)
            elif column == "#5":  # 循环模式列
                self._edit_element_loop_mode(elem_id, item_id)
            elif column == "#6":  # 预览列
                self._preview_element(elem_id)
    
    def _on_tree_double_click(self, event):
        """处理表格双击事件"""
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            item_id = self.tree.identify_row(event.y)
            if item_id:
                # 双击时编辑元素顺序
                elem_id = int(self.tree.item(item_id)["tags"][0])
                self._edit_element_order(elem_id, item_id)
    
    def _toggle_element_enabled(self, elem_id, item_id):
        """切换元素启用状态"""
        # 查找对应的变量
        for var in self.element_vars:
            if var["element_id"] == elem_id:
                # 切换状态
                new_state = not var["enabled_var"].get()
                var["enabled_var"].set(new_state)
                
                # 更新表格显示
                current_values = list(self.tree.item(item_id, "values"))
                current_values[0] = "✓" if new_state else ""
                self.tree.item(item_id, values=current_values)
                break
    
    def _edit_element_order(self, elem_id, item_id):
        """编辑元素顺序并自动调整其他元素的顺序"""
        # 查找对应的变量
        for var in self.element_vars:
            if var["element_id"] == elem_id:
                # 当前顺序
                current_order = var["order_var"].get()
                
                # 创建一个简单的输入对话框
                new_order = askinteger_topmost(self, "编辑顺序", f"请输入元素的执行顺序 (1-{len(self.elements_data)}):", initialvalue=current_order, minvalue=1, maxvalue=len(self.elements_data))
                
                # 如果用户取消或输入相同的值，则不进行任何更改
                if new_order is None or new_order == current_order:
                    return
                
                # 更新当前元素的顺序
                var["order_var"].set(new_order)
                
                # 调整其他元素的顺序以避免重复
                self._adjust_other_elements_order(elem_id, current_order, new_order)
                
                # 更新表格显示
                current_values = list(self.tree.item(item_id, "values"))
                current_values[1] = str(new_order)
                self.tree.item(item_id, values=tuple(current_values))
                
                # 重新排序表格
                self._resort_table()
                break
    
    def _edit_element_action(self, elem_id, item_id):
        """编辑元素操作类型"""
        # 查找对应的变量
        for var in self.element_vars:
            if var["element_id"] == elem_id:
                # 当前操作类型
                current_action = var["action_var"].get()
                
                # 创建操作类型选择对话框
                dialog = tk.Toplevel(self)
                dialog.title("选择操作类型")
                dialog.geometry("300x200")
                dialog.transient(self)
                dialog.grab_set()
                
                # 设置窗口置顶
                dialog.attributes('-topmost', True)
                
                # 居中显示
                dialog.update_idletasks()
                x = self.winfo_rootx() + (self.winfo_width() // 2) - (300 // 2)
                y = self.winfo_rooty() + (self.winfo_height() // 2) - (200 // 2)
                dialog.geometry(f"+{x}+{y}")
                
                # 创建单选按钮
                action_var = tk.StringVar(value=current_action)
                ttk.Label(dialog, text="请选择操作类型:").pack(pady=10)
                
                ttk.Radiobutton(
                    dialog, 
                    text="获取文本", 
                    variable=action_var, 
                    value="getText"
                ).pack(anchor=tk.W, padx=20, pady=5)
                
                ttk.Radiobutton(
                    dialog, 
                    text="点击元素", 
                    variable=action_var, 
                    value="click"
                ).pack(anchor=tk.W, padx=20, pady=5)
                
                ttk.Radiobutton(
                    dialog, 
                    text="点击并获取剪贴板内容", 
                    variable=action_var, 
                    value="clickAndGetClipboard"
                ).pack(anchor=tk.W, padx=20, pady=5)
                
                # 按钮区域
                btn_frame = ttk.Frame(dialog)
                btn_frame.pack(fill=tk.X, pady=15)
                
                # 确定按钮回调
                def on_ok():
                    new_action = action_var.get()
                    if new_action != current_action:
                        var["action_var"].set(new_action)
                        
                        # 更新表格显示
                        current_values = list(self.tree.item(item_id, "values"))
                        action_text = {
                            "getText": "获取文本",
                            "click": "点击元素",
                            "clickAndGetClipboard": "点击并获取剪贴板"
                        }.get(new_action, new_action)
                        current_values[3] = action_text
                        self.tree.item(item_id, values=current_values)
                    
                    dialog.destroy()
                    # --- 焦点恢复 ---
                    try:
                        if hasattr(self.parent, '_manage_focus'):
                            self.parent._manage_focus()
                            self.parent.root.after(100, self.parent._manage_focus)
                            self.parent._log_info("已尝试恢复主窗口焦点", "blue")
                    except Exception as e:
                        print(f"焦点恢复失败: {e}")
                
                ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.RIGHT, padx=5)
                ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
                
                # 等待对话框关闭
                self.wait_window(dialog)
                break
    
    def _edit_element_loop_mode(self, elem_id, item_id):
        """编辑元素循环模式"""
        # 查找对应的变量
        for var in self.element_vars:
            if var["element_id"] == elem_id:
                # 当前循环模式
                current_mode = var["loop_mode_var"].get()
                
                # 创建循环模式选择对话框
                dialog = tk.Toplevel(self)
                dialog.title("选择循环模式")
                dialog.geometry("300x200")
                dialog.transient(self)
                dialog.grab_set()
                
                # 设置窗口置顶
                dialog.attributes('-topmost', True)
                
                # 居中显示
                dialog.update_idletasks()
                x = self.winfo_rootx() + (self.winfo_width() // 2) - (300 // 2)
                y = self.winfo_rooty() + (self.winfo_height() // 2) - (200 // 2)
                dialog.geometry(f"+{x}+{y}")
                
                # 创建单选按钮
                mode_var = tk.StringVar(value=current_mode)
                ttk.Label(dialog, text="请选择循环模式:").pack(pady=10)
                
                ttk.Radiobutton(
                    dialog, 
                    text="单次循环 (仅在第一个订单执行)", 
                    variable=mode_var, 
                    value="once"
                ).pack(anchor=tk.W, padx=20, pady=5)
                
                ttk.Radiobutton(
                    dialog, 
                    text="始终循环 (每个订单都执行)", 
                    variable=mode_var, 
                    value="always"
                ).pack(anchor=tk.W, padx=20, pady=5)
                
                # 按钮区域
                btn_frame = ttk.Frame(dialog)
                btn_frame.pack(fill=tk.X, pady=15)
                
                # 确定按钮回调
                def on_ok():
                    new_mode = mode_var.get()
                    if new_mode != current_mode:
                        var["loop_mode_var"].set(new_mode)
                        
                        # 更新表格显示
                        current_values = list(self.tree.item(item_id, "values"))
                        loop_mode_text = "单次循环" if new_mode == "once" else "始终循环"
                        current_values[4] = loop_mode_text
                        self.tree.item(item_id, values=current_values)
                    
                    dialog.destroy()
                    # --- 焦点恢复 ---
                    try:
                        if hasattr(self.parent, '_manage_focus'):
                            self.parent._manage_focus()
                            self.parent.root.after(100, self.parent._manage_focus)
                            self.parent._log_info("已尝试恢复主窗口焦点", "blue")
                    except Exception as e:
                        print(f"焦点恢复失败: {e}")
                
                ttk.Button(btn_frame, text="确定", command=on_ok).pack(side=tk.RIGHT, padx=5)
                ttk.Button(btn_frame, text="取消", command=dialog.destroy).pack(side=tk.RIGHT, padx=5)
                
                # 等待对话框关闭
                self.wait_window(dialog)
                break
    
    def _preview_element(self, elem_id):
        """预览元素（触发外部预览回调）"""
        # 查找对应的元素数据
        elem = next((e for e in self.elements_data if e["element_id"] == elem_id), None)
        if elem and "xpath" in elem:
            # 通知父窗口进行预览
            # 注意：这里需要父窗口提供一个预览方法
            if hasattr(self.parent, "_preview_element"):
                xpath = elem["xpath"]
                success = self.parent._preview_element(xpath)
                if not success:
                    messagebox_topmost(self, "warning", "预览失败", "无法定位或高亮显示元素，请检查XPath路径是否正确")
    
    def _resort_table(self):
        """根据顺序重新排序表格"""
        # 获取当前表格中的所有项
        items = []
        for item_id in self.tree.get_children():
            elem_id = int(self.tree.item(item_id)["tags"][0])
            var = next((v for v in self.element_vars if v["element_id"] == elem_id), None)
            if var:
                order = var["order_var"].get()
                items.append((item_id, elem_id, order))
        
        # 按顺序排序
        items.sort(key=lambda x: x[2])
        
        # 重新排列表格
        for index, (item_id, _, _) in enumerate(items):
            self.tree.move(item_id, "", index)
    
    def _adjust_other_elements_order(self, changed_elem_id, old_order, new_order):
        """调整其他元素的顺序以避免重复"""
        # 如果顺序没有变化，则不需要调整
        if old_order == new_order:
            return
            
        # 确定调整方向和范围
        if old_order < new_order:
            # 向下移动：将中间的元素顺序值减1
            for var in self.element_vars:
                if var["element_id"] != changed_elem_id:
                    current_order = var["order_var"].get()
                    if old_order < current_order <= new_order:
                        var["order_var"].set(current_order - 1)
        else:
            # 向上移动：将中间的元素顺序值加1
            for var in self.element_vars:
                if var["element_id"] != changed_elem_id:
                    current_order = var["order_var"].get()
                    if new_order <= current_order < old_order:
                        var["order_var"].set(current_order + 1)
        
        # 更新所有表格项的显示
        for item_id in self.tree.get_children():
            elem_id = int(self.tree.item(item_id)["tags"][0])
            var = next((v for v in self.element_vars if v["element_id"] == elem_id), None)
            if var:
                current_values = list(self.tree.item(item_id, "values"))
                current_values[1] = str(var["order_var"].get())
                self.tree.item(item_id, values=tuple(current_values))
    
    def _select_all(self):
        """选择所有元素"""
        for var in self.element_vars:
            var["enabled_var"].set(True)
        
        # 更新表格显示
        for item_id in self.tree.get_children():
            current_values = list(self.tree.item(item_id, "values"))
            current_values[0] = "✓"
            self.tree.item(item_id, values=current_values)
    
    def _deselect_all(self):
        """取消选择所有元素"""
        for var in self.element_vars:
            var["enabled_var"].set(False)
        
        # 更新表格显示
        for item_id in self.tree.get_children():
            current_values = list(self.tree.item(item_id, "values"))
            current_values[0] = ""
            self.tree.item(item_id, values=current_values)
    
    def _sort_by_name(self):
        """按名称排序元素顺序"""
        # 按名称排序元素
        named_elements = [(elem["element_id"], elem.get("name", "")) for elem in self.elements_data]
        named_elements.sort(key=lambda x: x[1])
        
        # 保存原始顺序，以便后续调整
        original_orders = {}
        for var in self.element_vars:
            original_orders[var["element_id"]] = var["order_var"].get()
        
        # 按名称顺序设置新的顺序值
        for i, (elem_id, _) in enumerate(named_elements):
            var = next((v for v in self.element_vars if v["element_id"] == elem_id), None)
            if var:
                # 记录原始顺序
                old_order = var["order_var"].get()
                # 设置新顺序
                new_order = i + 1
                
                if old_order != new_order:
                    var["order_var"].set(new_order)
        
        # 更新表格显示
        for item_id in self.tree.get_children():
            elem_id = int(self.tree.item(item_id)["tags"][0])
            var = next((v for v in self.element_vars if v["element_id"] == elem_id), None)
            if var:
                current_values = list(self.tree.item(item_id, "values"))
                current_values[1] = var["order_var"].get()
                self.tree.item(item_id, values=current_values)
        
        # 重新排序表格
        self._resort_table()
    
    def _on_ok(self):
        """确定按钮回调"""
        # 收集结果
        result = []
        
        # 处理订单数量元素选择
        count_element_name = self.count_element_var.get()
        
        for i, elem in enumerate(self.elements_data):
            var = next((v for v in self.element_vars if v["element_id"] == elem["element_id"]), None)
            if var:
                # 检查是否为订单数量元素
                # 如果选择了"(不使用自动检测)"，则所有元素的is_order_count都为False
                if count_element_name == "(不使用自动检测)":
                    is_order_count = False
                else:
                    is_order_count = (count_element_name == elem.get("name", ""))
                
                result.append({
                    "element_id": elem["element_id"],
                    "name": elem.get("name", f"元素{i+1}"),
                    "xpath": elem.get("xpath", ""),
                    "action": var["action_var"].get(),
                    "order": var["order_var"].get(),
                    "enabled": var["enabled_var"].get(),
                    "is_order_count": is_order_count,
                    "loop_mode": var["loop_mode_var"].get()  # 添加循环模式
                })
        
        # 保存结果
        self.result = result
        self.destroy()
    
    def _on_cancel(self):
        """取消按钮回调"""
        self.result = None
        self.destroy()