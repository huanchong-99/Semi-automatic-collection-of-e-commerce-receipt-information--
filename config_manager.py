#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置管理相关

从原始main.py文件中拆分出来的模块
"""

from utils import *

class ConfigManager:
    """配置管理相关"""
    
    def __init__(self):
        pass

    def _load_offset_config(self):
        """加载元素偏移量配置"""
        try:
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "offset_config.json")
            if os.path.exists(config_path):
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    
                # 兼容旧版配置
                if "offset" in config and "x" in config["offset"] and "y" in config["offset"]:
                    # 旧格式，创建默认字典，并填入旧的偏移量
                    default_x = config["offset"]["x"]
                    default_y = config["offset"]["y"]
                    self._log_info(f"已从配置文件加载偏移量: X={default_x}, Y={default_y}", "blue")
                    # 初始化元素偏移量字典
                    self.element_offsets = {elem: {"x": default_x, "y": default_y} for elem in ["订单编号", "商品名称", "成交金额", "查看1", "查看2", "复制完整收货信息"]}
                elif "element_offsets" in config:
                    # 新格式，直接加载
                    self.element_offsets = config["element_offsets"]
                    elements_loaded = len(self.element_offsets)
                    self._log_info(f"已加载{elements_loaded}个元素的偏移量配置", "blue")
                    # 显示部分元素的偏移量
                    for name, offset in list(self.element_offsets.items())[:3]:
                        self._log_info(f"元素'{name}'偏移量: X={offset['x']}, Y={offset['y']}", "blue")
                        
            else:
                # 文件不存在，创建默认字典
                self.element_offsets = {elem: {"x": 0, "y": 0} for elem in ["订单编号", "商品名称", "成交金额", "查看1", "查看2", "复制完整收货信息"]}
                self._log_info("偏移量配置文件不存在，已创建默认配置", "blue")
                # 保存默认配置
                self._save_offset_config()
        except Exception as e:
            self.element_offsets = {elem: {"x": 0, "y": 0} for elem in ["订单编号", "商品名称", "成交金额", "查看1", "查看2", "复制完整收货信息"]}
            self._log_info(f"加载偏移量配置失败: {str(e)}，使用默认值", "orange")
    

    def _save_offset_config(self):
        """保存元素偏移量配置"""
        try:
            config = {
                "element_offsets": self.element_offsets,
                "updated_at": time.strftime("%Y-%m-%d %H:%M:%S")
            }
            
            config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "offset_config.json")
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            self._log_info(f"已保存{len(self.element_offsets)}个元素的偏移量配置", "green")
            
        except Exception as e:
            self._log_info(f"保存偏移量配置失败: {str(e)}", "red")


    def _reset_offsets(self):
        """重置所有元素的偏移量"""
        # 重置为空字典或默认值
        self.element_offsets = {elem: {"x": 0, "y": 0} for elem in ["订单编号", "商品名称", "成交金额", "查看1", "查看2", "复制完整收货信息"]}
        
        # 更新显示
        if hasattr(self, 'offset_label') and self.offset_label:
            self.offset_label.config(text="元素偏移量: 已重置")
        
        # 保存到配置文件
        self._save_offset_config()
        self._log_info("已重置所有元素的偏移量", "blue")
    

    def _show_offset_manager(self):
        """显示元素偏移量管理界面"""
        # 创建偏移量管理窗口
        offset_win = tk.Toplevel(self.root)
        offset_win.title("元素偏移量管理")
        offset_win.geometry("500x400")
        offset_win.transient(self.root)
        offset_win.grab_set()
        
        # 创建表格显示所有元素的偏移量
        columns = ("元素名称", "X偏移", "Y偏移", "操作")
        tree = ttk.Treeview(offset_win, columns=columns, show="headings", selectmode="browse")
        
        # 设置列标题
        for col in columns:
            tree.heading(col, text=col)
            if col == "元素名称":
                tree.column(col, width=200, anchor="w")
            elif col in ("X偏移", "Y偏移"):
                tree.column(col, width=80, anchor="center")
            else:
                tree.column(col, width=100, anchor="center")
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(offset_win, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 填充数据
        def refresh_tree():
            # 清空现有数据
            tree.delete(*tree.get_children())
            
            # 添加元素数据
            for name, offset in self.element_offsets.items():
                tree.insert("", "end", values=(name, offset.get("x", 0), offset.get("y", 0), "重置"))
        
        refresh_tree()
        
        # 创建按钮框架
        btn_frame = ttk.Frame(offset_win)
        btn_frame.pack(fill="x", padx=10, pady=10)
        
        # 添加重置所有按钮
        reset_all_btn = ttk.Button(
            btn_frame, 
            text="重置所有偏移量", 
            command=lambda: [self._reset_offsets(), refresh_tree()]
        )
        reset_all_btn.pack(side="left", padx=5)
        
        # 添加导出按钮
        export_btn = ttk.Button(
            btn_frame, 
            text="导出配置", 
            command=self._export_offset_config
        )
        export_btn.pack(side="left", padx=5)
        
        # 添加导入按钮
        import_btn = ttk.Button(
            btn_frame, 
            text="导入配置", 
            command=lambda: [self._import_offset_config(), refresh_tree()]
        )
        import_btn.pack(side="left", padx=5)
        
        # 添加关闭按钮
        close_btn = ttk.Button(
            btn_frame, 
            text="关闭", 
            command=offset_win.destroy
        )
        close_btn.pack(side="right", padx=5)
        
        # 处理点击"重置"操作
        def on_tree_click(event):
            # 获取点击的行和列
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            
            if item and column == "#4":  # 操作列
                # 获取元素名称
                values = tree.item(item, "values")
                elem_name = values[0]
                
                # 重置该元素的偏移量
                if elem_name in self.element_offsets:
                    self.element_offsets[elem_name] = {"x": 0, "y": 0}
                    self._save_offset_config()
                    self._log_info(f"已重置元素'{elem_name}'的偏移量", "blue")
                    refresh_tree()
        
        tree.bind("<Button-1>", on_tree_click)
    

    def _export_offset_config(self):
        """导出元素偏移量配置到文件"""
        try:
            file_path = filedialog.asksaveasfilename(
                title="导出偏移量配置",
                defaultextension=".json",
                filetypes=[("JSON文件", "*.json")],
                initialfile="element_offsets_config.json"
            )
            
            if not file_path:
                return
                
            config = {
                "element_offsets": self.element_offsets,
                "exported_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "description": "元素偏移量配置文件"
            }
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
                
            self._log_info(f"已导出偏移量配置到: {file_path}", "green")
            
        except Exception as e:
            self._log_info(f"导出偏移量配置失败: {str(e)}", "red")
            

    def _import_offset_config(self):
        """从文件导入元素偏移量配置"""
        try:
            file_path = filedialog.askopenfilename(
                title="导入偏移量配置",
                filetypes=[("JSON文件", "*.json")],
                initialdir=os.path.dirname(os.path.abspath(__file__))
            )
            
            if not file_path:
                return
                
            with open(file_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            if "element_offsets" in config:
                # 保存原始配置以防导入失败
                original_offsets = copy.deepcopy(self.element_offsets)
                
                try:
                    # 验证导入的配置格式
                    for name, offset in config["element_offsets"].items():
                        if not isinstance(name, str) or not isinstance(offset, dict):
                            raise ValueError("配置格式错误")
                        if "x" not in offset or "y" not in offset:
                            raise ValueError(f"元素'{name}'的偏移量格式错误")
                            
                    # 更新配置
                    self.element_offsets = config["element_offsets"]
                    self._save_offset_config()
                    self._log_info(f"已导入偏移量配置，包含{len(self.element_offsets)}个元素", "green")
                    
                except Exception as e:
                    # 恢复原始配置
                    self.element_offsets = original_offsets
                    self._log_info(f"导入偏移量配置验证失败: {str(e)}", "red")
            else:
                self._log_info("无效的偏移量配置文件", "red")
                
        except Exception as e:
            self._log_info(f"导入偏移量配置失败: {str(e)}", "red")

