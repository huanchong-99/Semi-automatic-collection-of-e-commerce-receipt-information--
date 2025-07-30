#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
数据处理和导出相关

从原始main.py文件中拆分出来的模块
"""

from utils import *
from coordinate_cache import CoordinateCache
from data_cache_manager import get_cache_manager

class DataProcessor:
    """数据处理和导出相关"""
    
    def __init__(self):
        """初始化数据处理器"""
        self.coordinate_cache = CoordinateCache()
        self.cache_manager = get_cache_manager()  # 获取数据缓存管理器

    def run_actions_loop(self, manual_order_count=None):
        """主循环入口 - 支持模块化翻页"""
        if not self.driver:
            self._log_info('循环模式错误: 浏览器未连接。', 'red')
            return
            
        # 同步UI状态到实例属性
        self._sync_ui_modular_paging_state()
            
        # 检查是否启用模块化翻页
        if hasattr(self, 'use_modular_paging') and self.use_modular_paging:
            return self._run_modular_paging_loop(manual_order_count)
        else:
            return self._run_original_loop(manual_order_count)
    
    def _run_original_loop(self, manual_order_count=None):
        """原有的循环逻辑（保持不变）"""
        if not self.driver:
            self._log_info('循环模式错误: 浏览器未连接。', 'red')
            return
            
        # 确保order_clipboard_contents字典已初始化
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
            self._log_info("初始化订单ID与收货信息映射字典", "blue")
            print("DEBUG-RUN-ACTIONS-LOOP: 初始化订单ID与收货信息映射字典")
        
        # 检查是否有订单数量元素被标记为自动检测
        order_count_elements = [op for op in self.operation_sequence if op.get("is_order_count", False)]
        
        if order_count_elements:
            # 有订单数量元素，使用自动检测模式
            count_action = order_count_elements[0]
            num_items = 0
            try:
                self._log_info(f"执行: {count_action['name']} (自动获取总数)")
                # 使用智能查找元素
                element = self._find_element_smart(count_action['name'], count_action['xpath'])
                if not element:
                    self._log_info('无法找到总数元素，请检查XPath。', 'red')
                    return
                    
                count_text = element.text.strip()
                import re
                numbers = re.findall(r'\d+', count_text)
                if not numbers:
                    self._log_info(f'错误: 在 "{count_text}" 中未找到数字。', 'red')
                    return
                num_items = int(numbers[0])
                self._log_info(f'自动检测到项目总数: {num_items}')
            except Exception as e:
                self._log_info(f'自动获取总数失败: {e}', 'red')
                return
            if num_items == 0:
                self._log_info('项目总数为0，无需执行。')
                return
            
            # 过滤掉订单数量元素，只处理其他元素
            actions_to_loop = [op for op in self.operation_sequence if not op.get("is_order_count", False)]
        else:
            # 没有订单数量元素，使用手动输入模式
            self._log_info("未选择订单数量元素，使用手动输入的订单数量", "blue")
            if manual_order_count is None:
                self._log_info('错误：未提供手动输入的订单数量', 'red')
                return
            num_items = manual_order_count
            self._log_info(f"使用手动输入的订单数量: {num_items}", "blue")
            
            # 跳过名为"待发货的数量"的元素，处理其他所有元素
            actions_to_loop = [op for op in self.operation_sequence if op.get('name') != '待发货的数量']
            self._log_info(f"手动输入模式：跳过'待发货的数量'元素，将处理{len(actions_to_loop)}个其他元素", "blue")
            
        # 限制处理数量（如果启用了调试模式）
        if hasattr(self, 'confirm_click') and self.confirm_click.get():
            debug_limit = 3  # 调试模式下限制处理前3个订单
            if num_items > debug_limit:
                self._log_info(f'调试模式已启用，仅处理前{debug_limit}个订单。', 'blue')
                num_items = debug_limit
                
        # 确保actions_to_loop不为空
        if not actions_to_loop:
            self._log_info('错误：没有可执行的操作元素', 'red')
            return
            
        first_action_xpath = actions_to_loop[0]['xpath']
        xpath_pattern = None
        
        # 尝试使用参照点学习XPath模式
        if hasattr(self, 'ref2_xpath') and self.ref2_xpath:
            xpath_pattern = self._learn_xpath_pattern(first_action_xpath, self.ref2_xpath)
            if not xpath_pattern:
                self._log_info('无法从参照XPath中学习到规律，将回退到基本模式。', 'orange')
                
        # 如果无法从参照点学习，尝试基本模式
        if not xpath_pattern:
            self._log_info('使用基本循环模式（递增最后一个数字）...', 'info')
            import re
            base_index = -1
            matches = list(re.finditer(r'\[(\d+)\]', first_action_xpath))
            if matches:
                base_index = int(matches[-1].group(1))
                self._log_info(f'检测到列表起始索引为: {base_index}')
                xpath_pattern = {'diff_segment_index': len(first_action_xpath.split('/'))-1, 'template': '', 'start_index': base_index}
            else:
                self._log_info('循环模式错误: 无法在第一个操作的XPath中找到列表索引（如 [1], [2]）。无法继续循环。', 'red')
                return
                
        # 设置进度条最大值
        self.progress_bar["maximum"] = num_items
        self.progress_bar["value"] = 0
        self.progress_label.config(text=f"0/{num_items}")
        self._log_info(f"设置进度条最大值为: {num_items}", "blue")
        
        self.collected_data = []
        processed_order_ids = set()  # 用于检测重复订单
        consecutive_same_order = 0   # 连续重复订单计数
        
        for i in range(1, num_items+1):
            if not self.is_running:
                break
            
            # 检查重试标志 - 阶段3增强：配置化重试机制
            if hasattr(self, 'retry_current_order') and self.retry_current_order:
                # 使用重试管理器检查是否应该重试
                if hasattr(self, 'retry_manager') and not self.retry_manager.should_retry("order_processing", i):
                    self._log_info(f"[重试] 订单 {i} 已达到最大重试次数，跳过", "red")
                    from utils import create_retry_log_entry, write_retry_log
                    log_entry = create_retry_log_entry("retry_abandoned", f"order_{i}", {"reason": "max_attempts_reached"})
                    write_retry_log(log_entry)
                    self.retry_current_order = False
                    continue
                
                # 记录重试开始
                if hasattr(self, 'retry_manager'):
                    from utils import create_retry_log_entry, write_retry_log
                    log_entry = create_retry_log_entry("retry_start", f"order_{i}", {"attempt": self.retry_manager.retry_attempts.get(f"order_processing_{i}", 0) + 1})
                    write_retry_log(log_entry)
                
                # 阶段3修复：重试前检查暂停状态，确保协调工作
                if hasattr(self, 'is_paused') and self.is_paused:
                    self._log_info("[重试] 检测到暂停状态，等待恢复后重试", "orange")
                    # 等待暂停状态解除
                    while hasattr(self, 'is_paused') and self.is_paused and self.is_running:
                        time.sleep(0.1)  # 更频繁检查以提高响应性
                        self.root.update()
                    if not self.is_running:
                        break
                    self._log_info("[重试] 暂停状态已解除，准备重试", "green")
                
                # 阶段3修复：重试前检查验证码状态
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info("[重试] 检测到验证码状态，等待验证码消失后重试", "red")
                    # 等待验证码状态解除
                    while hasattr(self, 'force_stop_flag') and self.force_stop_flag and self.is_running:
                        time.sleep(0.1)  # 更频繁检查以提高响应性
                        self.root.update()
                    if not self.is_running:
                        break
                    self._log_info("[重试] 验证码已消失，准备重试", "green")
                
                # 应用重试延迟
                if hasattr(self, 'retry_manager'):
                    retry_delay = self.retry_manager.get_retry_delay()
                    self._log_info(f"[重试] 等待 {retry_delay} 秒后重试", "blue")
                    time.sleep(retry_delay)
                
                self.retry_current_order = False  # 重置重试标志
                self.current_order_index = i  # 保存当前订单索引
                self._log_info(f"[重试] 重新处理第 {i}/{num_items} 个订单", "orange")
            else:
                self._log_info(f"[循环] 正在处理第 {i}/{num_items} 个订单", "blue")
                
            order_data = {}
            
            # 处理当前订单的所有操作
            for op in actions_to_loop:
                # 首先检查是否已终止操作
                if not self.is_running:
                    self._log_info("操作已终止，停止处理当前订单", "orange")
                    break
                    
                # 检查暂停状态
                if hasattr(self, 'is_paused') and self.is_paused:
                    self._log_info("操作已暂停，等待继续...", "orange")
                    # 等待暂停状态解除
                    while hasattr(self, 'is_paused') and self.is_paused and self.is_running:
                        time.sleep(0.5)
                        self.root.update()
                    if not self.is_running:
                        break
                    self._log_info("暂停状态已解除，继续执行操作", "green")
                
                # 检查验证码检测状态
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info("检测到验证码，暂停操作执行", "red")
                    # 等待验证码消失
                    while hasattr(self, 'force_stop_flag') and self.force_stop_flag and self.is_running:
                        time.sleep(0.5)
                        self.root.update()
                    if not self.is_running:
                        break
                    self._log_info("验证码已消失，继续执行操作", "green")
                
                # 移除'查看2'的特殊逻辑，现在使用与其他元素相同的定位点击方法
                # 原特殊逻辑已被移除，'查看2'现在将通过正常的元素定位和点击流程处理
                
                # 每个操作前都检测验证码
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info(f"操作前检测到验证码，暂停执行: {op['name']}", "red")
                    # 等待验证码消失
                    while hasattr(self, 'force_stop_flag') and self.force_stop_flag and self.is_running:
                        time.sleep(0.1)
                        self.root.update()
                    if not self.is_running:
                        break
                    self._log_info(f"验证码已消失，继续执行操作: {op['name']}", "green")
                
                # 正常处理其他元素
                op_xpath = op.get('xpath', '')
                op_item_xpath = self._generate_xpath_for_item(op_xpath, i, xpath_pattern)
                op_copy = op.copy()
                op_copy['xpath'] = op_item_xpath
                try:
                    result = self._execute_operation(op_copy)
                    if result is not None:
                        order_data[op['name']] = result
                    # 如果"点击前确认"未勾选，每个操作后添加延迟
                    if not self.confirm_click.get():
                        time.sleep(self.auto_action_interval)
                        self._log_info(f"自动执行模式：操作间延迟 {self.auto_action_interval} 秒", "blue")
                except Exception as e:
                    self._log_info(f"执行'{op['name']}'操作失败: {str(e)}", "red")
            
            # 检查是否成功采集了订单数据
            if order_data:
                # 检查是否是重复订单
                current_order_id = None
                for key in ['订单编号', '订单ID', 'order_id', 'orderid']:
                    if key in order_data:
                        current_order_id = order_data[key]
                        break
                
                if current_order_id:
                    if current_order_id in processed_order_ids:
                        consecutive_same_order += 1
                        self._log_info(f"检测到重复订单: {current_order_id}，这是第{consecutive_same_order}次重复", "orange")
                        
                        # 如果连续3次重复，尝试不同的滚动策略
                        if consecutive_same_order >= 3:
                            self._log_info("连续多次重复订单，尝试不同的滚动策略", "red")
                            # 检查是否已终止操作
                            if not self.is_running:
                                self._log_info("操作已终止，停止处理重复订单", "orange")
                                break
                            if not self._scroll_to_next_order():
                                self._log_info("所有滚动策略都失败，无法继续处理", "red")
                                break
                            consecutive_same_order = 0  # 重置计数器
                            continue  # 跳过当前订单，重新尝试
                    else:
                        processed_order_ids.add(current_order_id)
                        consecutive_same_order = 0  # 重置计数器
                        self.collected_data.append(order_data)
                else:
                    # 没有找到订单ID，但仍然添加数据
                    self.collected_data.append(order_data)
            
            # 更新进度条
            self.progress_bar["value"] = i
            self.progress_label.config(text=f"{i}/{num_items}")
            self.root.update()
            
            # 如果不是最后一个订单，滚动到下一个
            if i < num_items:
                # 检查是否已终止操作
                if not self.is_running:
                    self._log_info("操作已终止，停止滚动到下一个订单", "orange")
                    break
                    
                if not self._scroll_to_next_order():
                    self._log_info("无法滚动到下一个订单，尝试继续处理", "orange")
                    # 即使滚动失败，也尝试继续处理
                
                # 等待页面加载前再次检查状态
                if not self.is_running:
                    self._log_info("操作已终止，停止等待页面加载", "orange")
                    break
                time.sleep(1.5)
        
        self._log_info(f"[循环] 已完成所有 {len(self.collected_data)} 个订单的处理", "green")
        self._stop_collection()
        if len(self.collected_data) > 0:
            self.excel_button.config(state=tk.NORMAL)
            self.word_button.config(state=tk.NORMAL)


    def _execute_operation(self, operation):
        """重构：执行单个操作，采用pyautogui移动+WASD微调+剪贴板采集，支持用户验证，对齐代码逻辑.md"""
        import time  # 添加time模块导入，修复UnboundLocalError
        
        # 添加暂停检查 - 阶段1修复：方法开始时检查暂停状态
        if hasattr(self, 'is_paused') and self.is_paused:
            self._log_info(f"操作已暂停，跳过执行: {operation.get('name', '未知操作')}", "orange")
            return None
        
        if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
            self._log_info(f"检测到验证码，跳过执行: {operation.get('name', '未知操作')}", "red")
            return None
        
        # 检查pyautogui依赖
        if 'pyautogui' not in globals():
            self._log_info("pyautogui模块未正确导入，请安装: pip install pyautogui", "red")
            return None
            
        try:
            if not self.driver:
                self._log_info("浏览器未连接，无法执行操作", "red")
                return None
            xpath = operation.get("smart_xpath") or operation["xpath"]
            action = operation["action"]
            name = operation["name"]
            
            # 在元素查找前再次检查暂停状态 - 阶段1修复：元素查找前检查
            if hasattr(self, 'is_paused') and self.is_paused:
                self._log_info(f"元素查找前检测到暂停: {name}", "orange")
                return None
            
            if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                self._log_info(f"元素查找前检测到验证码: {name}", "red")
                return None
            
            # 使用智能定位查找元素
            element = self._find_element_smart(name, xpath)
                
            if not element:
                self._log_info(f"未找到元素: {name}", "red")
                return None
                
            # 获取元素特定的偏移量
            element_offset_x = 0
            element_offset_y = 0
            
            # 尝试从元素上获取名称属性，这是为了确保使用相对XPath等方法找到的元素也能正确应用偏移量
            try:
                element_name = self.driver.execute_script("return arguments[0].getAttribute('data-element-name');", element)
                if element_name and element_name in self.element_offsets:
                    self._log_info(f"使用元素'{element_name}'的偏移量配置", "blue")
                    element_offset_x = self.element_offsets[element_name].get("x", 0)
                    element_offset_y = self.element_offsets[element_name].get("y", 0)
                elif name in self.element_offsets:
                    self._log_info(f"使用元素'{name}'的偏移量配置", "blue")
                    element_offset_x = self.element_offsets[name].get("x", 0)
                    element_offset_y = self.element_offsets[name].get("y", 0)
                else:
                    self._log_info(f"元素'{name}'没有偏移量配置，使用默认值(0,0)", "blue")
            except Exception as e:
                # 如果获取属性失败，回退到使用操作名称
                if name in self.element_offsets:
                    element_offset_x = self.element_offsets[name].get("x", 0)
                    element_offset_y = self.element_offsets[name].get("y", 0)
                    self._log_info(f"从属性获取元素名称失败，使用操作名称'{name}'的偏移量: X={element_offset_x}, Y={element_offset_y}", "orange")
            
            if action == "getText":
                import re
                text = element.text.strip()
                self._log_info(f"获取文本 '{name}': {text}", "green")
                
                # 如果是订单编号元素，解析并保存订单ID
                if name == "订单编号" or "订单编号" in name:
                    match = re.search(r"订单编号[：: ]*([0-9a-zA-Z\-]+)", text)
                    if match:
                        self.last_captured_order_id = match.group(1)
                        self._log_info(f"已保存订单编号: {self.last_captured_order_id}", "blue")
                        # 返回纯订单ID而不是完整文本，确保数据一致性
                        return self.last_captured_order_id
                    else:
                        self.last_captured_order_id = None
                        self._log_info("未能解析订单编号", "red")
                return text
            elif action in ["click", "clickAndGetClipboard"]:
                # 点击前1.2秒延迟
                self._log_info(f"点击前延迟1.2秒: {name}", "blue")
                time.sleep(1.2)
                
                # 在点击前再次检查暂停状态和验证码 - 阶段1修复：点击前检查
                if hasattr(self, 'is_paused') and self.is_paused:
                    self._log_info(f"点击前检测到暂停: {name}", "orange")
                    return None
                
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info(f"点击前检测到验证码: {name}", "red")
                    return None
                
                # 检查是否为虚拟元素（使用缓存坐标）
                if hasattr(element, 'is_virtual') and element.is_virtual:
                    # 使用缓存坐标直接点击
                    cached_coords = element.cached_coords
                    # 修复坐标字段名称
                    if 'screen_x' in cached_coords and 'screen_y' in cached_coords:
                        screen_x = cached_coords['screen_x']
                        screen_y = cached_coords['screen_y']
                    elif 'x' in cached_coords and 'y' in cached_coords:
                        screen_x = cached_coords['x']
                        screen_y = cached_coords['y']
                    else:
                        self._log_info(f"虚拟元素'{name}'的坐标格式错误: {cached_coords}", "red")
                        return None
                    self._log_info(f"使用缓存坐标点击元素 '{name}': X={screen_x}, Y={screen_y}", "orange")
                    
                    # 应用元素特定的偏移量
                    if name in self.element_offsets:
                        element_offset_x = self.element_offsets[name].get("x", 0)
                        element_offset_y = self.element_offsets[name].get("y", 0)
                        screen_x += element_offset_x
                        screen_y += element_offset_y
                        self._log_info(f"应用偏移量后的坐标: X={screen_x}, Y={screen_y}", "orange")
                    
                    # 执行点击
                    self._switch_focus_to_browser()
                    time.sleep(0.3)
                    pyautogui.moveTo(int(screen_x), int(screen_y))
                    pyautogui.click()
                    self._log_info(f"已使用缓存坐标点击 '{name}'", "green")
                    
                    # 更新缓存坐标
                    try:
                        self.coordinate_cache.save_coordinate(name, screen_x, screen_y)
                        self._log_info(f"已更新缓存坐标 '{name}': ({screen_x}, {screen_y})", "blue")
                    except Exception as e:
                        self._log_info(f"更新缓存坐标失败 '{name}': {str(e)}", "orange")
                    
                    self._manage_focus()
                    
                    # 虚拟元素点击成功，返回True
                    if action == "click":
                        return True
                    
                else:
                    # 正常元素处理流程
                    # 滚动到元素可见
                    self.driver.execute_script('arguments[0].scrollIntoView({behavior: "smooth", block: "center"});', element)
                    time.sleep(0.5)
                    
                    # 获取浏览器窗口位置和尺寸
                    window_pos = self.driver.get_window_position()
                    window_size = self.driver.get_window_size()
                    
                    # 获取元素在视口中的位置和尺寸
                    rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
                
                    # 记录元素的原始位置信息
                    self._log_info(f"元素'{name}'的原始位置: left={rect['left']}, top={rect['top']}, width={rect['width']}, height={rect['height']}", "blue")
                    
                    # 获取浏览器内容区域的偏移量
                    js = """
                    return {
                        screenX: window.screenX, 
                        screenY: window.screenY, 
                        outerHeight: window.outerHeight, 
                        innerHeight: window.innerHeight,
                        outerWidth: window.outerWidth,
                        innerWidth: window.innerWidth
                    };
                    """
                    win_metrics = self.driver.execute_script(js)
                    
                    # 计算内容区域的左上角位置
                    content_left = win_metrics['screenX'] + (win_metrics['outerWidth'] - win_metrics['innerWidth']) / 2
                    content_top = win_metrics['screenY'] + (win_metrics['outerHeight'] - win_metrics['innerHeight'])
                    
                    self._log_info(f"浏览器内容区域: left={content_left}, top={content_top}", "blue")
                    
                    # 计算元素的绝对屏幕位置（取元素中心点）
                    screen_x = content_left + rect['left'] + rect['width'] / 2
                    screen_y = content_top + rect['top'] + rect['height'] / 2
                    
                    self._log_info(f"元素'{name}'的中心点位置: X={screen_x}, Y={screen_y}", "blue")
                
                    # 检查元素位置是否合理
                    if rect['left'] < 0 or rect['top'] < 0 or rect['width'] <= 0 or rect['height'] <= 0:
                        self._log_info(f"警告: 元素'{name}'的位置或尺寸不合理，尝试重新获取", "orange")
                        
                        # 尝试使用JavaScript重新获取元素位置
                        try:
                            js_rect = """
                            var element = arguments[0];
                            var rect = element.getBoundingClientRect();
                            
                            // 确保元素可见
                            element.scrollIntoView({behavior: 'auto', block: 'center'});
                            
                            // 等待一下让滚动完成
                            return new Promise(resolve => {
                                setTimeout(() => {
                                    var updatedRect = element.getBoundingClientRect();
                                    resolve({
                                        left: updatedRect.left,
                                        top: updatedRect.top,
                                        width: updatedRect.width,
                                        height: updatedRect.height
                                    });
                                }, 500);
                            });
                            """
                            
                            updated_rect = self.driver.execute_script(js_rect, element)
                            self._log_info(f"重新获取元素'{name}'的位置: left={updated_rect['left']}, top={updated_rect['top']}, width={updated_rect['width']}, height={updated_rect['height']}", "blue")
                            
                            # 使用更新后的位置
                            if updated_rect['width'] > 0 and updated_rect['height'] > 0:
                                rect = updated_rect
                                # 重新计算元素的绝对屏幕位置
                                screen_x = content_left + rect['left'] + rect['width'] / 2
                                screen_y = content_top + rect['top'] + rect['height'] / 2
                                self._log_info(f"更新后元素'{name}'的中心点位置: X={screen_x}, Y={screen_y}", "blue")
                        except Exception as e:
                            self._log_info(f"重新获取元素位置失败: {str(e)}", "red")
                    
                    # 应用元素特定的偏移量
                    screen_x += element_offset_x
                    screen_y += element_offset_y
                    
                    # 记录应用的偏移量
                    self._log_info(f"应用元素'{name}'的偏移量: X={element_offset_x}, Y={element_offset_y}", "blue")
                    self._log_info(f"最终点击位置: X={screen_x}, Y={screen_y}", "blue")
                    
                    # 检查点击位置是否在屏幕范围内
                    screen_width, screen_height = pyautogui.size()
                    if screen_x < 0 or screen_x > screen_width or screen_y < 0 or screen_y > screen_height:
                        self._log_info(f"警告: 计算的点击位置({screen_x}, {screen_y})超出屏幕范围({screen_width}x{screen_height})", "red")
                        # 调整到屏幕范围内
                        screen_x = max(0, min(screen_x, screen_width))
                        screen_y = max(0, min(screen_y, screen_height))
                        self._log_info(f"调整后的点击位置: X={screen_x}, Y={screen_y}", "orange")
                
                    # 支持WASD微调
                    if self.confirm_click.get():
                        ok = self._show_verification_dialog_and_wait(name, xpath, screen_x, screen_y, element_offset_x, element_offset_y)
                        if not ok:
                            self._log_info(f"用户取消了点击: {name}", "orange")
                            return None
                        # 在_show_verification_dialog_and_wait中已经执行了点击，直接返回
                        return True
                    else:
                        # 没有确认对话框，直接在应用了元素偏移量的位置点击
                        
                        # 点击前确保浏览器窗口有焦点
                        self._switch_focus_to_browser()
                        time.sleep(0.3)
                        
                        # 记录移动前的鼠标位置
                        current_pos = pyautogui.position()
                        self._log_info(f"[坐标日志] 元素'{name}' - 移动前鼠标位置: X={current_pos.x}, Y={current_pos.y}", "cyan")
                        
                        # 移动到目标位置
                        pyautogui.moveTo(int(screen_x), int(screen_y))
                        
                        # 记录移动后的鼠标位置
                        moved_pos = pyautogui.position()
                        self._log_info(f"[坐标日志] 元素'{name}' - 移动后鼠标位置: X={moved_pos.x}, Y={moved_pos.y}", "cyan")
                        
                        # 执行点击
                        pyautogui.click()
                        
                        # 记录点击时的坐标
                        click_pos = pyautogui.position()
                        self._log_info(f"[坐标日志] 元素'{name}' - 点击时坐标: X={click_pos.x}, Y={click_pos.y}", "green")
                        self._log_info(f"已通过 PyAutoGUI 点击 '{name}'", "blue")
                        
                        # 点击成功后保存坐标到缓存（无论是否为重试模式） - 阶段3增强
                        try:
                            self._save_successful_coordinates_enhanced(name, click_pos.x, click_pos.y, element_offset_x, element_offset_y)
                            self._log_info(f"已缓存元素 '{name}' 的坐标: ({click_pos.x}, {click_pos.y})", "blue")
                        except Exception as e:
                            self._log_info(f"缓存坐标失败 '{name}': {str(e)}", "orange")
                        
                        # 点击后延迟一下，让浏览器有时间响应
                        time.sleep(0.5)
                        
                        # 对于'复制完整的收货信息'元素，跳过额外点击以避免剪贴板内容重复
                        if name != '复制完整的收货信息':
                            # 执行一次额外的原地点击，与点击前确认模式行为一致
                            extra_click_pos = pyautogui.position()
                            self._log_info(f"[坐标日志] 元素'{name}' - 额外点击前坐标: X={extra_click_pos.x}, Y={extra_click_pos.y}", "cyan")
                            
                            pyautogui.click()
                            
                            # 记录额外点击时的坐标
                            extra_click_after_pos = pyautogui.position()
                            self._log_info(f"[坐标日志] 元素'{name}' - 额外点击时坐标: X={extra_click_after_pos.x}, Y={extra_click_after_pos.y}", "green")
                            self._log_info(f"已执行额外的原地点击 '{name}'", "blue")
                            time.sleep(0.3)
                        else:
                            self._log_info(f"跳过'{name}'的额外点击，避免剪贴板内容重复", "blue")
                        
                        # 恢复焦点到采集工具窗口
                        self._manage_focus()
                        
                        # 点击操作成功，返回True
                        if action == "click":
                            return True
                    
                # 统一处理clickAndGetClipboard动作
                if action == "clickAndGetClipboard":
                     # 对于'复制完整的收货信息'元素，如果跳过了额外点击，直接获取剪贴板内容
                     if name == '复制完整的收货信息':
                         self._log_info(f"直接获取剪贴板内容，无需等待更新", "blue")
                         time.sleep(0.5)
                         self._manage_focus()
                         clipboard_content = pyperclip.paste()
                     else:
                         self._log_info(f"点击后等待剪贴板内容更新...", "blue")
                         time.sleep(1.5)
                         self._manage_focus()
                         time.sleep(0.5)
                         clipboard_content = self._wait_for_clipboard_content(
                             timeout=12.0,
                             check_interval=0.5,
                             min_length=10
                         )
                     # 使用之前采集的订单ID
                     current_order_id = getattr(self, 'last_captured_order_id', None)
                     if not current_order_id:
                         # 弹窗要求用户输入订单ID
                         from tkinter import simpledialog
                         current_order_id = simpledialog.askstring(
                             "订单ID缺失", "未能自动提取订单ID，请手动输入当前订单ID：", parent=self.root)
                         if not current_order_id or not current_order_id.strip():
                             self._log_info("用户未输入订单ID，跳过本次映射", "red")
                             return clipboard_content
                     if clipboard_content and clipboard_content.strip():
                         self._store_clipboard_content(name, clipboard_content, current_order_id)
                         self._log_info(f"[映射写入] 订单ID: {current_order_id}, 长度: {len(clipboard_content)}, 内容: {clipboard_content[:30]}...", "green")
                     else:
                         self._log_info(f"[映射跳过] 订单ID: {current_order_id}, 内容无效或为空", "orange")
                     return clipboard_content
                    
        except Exception as e:
            self._log_info(f"执行'{name}'操作失败: {str(e)}", "red")
            import traceback
            self._log_info(traceback.format_exc(), "red")
            return None


    def _find_element_smart(self, name, original_xpath):
        """智能元素查找，使用多种策略定位元素 - 阶段3增强：配置化重试策略"""
        if not self.driver:
            self._log_info(f"浏览器未连接，无法查找元素 '{name}'", "red")
            return None
            
        self._log_info(f"智能查找元素: '{name}'", "blue")
        element = None
        
        # 阶段3增强：获取重试策略列表
        retry_strategies = ["smart_element_search", "fallback_xpath"]
        if hasattr(self, 'retry_manager'):
            retry_strategies = self.retry_manager.get_retry_strategies()
            # 注意：重试模式下不直接使用缓存坐标，而是先尝试正常的元素查找方法
            # 缓存坐标只作为所有策略都失败时的最后备选方案
        
        # 策略1: 使用原始XPath
        try:
            element = self.driver.find_element(By.XPATH, original_xpath)
            self._log_info(f"使用原始XPath找到元素 '{name}'", "green")
            # 给元素添加一个属性，标记它的名称，用于后续应用偏移量
            self.driver.execute_script("arguments[0].setAttribute('data-element-name', arguments[1]);", element, name)
            
            # 检查元素位置
            rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
            if rect['width'] <= 0 or rect['height'] <= 0:
                self._log_info(f"警告: 使用原始XPath找到的元素'{name}'尺寸异常: width={rect['width']}, height={rect['height']}", "orange")
                # 尝试滚动到元素
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", element)
                time.sleep(0.5)
            
            return element
        except Exception as e:
            self._log_info(f"原始XPath未找到元素 '{name}': {str(e)}", "orange")
        
        # 策略2: 使用相对XPath
        try:
            # 尝试生成更健壮的相对XPath
            relative_xpath = self._generate_relative_xpath(original_xpath)
            if relative_xpath:
                element = self.driver.find_element(By.XPATH, relative_xpath)
                self._log_info(f"使用相对XPath找到元素 '{name}'", "green")
                
                # 检查元素位置并确保元素可见
                rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
                self._log_info(f"相对XPath找到的元素'{name}'位置: left={rect['left']}, top={rect['top']}, width={rect['width']}, height={rect['height']}", "blue")
                
                # 滚动到相对XPath找到的元素位置，使目标区域可见
                self._log_info(f"滚动到相对XPath找到的元素位置，尝试让原始XPath重新生效", "blue")
                self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", element)
                time.sleep(0.8)  # 等待滚动完成
                
                # 滚动后重新尝试使用原始XPath查找
                try:
                    original_element = self.driver.find_element(By.XPATH, original_xpath)
                    original_rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', original_element)
                    self._log_info(f"滚动后原始XPath重新找到元素'{name}': left={original_rect['left']}, top={original_rect['top']}, width={original_rect['width']}, height={original_rect['height']}", "green")
                    
                    # 检查原始元素位置是否正常
                    if original_rect['width'] > 0 and original_rect['height'] > 0:
                        # 给元素添加一个属性，标记它的名称，用于后续应用偏移量
                        self.driver.execute_script("arguments[0].setAttribute('data-element-name', arguments[1]);", original_element, name)
                        self._log_info(f"使用滚动后重新找到的原始XPath元素 '{name}'", "green")
                        return original_element
                    else:
                        self._log_info(f"滚动后原始XPath元素位置仍异常，继续使用相对XPath元素", "orange")
                except Exception as e:
                    self._log_info(f"滚动后原始XPath仍未找到元素 '{name}': {str(e)}", "orange")
                
                # 如果原始XPath仍然失败，检查相对XPath元素的位置合理性
                if rect['width'] <= 0 or rect['height'] <= 0 or rect['left'] < 0 or rect['top'] < 0:
                    self._log_info(f"警告: 相对XPath找到的元素'{name}'位置异常，可能不是目标元素", "orange")
                    # 如果位置明显异常（如left<100, top<100），很可能找错了元素
                    if rect['left'] < 100 and rect['top'] < 100:
                        self._log_info(f"相对XPath找到的元素位置过于靠近页面左上角，可能是错误元素，跳过此策略", "red")
                        raise Exception("相对XPath找到错误元素")
                
                # 给元素添加一个属性，标记它的名称，用于后续应用偏移量
                self.driver.execute_script("arguments[0].setAttribute('data-element-name', arguments[1]);", element, name)
                return element
        except Exception as e:
            self._log_info(f"相对XPath未找到元素 '{name}': {str(e)}", "orange")
        
        # 策略3: 使用文本内容查找
        try:
            element = self._find_by_text_content(name)
            if element:
                self._log_info(f"使用文本内容找到元素 '{name}'", "green")
                # 给元素添加一个属性，标记它的名称，用于后续应用偏移量
                self.driver.execute_script("arguments[0].setAttribute('data-element-name', arguments[1]);", element, name)
                
                # 检查元素位置
                rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
                if rect['width'] <= 0 or rect['height'] <= 0:
                    self._log_info(f"警告: 使用文本内容找到的元素'{name}'尺寸异常: width={rect['width']}, height={rect['height']}", "orange")
                    # 尝试滚动到元素
                    self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", element)
                    time.sleep(0.5)
                
                return element
        except Exception as e:
            self._log_info(f"文本内容未找到元素 '{name}': {str(e)}", "orange")
        
        # 策略4: 使用CSS选择器
        try:
            # 尝试从XPath转换为CSS选择器
            css_selector = self._xpath_to_css(original_xpath)
            if css_selector:
                element = self.driver.find_element(By.CSS_SELECTOR, css_selector)
                self._log_info(f"使用CSS选择器找到元素 '{name}'", "green")
                # 给元素添加一个属性，标记它的名称，用于后续应用偏移量
                self.driver.execute_script("arguments[0].setAttribute('data-element-name', arguments[1]);", element, name)
                
                # 检查元素位置
                rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
                if rect['width'] <= 0 or rect['height'] <= 0:
                    self._log_info(f"警告: 使用CSS选择器找到的元素'{name}'尺寸异常: width={rect['width']}, height={rect['height']}", "orange")
                    # 尝试滚动到元素
                    self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", element)
                    time.sleep(0.5)
                
                return element
        except Exception as e:
            self._log_info(f"CSS选择器未找到元素 '{name}': {str(e)}", "orange")
        
        # 策略5: 使用JavaScript查找
        try:
            js_script = f"""
            function findElementByContent(text) {{
                const elements = Array.from(document.querySelectorAll('*'));
                return elements.find(el => {{
                    const content = el.textContent || '';
                    return content.includes(text) && content.length < 100;
                }});
            }}
            
            const element = findElementByContent('{name}');
            if (element) {{
                element.style.border = '2px solid red';
                element.scrollIntoView({{behavior: 'smooth', block: 'center'}});
                element.setAttribute('data-element-name', '{name}');
                return true;
            }}
            return false;
            """
            
            found = self.driver.execute_script(js_script)
            if found:
                # 使用JavaScript找到元素后，我们需要重新定位它以返回WebElement对象
                # 这里使用一个简单的策略：查找带有红色边框的元素
                element = self.driver.find_element(By.XPATH, "//*[contains(@style, 'border: 2px solid red')]")
                self._log_info(f"使用JavaScript找到元素 '{name}'", "green")
                
                # 检查元素位置
                rect = self.driver.execute_script('return arguments[0].getBoundingClientRect();', element)
                if rect['width'] <= 0 or rect['height'] <= 0:
                    self._log_info(f"警告: 使用JavaScript找到的元素'{name}'尺寸异常: width={rect['width']}, height={rect['height']}", "orange")
                    # 尝试滚动到元素
                    self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'auto', block: 'center'});", element)
                    time.sleep(0.5)
                
                return element
        except Exception as e:
            self._log_info(f"JavaScript未找到元素 '{name}': {str(e)}", "orange")
        
        # 策略6: 在重试模式下使用缓存坐标（仅当所有其他策略都失败时）
        # 只有在明确的重试模式下才使用缓存坐标
        if (hasattr(self, 'retry_current_order') and self.retry_current_order and 
            hasattr(self, 'retry_manager') and self.retry_manager.is_coordinate_cache_enabled()):
            try:
                cached_coords = self._load_cached_coordinates(name)
                if cached_coords and self.retry_manager.is_coordinate_valid(cached_coords):
                    self._log_info(f"[重试模式] 所有策略失败，尝试使用缓存坐标查找元素 '{name}'", "orange")
                    # 创建虚拟元素对象，使用增强的缓存坐标信息
                    virtual_element = self._create_virtual_element_enhanced(name, cached_coords)
                    if virtual_element:
                        self._log_info(f"[重试模式] 使用缓存坐标成功创建虚拟元素 '{name}'", "green")
                        return virtual_element
            except Exception as e:
                self._log_info(f"[重试模式] 使用缓存坐标失败 '{name}': {str(e)}", "orange")
        
        self._log_info(f"所有策略都未能找到元素 '{name}'", "red")
        return None
    
    def _load_cached_coordinates(self, element_name):
        """加载缓存的坐标信息 - 阶段3增强方法"""
        try:
            cache_file = "coordinate_cache.json"
            if not os.path.exists(cache_file):
                return None
            
            with open(cache_file, 'r', encoding='utf-8') as f:
                cache_data = json.load(f)
            
            coords = cache_data.get('coordinates', {}).get(element_name)
            if coords and coords.get('success_count', 0) > 0:
                return coords
            
            return None
            
        except Exception as e:
            self._log_info(f"加载坐标缓存失败: {e}", "error")
            return None
    
    def _create_virtual_element_enhanced(self, name, cached_coords):
        """基于缓存坐标创建增强虚拟元素 - 阶段3增强方法"""
        try:
            # 验证缓存坐标的完整性
            required_fields = ['screen_x', 'screen_y', 'element_offset_x', 'element_offset_y']
            for field in required_fields:
                if field not in cached_coords:
                    self._log_info(f"缓存坐标缺少必要字段: {field}", "red")
                    return None
            
            # 获取当前页面滚动位置，动态调整坐标
            try:
                current_scroll_y = self.driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop;")
                cached_scroll_y = cached_coords.get('scroll_y', 0)  # 缓存时的滚动位置
                scroll_offset = current_scroll_y - cached_scroll_y
                
                # 调整Y坐标
                adjusted_screen_y = cached_coords['screen_y'] - scroll_offset
                
                self._log_info(f"[坐标调整] 元素'{name}' - 当前滚动: {current_scroll_y}, 缓存滚动: {cached_scroll_y}, 偏移: {scroll_offset}", "cyan")
                self._log_info(f"[坐标调整] 元素'{name}' - 原始Y: {cached_coords['screen_y']}, 调整后Y: {adjusted_screen_y}", "cyan")
            except Exception as e:
                self._log_info(f"获取滚动位置失败，使用原始坐标: {e}", "orange")
                adjusted_screen_y = cached_coords['screen_y']
            
            # 创建增强的虚拟元素类
            class EnhancedVirtualElement:
                def __init__(self, name, coords, adjusted_y):
                    self.name = name
                    self.cached_coords = coords
                    self.is_virtual = True
                    self.screen_x = coords['screen_x']
                    self.screen_y = adjusted_y  # 使用调整后的Y坐标
                    self.element_offset_x = coords['element_offset_x']
                    self.element_offset_y = coords['element_offset_y']
                    self.success_count = coords.get('success_count', 0)
                    self.last_success = coords.get('last_success', '')
                
                def get_attribute(self, attr):
                    if attr == 'data-element-name':
                        return self.name
                    return None
                
                def get_rect(self):
                    """返回虚拟元素的位置信息"""
                    return {
                        'x': self.screen_x,
                        'y': self.screen_y,
                        'width': 50,  # 默认宽度
                        'height': 20   # 默认高度
                    }
            
            virtual_element = EnhancedVirtualElement(name, cached_coords, adjusted_screen_y)
            self._log_info(f"成功创建增强虚拟元素 '{name}' 使用坐标: ({cached_coords['screen_x']}, {adjusted_screen_y})", "green")
            
            # 记录缓存坐标使用
            if hasattr(self, 'retry_manager'):
                from utils import create_retry_log_entry, write_retry_log
                log_entry = create_retry_log_entry("cached_coordinates_used", name, {
                    "coordinates": cached_coords,
                    "success_count": cached_coords.get('success_count', 0)
                })
                write_retry_log(log_entry)
            
            return virtual_element
            
        except Exception as e:
            self._log_info(f"创建增强虚拟元素失败 '{name}': {str(e)}", "red")
            return None
    
    def _save_successful_coordinates_enhanced(self, element_name, screen_x, screen_y, offset_x, offset_y):
        """保存成功的点击坐标 - 阶段3增强方法"""
        try:
            cache_file = "coordinate_cache.json"
            cache_data = {}
            
            # 读取现有缓存
            if os.path.exists(cache_file):
                with open(cache_file, 'r', encoding='utf-8') as f:
                    cache_data = json.load(f)
            
            # 初始化缓存结构
            if 'coordinates' not in cache_data:
                cache_data['coordinates'] = {}
            
            # 获取当前页面滚动位置
            try:
                current_scroll_y = self.driver.execute_script("return window.pageYOffset || document.documentElement.scrollTop;")
            except Exception as e:
                self._log_info(f"获取滚动位置失败: {e}", "orange")
                current_scroll_y = 0
            
            # 更新坐标信息
            if element_name not in cache_data['coordinates']:
                cache_data['coordinates'][element_name] = {'success_count': 0}
            
            coord_info = cache_data['coordinates'][element_name]
            coord_info.update({
                'screen_x': screen_x,
                'screen_y': screen_y,
                'element_offset_x': offset_x,
                'element_offset_y': offset_y,
                'scroll_y': current_scroll_y,  # 记录保存时的滚动位置
                'success_count': coord_info.get('success_count', 0) + 1,
                'last_success': datetime.now().isoformat()
            })
            
            self._log_info(f"[缓存] 保存坐标时滚动位置: {current_scroll_y}", "cyan")
            
            cache_data['cache_version'] = '1.0'
            cache_data['last_updated'] = datetime.now().isoformat()
            
            # 保存缓存
            with open(cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
            
            self._log_info(f"[缓存] 已保存元素'{element_name}'的成功坐标", "blue")
            
            # 记录坐标保存事件
            if hasattr(self, 'retry_manager') and self.retry_manager.is_retry_logging_enabled():
                from utils import create_retry_log_entry, write_retry_log
                log_entry = create_retry_log_entry("coordinates_saved", element_name, {
                    "screen_x": screen_x,
                    "screen_y": screen_y,
                    "offset_x": offset_x,
                    "offset_y": offset_y,
                    "success_count": coord_info['success_count']
                })
                write_retry_log(log_entry)
            
        except Exception as e:
            self._log_info(f"保存坐标缓存失败: {e}", "error")
    
    def _create_virtual_element(self, name, cached_coords):
        """创建虚拟元素对象，用于处理缓存坐标"""
        try:
            # 简单验证缓存坐标的有效性
            has_screen_coords = cached_coords and 'screen_x' in cached_coords and 'screen_y' in cached_coords
            has_xy_coords = cached_coords and 'x' in cached_coords and 'y' in cached_coords
            
            if not cached_coords or (not has_screen_coords and not has_xy_coords):
                self._log_info(f"缓存坐标无效: {cached_coords}", "red")
                return None
            
            # 创建一个简单的虚拟元素类
            class VirtualElement:
                def __init__(self, name, coords):
                    self.name = name
                    self.cached_coords = coords
                    self.is_virtual = True
                
                def get_attribute(self, attr):
                    if attr == 'data-element-name':
                        return self.name
                    return None
            
            virtual_element = VirtualElement(name, cached_coords)
            self._log_info(f"成功创建虚拟元素 '{name}' 使用坐标: {cached_coords}", "green")
            return virtual_element
            
        except Exception as e:
            self._log_info(f"创建虚拟元素失败 '{name}': {str(e)}", "red")
            return None

    def _get_order_count(self):
        """从页面获取待处理订单数量"""
        if not self.driver:
            return 0
            
        # 如果操作序列中包含了订单数量元素，使用它获取数量
        order_count_elements = [op for op in self.operation_sequence if op.get("is_order_count", False)]
        
        if order_count_elements:
            try:
                # 使用指定的XPath获取订单数量元素
                op = order_count_elements[0]
                xpath = op["xpath"]
                element = self.driver.find_element(By.XPATH, xpath)
                text = element.text
                
                # 提取数字
                import re
                numbers = re.findall(r'\d+', text)
                if numbers:
                    count = int(numbers[0])
                    self._log_info(f"自动获取订单数量: {count}", "green")
                    return count
            except Exception as e:
                self._log_info(f"获取订单数量失败: {str(e)}", "red")
        
        # 如果没有指定订单数量元素或获取失败，询问用户
        count = simpledialog.askinteger(
            "输入订单数量", 
            "请手动输入本次要处理的订单数量:", 
            minvalue=1, 
            maxvalue=1000
        )
        if count:
            self._log_info(f"手动输入订单数量: {count}", "blue")
            return count
        
        # 默认值
        default_count = 5
        self._log_info(f"使用默认订单数量: {default_count}", "orange")
        return default_count


    def _extract_current_order_id(self):
        """提取当前页面上的订单ID"""
        # 检查是否已终止操作
        if not self.is_running:
            self._log_info("操作已终止，停止提取订单ID", "orange")
            return None
            
        import re  # 导入re模块解决"cannot access local variable 're'"问题
        try:
            # 策略1: 尝试从URL中提取
            url = self.driver.current_url
            url_match = re.search(r'order[_=-]?id=([A-Za-z0-9-]+)', url, re.IGNORECASE)
            if url_match:
                order_id = url_match.group(1)
                self._log_info(f"从URL成功提取订单ID: {order_id}", "green")
                print(f"DEBUG-URL-ID-EXTRACTED: '{order_id}'")
                
                # 保存到最近使用的ID列表
                if order_id not in self.last_order_ids:
                    self.last_order_ids.append(order_id)
                    if len(self.last_order_ids) > 5:
                        self.last_order_ids.pop(0)
                
                # 更新当前订单ID
                self.current_order_id = order_id
                
                return order_id
            
            # 策略2: 尝试多种可能的订单ID选择器
            order_id_selectors = [
                "//span[contains(text(), '订单编号')]/following-sibling::span",
                "//span[contains(text(), '订单编号')]",
                "//div[contains(text(), '订单编号')]",
                "//td[contains(text(), '订单编号')]",
                "//th[contains(text(), '订单编号')]/following-sibling::td",
                "//label[contains(text(), '订单编号')]/following-sibling::*"
            ]
            
            self._log_info("开始提取当前可见订单ID...", "blue")
            print("DEBUG-EXTRACT-ID: 开始提取当前可见订单ID")
            
            # 添加重试机制
            max_retries = 3
            for retry in range(max_retries):
                if retry > 0:
                    self._log_info(f"第{retry+1}次尝试提取订单ID...", "orange")
                    print(f"DEBUG-EXTRACT-RETRY: 第{retry+1}次尝试提取订单ID")
                    # 短暂等待页面可能的更新
                    time.sleep(0.5)
                
                # 获取当前视口内可见元素的信息
                visible_elements_js = """
                return Array.from(document.querySelectorAll('*')).filter(el => {
                    const rect = el.getBoundingClientRect();
                    return rect.top >= 0 && 
                           rect.left >= 0 && 
                           rect.bottom <= window.innerHeight && 
                           rect.right <= window.innerWidth &&
                           el.textContent.includes('订单编号');
                }).map(el => ({
                    text: el.textContent,
                    top: el.getBoundingClientRect().top,
                    visible: true
                }));
                """
                
                try:
                    visible_elements = self.driver.execute_script(visible_elements_js)
                    if visible_elements:
                        print(f"DEBUG-VISIBLE-ELEMENTS: 找到{len(visible_elements)}个可见元素包含'订单编号'")
                        for i, elem in enumerate(visible_elements):
                            print(f"DEBUG-VISIBLE-ELEMENT-{i+1}: {elem['text'][:50]}, top={elem['top']}")
                        
                        # 优先使用可见元素中最靠近视口中心的元素
                        if len(visible_elements) > 0:
                            # 按照元素距离视口顶部的距离排序
                            sorted_elements = sorted(visible_elements, key=lambda e: abs(e['top'] - (self.driver.execute_script("return window.innerHeight") / 2)))
                            center_element_text = sorted_elements[0]['text']
                            print(f"DEBUG-CENTER-ELEMENT: 选择最靠近视口中心的元素: '{center_element_text[:50]}'")
                            
                            # 从文本中提取订单ID
                            match = re.search(r'订单编号[：:]\s*([0-9a-zA-Z-]+)', center_element_text)
                            if match:
                                extracted_id = match.group(1)
                                self._log_info(f"从可见元素中提取到订单ID: '{extracted_id}'", "green")
                                print(f"DEBUG-VISIBLE-ID-EXTRACTED: '{extracted_id}'")
                                
                                # 保存到最近使用的ID列表
                                if extracted_id not in self.last_order_ids:
                                    self.last_order_ids.append(extracted_id)
                                    if len(self.last_order_ids) > 5:
                                        self.last_order_ids.pop(0)
                                
                                # 更新当前订单ID
                                self.current_order_id = extracted_id
                                
                                # 验证提取的ID格式是否合理
                                if len(extracted_id) >= 5:  # 订单ID通常较长
                                    return extracted_id
                                else:
                                    self._log_info(f"提取的订单ID '{extracted_id}' 格式可疑，尝试其他方法", "orange")
                except Exception as js_e:
                    print(f"DEBUG-JS-ERROR: 获取可见元素失败: {str(js_e)}")
                
                for selector in order_id_selectors:
                    try:
                        elements = self.driver.find_elements(By.XPATH, selector)
                        if elements:
                            order_id = elements[0].text
                            self._log_info(f"找到订单ID元素，原始文本: '{order_id}'", "blue")
                            print(f"DEBUG-ORDER-ID-ELEMENT: '{order_id}'")
                            
                            # 提取数字和字母部分作为订单ID
                            import re
                            match = re.search(r'[0-9a-zA-Z-]+', order_id)
                            if match:
                                extracted_id = match.group(0)
                                self._log_info(f"提取到订单ID: '{extracted_id}'", "green")
                                print(f"DEBUG-ORDER-ID-EXTRACTED: '{extracted_id}'")
                                
                                # 保存到最近使用的ID列表
                                if extracted_id not in self.last_order_ids:
                                    self.last_order_ids.append(extracted_id)
                                    if len(self.last_order_ids) > 5:
                                        self.last_order_ids.pop(0)
                                
                                # 更新当前订单ID
                                self.current_order_id = extracted_id
                                
                                # 验证提取的ID格式是否合理
                                if len(extracted_id) >= 5:  # 订单ID通常较长
                                    return extracted_id
                                else:
                                    self._log_info(f"提取的订单ID '{extracted_id}' 格式可疑，尝试下一个选择器", "orange")
                                    continue
                            
                            cleaned_id = order_id.strip()
                            self._log_info(f"清理后的订单ID: '{cleaned_id}'", "green")
                            print(f"DEBUG-ORDER-ID-CLEANED: '{cleaned_id}'")
                            
                            # 保存到最近使用的ID列表
                            if cleaned_id not in self.last_order_ids:
                                self.last_order_ids.append(cleaned_id)
                                if len(self.last_order_ids) > 5:
                                    self.last_order_ids.pop(0)
                            
                            # 更新当前订单ID
                            self.current_order_id = cleaned_id
                            
                            # 验证清理后的ID格式是否合理
                            if len(cleaned_id) >= 5:  # 订单ID通常较长
                                return cleaned_id
                            else:
                                self._log_info(f"清理后的订单ID '{cleaned_id}' 太短，尝试下一个选择器", "orange")
                    except Exception as e:
                        print(f"DEBUG-SELECTOR-ERROR: 选择器 '{selector}' 失败: {str(e)}")
            
            # 策略3: 尝试从页面源代码中查找订单ID模式
            page_source = self.driver.page_source
            patterns = [
                r'订单编号[：:]\s*([0-9a-zA-Z-]{5,})',
                r'订单号[：:]\s*([0-9a-zA-Z-]{5,})',
                r'订单[：:]\s*([0-9a-zA-Z-]{5,})',
                r'单号[：:]\s*([0-9a-zA-Z-]{5,})',
                r'order[_\s]?id[：:=]\s*([0-9a-zA-Z-]{5,})',
                r'order[_\s]?number[：:=]\s*([0-9a-zA-Z-]{5,})'
            ]
            
            for pattern in patterns:
                match = re.search(pattern, page_source, re.IGNORECASE)
                if match:
                    extracted_id = match.group(1)
                    self._log_info(f"从页面源码提取订单ID: '{extracted_id}'", "green")
                    print(f"DEBUG-SOURCE-ID-EXTRACTED: '{extracted_id}'")
                    
                    # 保存到最近使用的ID列表
                    if extracted_id not in self.last_order_ids:
                        self.last_order_ids.append(extracted_id)
                        if len(self.last_order_ids) > 5:
                            self.last_order_ids.pop(0)
                    
                    # 更新当前订单ID
                    self.current_order_id = extracted_id
                    
                    return extracted_id
            
            # 策略4: 如果有当前缓存的订单ID，使用它
            if hasattr(self, 'current_order_id') and self.current_order_id:
                self._log_info(f"使用当前缓存的订单ID: '{self.current_order_id}'", "orange")
                print(f"DEBUG-CACHED-ID-USED: '{self.current_order_id}'")
                return self.current_order_id
            
            self._log_info("无法提取订单ID", "red")
            print("DEBUG-NO-ID-EXTRACTED")
            return None
        except Exception as e:
            self._log_info(f"提取订单ID时发生异常: {str(e)}", "red")
            print(f"DEBUG-EXTRACT-ID-ERROR: {str(e)}")
            
            # 在异常情况下，如果有当前缓存的订单ID，使用它
            if hasattr(self, 'current_order_id') and self.current_order_id:
                self._log_info(f"异常情况下使用当前缓存的订单ID: '{self.current_order_id}'", "orange")
                print(f"DEBUG-ERROR-CACHED-ID-USED: '{self.current_order_id}'")
                return self.current_order_id
                
            return None
    

    def _scroll_to_next_order(self):
        """使用多种滚动策略尝试滚动到下一个订单"""
        # 检查是否已终止操作
        if not self.is_running:
            self._log_info("操作已终止，停止滚动操作", "orange")
            return False
            
        if not self.driver:
            self._log_info("无法滚动：浏览器未连接", "red")
            return False
            
        # 获取当前订单ID，用于后续比较
        current_order_id = self._extract_current_order_id()
        self._log_info(f"滚动前订单ID: {current_order_id}", "blue")
        
        # 保存滚动前的截图（如果启用了调试模式）
        if hasattr(self, 'confirm_click') and self.confirm_click.get():
            self._save_screenshot("before_scroll")
        
        # 根据是否是调试模式选择不同的滚动策略
        if hasattr(self, 'confirm_click') and self.confirm_click.get():
            # 调试模式下，只使用JavaScript滚动和键盘滚动，不使用鼠标拖动
            scroll_methods = [
                self._scroll_with_javascript,
                self._scroll_with_keys,
                self._click_next_page
            ]
            self._log_info("调试模式下使用精确滚动", "blue")
            
            # 调试模式下增加重试次数和等待时间
            max_retries = 5  # 增加到5次尝试
            wait_time = 3.0  # 调试模式下等待更长时间
        else:
            # 正常模式下，使用JavaScript滚动和键盘滚动，不使用鼠标拖动
            scroll_methods = [
                self._scroll_with_javascript,
                self._scroll_with_keys,
                self._click_next_page
            ]
            max_retries = 1
            wait_time = 1.5
        
        # 尝试多次滚动
        for retry in range(max_retries):
            # 检查验证码和暂停状态 - 修复点3
            if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                self._log_info(f"滚动重试 {retry+1}/{max_retries} 中检测到验证码，停止滚动", "red")
                return False
                
            if hasattr(self, 'is_paused') and self.is_paused:
                self._log_info(f"滚动重试 {retry+1}/{max_retries} 中检测到暂停状态，停止滚动", "orange")
                return False
            
            if retry > 0:
                self._log_info(f"第 {retry+1}/{max_retries} 次尝试滚动", "blue")
                
                # 重试时增加滚动距离
                multiplier = (retry + 1) * 0.8  # 增加倍数，使滚动距离更接近正确位置
                self._log_info(f"增加滚动距离倍数: {multiplier}x", "blue")
                
                # 在调试模式下，每次重试前等待更长时间让页面稳定
                if hasattr(self, 'confirm_click') and self.confirm_click.get():
                    time.sleep(1.0)
            else:
                multiplier = 1.0
            
            # 在调试模式下，先尝试使用JavaScript滚动
            if hasattr(self, 'confirm_click') and self.confirm_click.get():
                try:
                    self._log_info(f"调试模式下先尝试JavaScript滚动 (倍数: {multiplier}x)", "blue")
                    success = self._scroll_with_javascript(multiplier=multiplier)
                    if not success:
                        self._log_info("JavaScript滚动返回失败状态", "orange")
                    
                    # 增加滚动后等待时间，让页面有足够时间更新
                    time.sleep(wait_time * 1.5)
                    
                    # 检查是否滚动到了新订单
                    new_order_id = self._extract_current_order_id()
                    self._log_info(f"滚动后订单ID: {new_order_id}", "blue")
                    
                    if new_order_id and new_order_id != current_order_id:
                        self._log_info(f"成功滚动到新订单: {new_order_id}", "green")
                        return True
                    
                    # 参考代码的策略：即使ID没变，也尝试查找下一个订单的元素
                    # 这是关键改进：先滚动，然后尝试查找下一个订单的元素
                    try:
                        # 尝试定位第二个订单的元素
                        if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                            self._log_info("尝试直接查找第二个订单的参照元素", "blue")
                            element = self.driver.find_element(By.XPATH, self.ref2_xpath)
                            if element:
                                self._log_info("成功找到第二个订单的参照元素，滚动成功", "green")
                                return True
                    except Exception as e:
                        self._log_info(f"查找第二个订单参照元素失败: {str(e)}", "orange")
                except Exception as e:
                    self._log_info(f"JavaScript滚动失败: {str(e)}", "red")
            
            # 尝试所有滚动方法
            for method in scroll_methods:
                # 检查验证码和暂停状态 - 修复点4
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info(f"滚动方法 {method.__name__} 执行前检测到验证码，停止滚动", "red")
                    return False
                    
                if hasattr(self, 'is_paused') and self.is_paused:
                    self._log_info(f"滚动方法 {method.__name__} 执行前检测到暂停状态，停止滚动", "orange")
                    return False
                
                try:
                    method_name = method.__name__
                    self._log_info(f"尝试使用 {method_name} 滚动到下一个订单", "blue")
                    
                    # 对JavaScript滚动方法传递倍数参数
                    if method_name == '_scroll_with_javascript':
                        success = method(multiplier=multiplier)
                        if not success:
                            self._log_info(f"{method_name} 返回失败状态", "orange")
                            continue
                    else:
                        method()
                    
                    # 增加滚动后等待时间，让页面有足够时间更新
                    time.sleep(wait_time * 1.5)
                    
                    # 检查是否滚动到了新订单
                    new_order_id = self._extract_current_order_id()
                    self._log_info(f"滚动后订单ID: {new_order_id}", "blue")
                    
                    if new_order_id and new_order_id != current_order_id:
                        self._log_info(f"成功滚动到新订单: {new_order_id}", "green")
                        return True
                    
                    # 参考代码的策略：即使ID没变，也尝试查找下一个订单的元素
                    try:
                        # 尝试定位第二个订单的元素
                        if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                            self._log_info("尝试直接查找第二个订单的参照元素", "blue")
                            element = self.driver.find_element(By.XPATH, self.ref2_xpath)
                            if element:
                                self._log_info("成功找到第二个订单的参照元素，滚动成功", "green")
                                return True
                    except Exception as e:
                        self._log_info(f"查找第二个订单参照元素失败: {str(e)}", "orange")
                    
                    self._log_info(f"{method_name} 滚动后订单未变化，尝试下一种方法", "orange")
                    
                    # 验证页面是否有变化（通过检查页面高度或其他元素）
                    try:
                        scroll_position_before = self.driver.execute_script("return window.pageYOffset;")
                        page_height_before = self.driver.execute_script("return document.body.scrollHeight;")
                        
                        # 尝试再次小幅滚动
                        self.driver.execute_script("window.scrollBy(0, 50);")
                        time.sleep(0.5)
                        
                        scroll_position_after = self.driver.execute_script("return window.pageYOffset;")
                        page_height_after = self.driver.execute_script("return document.body.scrollHeight;")
                        
                        if scroll_position_after > scroll_position_before or page_height_after != page_height_before:
                            self._log_info("检测到页面有变化，可能正在加载新内容", "blue")
                            time.sleep(1.0)  # 等待更长时间让内容加载
                            
                            # 再次检查订单ID
                            new_order_id = self._extract_current_order_id()
                            if new_order_id and new_order_id != current_order_id:
                                self._log_info(f"延迟检测到新订单: {new_order_id}", "green")
                                return True
                            # 再次尝试查找第二个订单的元素
                            try:
                                if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                                    element = self.driver.find_element(By.XPATH, self.ref2_xpath)
                                    if element:
                                        self._log_info("延迟成功找到第二个订单的参照元素，滚动成功", "green")
                                        return True
                            except Exception:
                                pass
                    except Exception as e:
                        self._log_info(f"页面变化检测失败: {str(e)}", "orange")
                except Exception as e:
                    self._log_info(f"{method_name} 滚动失败: {str(e)}", "red")
                    import traceback
                    logging.error(f"滚动异常: {traceback.format_exc()}")
            
            # 如果所有方法都失败，尝试点击页面中间然后按Page Down键
            if retry == max_retries - 1:  # 最后一次重试
                try:
                    self._log_info("尝试点击页面中间并按Page Down键", "blue")
                    
                    # 获取浏览器窗口位置和大小
                    window_rect = self.driver.execute_script("""
                        return {
                            width: window.innerWidth,
                            height: window.innerHeight
                        };
                    """)
                    
                    # 点击页面中间
                    actions = ActionChains(self.driver)
                    actions.move_by_offset(window_rect['width'] // 2, window_rect['height'] // 2)
                    actions.click()
                    actions.perform()
                    
                    # 等待一下
                    time.sleep(0.5)
                    
                    # 发送Page Down键
                    actions = ActionChains(self.driver)
                    actions.send_keys(Keys.PAGE_DOWN)
                    actions.perform()
                    
                    # 增加滚动后等待时间
                    time.sleep(wait_time * 1.5)
                    
                    # 检查是否滚动到了新订单
                    new_order_id = self._extract_current_order_id()
                    if new_order_id and new_order_id != current_order_id:
                        self._log_info(f"点击页面中间并按Page Down键成功滚动到新订单: {new_order_id}", "green")
                        return True
                    
                    # 参考代码的策略：即使ID没变，也尝试查找下一个订单的元素
                    try:
                        if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                            element = self.driver.find_element(By.XPATH, self.ref2_xpath)
                            if element:
                                self._log_info("通过Page Down成功找到第二个订单的参照元素，滚动成功", "green")
                                return True
                    except Exception:
                        pass
                except Exception as e:
                    self._log_info(f"点击页面中间并按Page Down键失败: {str(e)}", "red")
        
        # 如果所有方法都失败了，尝试增加滚动距离并重试
        self._log_info("所有滚动方法都失败，尝试增加滚动距离", "orange")
        try:
            # 增加滚动距离，再次尝试JavaScript滚动
            multiplier = 2.0 if hasattr(self, 'confirm_click') and self.confirm_click.get() else 1.5
            self._log_info(f"使用 {multiplier}x 倍滚动距离", "blue")
            success = self._scroll_with_javascript(multiplier=multiplier)
            if not success:
                self._log_info("最终JavaScript滚动返回失败状态", "orange")
                
            # 增加滚动后等待时间
            time.sleep(wait_time * 2)
            
            # 再次检查是否滚动到了新订单
            new_order_id = self._extract_current_order_id()
            if new_order_id and new_order_id != current_order_id:
                self._log_info(f"增加滚动距离后成功到达新订单: {new_order_id}", "green")
                return True
            
            # 参考代码的策略：即使ID没变，也尝试查找下一个订单的元素
            try:
                if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                    element = self.driver.find_element(By.XPATH, self.ref2_xpath)
                    if element:
                        self._log_info("最终成功找到第二个订单的参照元素，滚动成功", "green")
                        return True
            except Exception:
                pass
        except Exception as e:
            self._log_info(f"增加滚动距离失败: {str(e)}", "red")
        
            self._log_info("所有滚动方法都失败，无法到达下一个订单", "red")
            return False
        

    def _scroll_with_javascript(self, multiplier=1.0):
        """使用JavaScript滚动页面，确保第二个容器滚动到第一个容器的位置"""
        # 检查验证码和暂停状态 - 修复点1
        if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
            self._log_info("检测到验证码，停止滚动操作", "red")
            return False
            
        if hasattr(self, 'is_paused') and self.is_paused:
            self._log_info("检测到暂停状态，停止滚动操作", "orange")
            return False
            
        # 默认滚动距离，可以根据实际情况调整
        scroll_distance = 150 * multiplier  # 增加基础滚动距离
        
        # 调试模式下使用更精确的滚动
        if hasattr(self, 'confirm_click') and self.confirm_click.get():
            scroll_distance = 200 * multiplier  # 调试模式下使用更大的滚动距离
            self._log_info(f"调试模式下使用精确滚动距离: {scroll_distance}px", "blue")
            
            # 在调试模式下，保存滚动前的截图
            self._save_screenshot("before_js_scroll")
        
        # 尝试通过参照点计算精确滚动距离
        try:
            # 如果有两个参照点，尝试计算它们之间的距离
            if hasattr(self, 'ref1_xpath') and hasattr(self, 'ref2_xpath') and self.ref1_xpath and self.ref2_xpath:
                self._log_info("尝试使用参照点计算滚动距离", "blue")
                
                # 获取第一个参照点的位置
                ref1_position = self.driver.execute_script(f"""
                    const element = document.evaluate('{self.ref1_xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (element) {{
                        const rect = element.getBoundingClientRect();
                        return {{
                            top: rect.top,
                            left: rect.left,
                            height: rect.height,
                            exists: true
                        }};
                    }}
                    return {{ exists: false }};
                """)
                
                # 获取第二个参照点的位置
                ref2_position = self.driver.execute_script(f"""
                    const element = document.evaluate('{self.ref2_xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                    if (element) {{
                        const rect = element.getBoundingClientRect();
                        return {{
                            top: rect.top,
                            left: rect.left,
                            height: rect.height,
                            exists: true
                        }};
                    }}
                    return {{ exists: false }};
                """)
                
                # 如果两个参照点都存在，计算它们之间的距离
                if ref1_position.get('exists', False) and ref2_position.get('exists', False):
                    # 计算两个参照点之间的垂直距离
                    distance = ref2_position['top'] - ref1_position['top']
                    
                    if distance > 0:
                        self._log_info(f"参照点之间的垂直距离: {distance}px", "blue")
                        
                        # 获取页面总高度和可视区域高度，以便计算滚动比例
                        page_metrics = self.driver.execute_script("""
                            return {
                                scrollHeight: document.documentElement.scrollHeight,
                                clientHeight: document.documentElement.clientHeight,
                                scrollTop: document.documentElement.scrollTop
                            };
                        """)
                        
                        # 计算滚动比例
                        scroll_height = page_metrics.get('scrollHeight', 3000)
                        client_height = page_metrics.get('clientHeight', 800)
                        
                        # 计算滚动条移动距离与实际页面移动距离的比例
                        if scroll_height > client_height:
                            scroll_ratio = (scroll_height - client_height) / scroll_height
                            # 使用参照点之间的距离作为滚动距离，考虑滚动条比例并增加一点额外距离
                            scroll_distance = distance * scroll_ratio * 1.2  # 增加20%的距离
                            self._log_info(f"页面高度: {scroll_height}px, 可视区域: {client_height}px, 滚动比例: {scroll_ratio:.2f}", "blue")
                        else:
                            # 如果页面不需要滚动，使用默认值
                            scroll_distance = distance * 0.8  # 增加滚动距离
                        
                        self._log_info(f"使用调整后的滚动距离: {scroll_distance}px", "blue")
                    else:
                        self._log_info(f"参照点之间的垂直距离为负或零: {distance}px，使用默认距离", "orange")
                else:
                    self._log_info("无法获取参照点位置，使用默认滚动距离", "orange")
        except Exception as e:
            self._log_info(f"计算参照点距离失败: {str(e)}", "orange")
        
        # 直接滚动页面，不尝试滚动容器
        self._log_info(f"直接滚动整个页面 {scroll_distance}px", "blue")
        
        # 使用更精确的滚动方法，考虑滚动条和页面的比例关系
        try:
            # 获取页面总高度和可视区域高度，以便计算滚动比例
            page_metrics = self.driver.execute_script("""
                return {
                    scrollHeight: document.documentElement.scrollHeight,
                    clientHeight: document.documentElement.clientHeight,
                    scrollTop: document.documentElement.scrollTop,
                    pageYOffset: window.pageYOffset
                };
            """)
            
            self._log_info(f"页面高度: {page_metrics.get('scrollHeight', 0)}px, 可视区域高度: {page_metrics.get('clientHeight', 0)}px, 当前滚动位置: {page_metrics.get('scrollTop', 0)}px", "blue")
            
            # 计算滚动步长，分多次小幅度滚动以确保页面内容能够正确加载
            total_scroll = scroll_distance
            steps = 2  # 减少滚动步骤，每步滚动更多
            step_distance = total_scroll / steps
            
            for i in range(steps):
                # 检查验证码和暂停状态 - 修复点2
                if hasattr(self, 'force_stop_flag') and self.force_stop_flag:
                    self._log_info(f"滚动步骤 {i+1}/{steps} 中检测到验证码，停止滚动", "red")
                    return False
                    
                if hasattr(self, 'is_paused') and self.is_paused:
                    self._log_info(f"滚动步骤 {i+1}/{steps} 中检测到暂停状态，停止滚动", "orange")
                    return False
                
                # 执行滚动
                result = self.driver.execute_script(f"""
                    const originalScrollTop = window.pageYOffset;
                    window.scrollBy(0, {step_distance});
                    return {{
                        before: originalScrollTop,
                        after: window.pageYOffset,
                        scrolled: window.pageYOffset - originalScrollTop
                    }};
                """)
                
                before_position = result.get('before', 0)
                after_position = result.get('after', 0)
                scrolled = result.get('scrolled', 0)
                
                self._log_info(f"滚动步骤 {i+1}/{steps}: 从 {before_position} 到 {after_position} (滚动: {scrolled}px)", "blue")
                
                # 短暂等待，让页面内容加载
                time.sleep(0.3)
            
            # 在调试模式下，保存滚动后的截图
            if hasattr(self, 'confirm_click') and self.confirm_click.get():
                self._save_screenshot("after_window_scroll")
                
            # 如果滚动距离太小，尝试一次更大的滚动
            if total_scroll < 100:
                self._log_info("滚动距离太小，尝试更大的滚动", "blue")
                result = self.driver.execute_script(f"""
                    const originalScrollTop = window.pageYOffset;
                    window.scrollBy(0, 200);
                    return {{
                        before: originalScrollTop,
                        after: window.pageYOffset,
                        scrolled: window.pageYOffset - originalScrollTop
                    }};
                """)
                
                before_position = result.get('before', 0)
                after_position = result.get('after', 0)
                scrolled = result.get('scrolled', 0)
                
                self._log_info(f"额外滚动: 从 {before_position} 到 {after_position} (滚动: {scrolled}px)", "blue")
            
            return True
        except Exception as e:
            self._log_info(f"精确滚动失败: {str(e)}", "orange")
        
        # 如果精确滚动失败，尝试常规滚动方法
        scroll_methods = [
            f"window.scrollBy(0, {scroll_distance}); return {{before: window.pageYOffset, after: window.pageYOffset + {scroll_distance}}};",
            f"const before = window.pageYOffset; window.scrollTo(0, before + {scroll_distance}); return {{before: before, after: window.pageYOffset}};",
            f"const before = document.documentElement.scrollTop; document.documentElement.scrollTop += {scroll_distance}; return {{before: before, after: document.documentElement.scrollTop}};",
            f"const before = document.body.scrollTop; document.body.scrollTop += {scroll_distance}; return {{before: before, after: document.body.scrollTop}};"
        ]
        
        for method in scroll_methods:
            try:
                result = self.driver.execute_script(method)
                before_position = result.get('before', 0)
                after_position = result.get('after', 0)
                
                if after_position > before_position:
                    self._log_info(f"已滚动窗口从 {before_position} 到 {after_position} (距离: {after_position - before_position}px)", "blue")
                    
                    # 在调试模式下，保存滚动后的截图
                    if hasattr(self, 'confirm_click') and self.confirm_click.get():
                        self._save_screenshot("after_window_scroll")
                        
                    return True
            except Exception as e:
                continue
        
        # 最后尝试使用键盘滚动
        try:
            actions = ActionChains(self.driver)
            actions.send_keys(Keys.PAGE_DOWN)
            actions.perform()
            self._log_info("使用键盘PAGE_DOWN滚动", "blue")
            
            # 在调试模式下，保存滚动后的截图
            if hasattr(self, 'confirm_click') and self.confirm_click.get():
                self._save_screenshot("after_pagedown_scroll")
                
            return True
        except Exception as e:
            self._log_info(f"键盘滚动失败: {str(e)}", "orange")
        
        self._log_info("所有窗口滚动方法都失败", "orange")
        return False
    

    def _swipe_with_pyautogui(self):
        """使用PyAutoGUI模拟滑动手势"""
        try:
            # 获取浏览器窗口位置和大小
            window_rect = self.driver.execute_script("""
                return {
                    left: window.screenX || window.screenLeft,
                    top: window.screenY || window.screenTop,
                    width: window.outerWidth,
                    height: window.innerHeight
                };
            """)
            
            # 计算滑动的起点和终点（从窗口中心开始，向上滑动）
            start_x = window_rect['left'] + window_rect['width'] // 2
            start_y = window_rect['top'] + window_rect['height'] // 2
            end_y = start_y - 300  # 向上滑动300像素
            
            # 执行滑动操作
            pyautogui.moveTo(start_x, start_y)
            pyautogui.mouseDown()
            pyautogui.moveTo(start_x, end_y, duration=0.5)
            pyautogui.mouseUp()
            
            self._log_info(f"已执行滑动: 从 ({start_x}, {start_y}) 到 ({start_x}, {end_y})", "blue")
        except Exception as e:
            self._log_info(f"PyAutoGUI滑动失败: {str(e)}", "red")
            raise
    

    def _scroll_with_keys(self):
        """使用键盘按键滚动页面"""
        try:
            # 确保页面有焦点
            self.driver.execute_script("window.focus();")
            
            # 发送Page Down键
            actions = ActionChains(self.driver)
            actions.send_keys(Keys.PAGE_DOWN)
            actions.perform()
            
            self._log_info("已发送Page Down键", "blue")
        except Exception as e:
            self._log_info(f"键盘滚动失败: {str(e)}", "red")
            raise
    

    def _click_next_page(self):
        """尝试点击"下一页"按钮"""
        try:
            # 尝试多种可能的"下一页"按钮选择器
            next_page_selectors = [
                "//a[contains(text(), '下一页')]",
                "//button[contains(text(), '下一页')]",
                "//span[contains(text(), '下一页')]",
                "//a[contains(@class, 'next')]",
                "//button[contains(@class, 'next')]",
                "//li[contains(@class, 'next')]",
                "//i[contains(@class, 'next')]"
            ]
            
            for selector in next_page_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    if elements:
                        elements[0].click()
                        self._log_info(f"已点击下一页按钮: {selector}", "blue")
                        return True
                except Exception:
                    pass
            
            self._log_info("未找到下一页按钮", "orange")
            return False
        except Exception as e:
            self._log_info(f"点击下一页失败: {str(e)}", "red")
            return False
    
    # ==================== 模块化翻页核心方法 ====================
    
    def _run_modular_paging_loop(self, manual_order_count=None):
        """模块化翻页循环"""
        import math
        
        # 获取总订单数
        total_orders = self._get_total_order_count(manual_order_count)
        if total_orders <= 0:
            return
            
        # 获取每页订单数
        page_size = int(self.page_count_var.get()) if hasattr(self, 'page_count_var') else 20
        total_pages = math.ceil(total_orders / page_size)
        
        self._log_info(f"开始模块化处理：总订单{total_orders}个，每页{page_size}个，共{total_pages}页", "green")
        
        # 页面循环 - 新增的外层循环
        for current_page in range(1, total_pages + 1):
            if not self.is_running:
                break
                
            # 计算当前页要处理的订单数
            remaining_orders = total_orders - (current_page - 1) * page_size
            current_page_orders = min(page_size, remaining_orders)
            
            self._log_info(f"开始处理第{current_page}页，共{current_page_orders}个订单", "blue")
            
            # 每页开始时重置状态 - 关键改进
            self._reset_page_state()
            
            # 调用原有的单页处理逻辑 - 保持原有代码不变
            success = self._process_single_page(current_page_orders, current_page)
            
            if not success:
                self._log_info(f"第{current_page}页处理失败，停止执行", "red")
                break
                
            # 如果不是最后一页，执行翻页
            if current_page < total_pages and self.is_running:
                if not self._execute_page_turn():
                    self._log_info(f"第{current_page}页翻页失败，停止执行", "red")
                    break
                self._log_info(f"第{current_page}页处理完成，已翻页到第{current_page+1}页", "green")
        
        self._log_info(f"模块化处理完成，共处理{len(self.collected_data)}个订单", "green")
        self._stop_collection()
    
    def _get_total_order_count(self, manual_order_count):
        """获取总订单数"""
        # 检查是否有订单数量元素被标记为自动检测
        order_count_elements = [op for op in self.operation_sequence if op.get("is_order_count", False)]
        
        if order_count_elements:
            # 有订单数量元素，使用自动检测模式
            count_action = order_count_elements[0]
            try:
                self._log_info(f"执行: {count_action['name']} (自动获取总数)")
                element = self._find_element_smart(count_action['name'], count_action['xpath'])
                if not element:
                    self._log_info('无法找到总数元素，请检查XPath。', 'red')
                    return 0
                    
                count_text = element.text.strip()
                import re
                numbers = re.findall(r'\d+', count_text)
                if not numbers:
                    self._log_info(f'错误: 在 "{count_text}" 中未找到数字。', 'red')
                    return 0
                num_items = int(numbers[0])
                self._log_info(f'自动检测到项目总数: {num_items}')
                return num_items
            except Exception as e:
                self._log_info(f'自动获取总数失败: {e}', 'red')
                return 0
        else:
            # 没有订单数量元素，使用手动输入模式
            if manual_order_count is None:
                self._log_info('错误：未提供手动输入的订单数量', 'red')
                return 0
            self._log_info(f"使用手动输入的订单数量: {manual_order_count}", "blue")
            return manual_order_count
    
    def _reset_page_state(self):
        """重置页面状态 - 解决跨页问题的关键"""
        # 清除XPath学习缓存
        if hasattr(self, '_xpath_pattern_cache'):
            self._xpath_pattern_cache = None
        
        # 重置订单ID检测
        if hasattr(self, 'processed_order_ids'):
            self.processed_order_ids.clear()
        else:
            self.processed_order_ids = set()
        
        # 重置重复订单计数
        self.consecutive_same_order = 0
        
        self._log_info("页面状态已重置", "blue")
    
    def _process_single_page(self, page_orders, page_num):
        """处理单页订单 - 原有逻辑的封装"""
        try:
            # 获取操作序列
            actions_to_loop = [op for op in self.operation_sequence if not op.get("is_order_count", False)]
            if not actions_to_loop:
                self._log_info('错误：没有可执行的操作元素', 'red')
                return False
            
            # 学习XPath模式（每页重新学习）
            first_action_xpath = actions_to_loop[0]['xpath']
            xpath_pattern = self._learn_xpath_pattern_for_page(first_action_xpath)
            
            # 处理当前页的所有订单
            for order_index in range(1, page_orders + 1):
                if not self.is_running:
                    return False
                    
                # 更新双层进度显示
                self._update_dual_progress(page_num, order_index, page_orders)
                
                # 处理当前订单
                success = self._process_single_order(order_index, actions_to_loop, xpath_pattern)
                if not success:
                    return False
                    
                # 滚动到下一个订单（如果不是最后一个）
                if order_index < page_orders:
                    if not self._scroll_to_next_order():
                        self._log_info("无法滚动到下一个订单，尝试继续处理", "orange")
                        # 即使滚动失败，也尝试继续处理
            
            return True
            
        except Exception as e:
            self._log_info(f"处理第{page_num}页时发生错误: {e}", "red")
            return False
    
    def _learn_xpath_pattern_for_page(self, first_action_xpath):
        """为当前页学习XPath模式"""
        xpath_pattern = None
        
        # 尝试使用参照点学习XPath模式
        if hasattr(self, 'ref2_xpath') and self.ref2_xpath:
            xpath_pattern = self._learn_xpath_pattern(first_action_xpath, self.ref2_xpath)
            if not xpath_pattern:
                self._log_info('无法从参照XPath中学习到规律，将回退到基本模式。', 'orange')
                
        # 如果无法从参照点学习，尝试基本模式
        if not xpath_pattern:
            self._log_info('使用基本循环模式（递增最后一个数字）...', 'info')
            import re
            base_index = -1
            matches = list(re.finditer(r'\[(\d+)\]', first_action_xpath))
            if matches:
                base_index = int(matches[-1].group(1))
                self._log_info(f'检测到列表起始索引为: {base_index}')
                xpath_pattern = {'diff_segment_index': len(first_action_xpath.split('/'))-1, 'template': '', 'start_index': base_index}
            else:
                self._log_info('循环模式错误: 无法在第一个操作的XPath中找到列表索引（如 [1], [2]）。', 'red')
                return None
                
        return xpath_pattern
    
    def _update_dual_progress(self, page_num, order_index, page_orders):
        """更新双层进度显示"""
        import math
        page_size = int(self.page_count_var.get()) if hasattr(self, 'page_count_var') else 20
        total_pages = math.ceil(len(self.collected_data) / page_size) if hasattr(self, 'collected_data') else page_num
        
        progress_text = f"第{page_num}页/共{total_pages}页 - 当前页第{order_index}个/共{page_orders}个"
        
        if hasattr(self, 'progress_label'):
            self.progress_label.config(text=progress_text)
        
        # 计算总体进度
        total_processed = (page_num - 1) * page_size + order_index
        if hasattr(self, 'progress_bar'):
            self.progress_bar["value"] = total_processed
            
        if hasattr(self, 'root'):
            self.root.update()
    
    def _process_single_order(self, order_index, actions_to_loop, xpath_pattern):
        """处理单个订单"""
        order_data = {}
        
        # 处理当前订单的所有操作
        for op in actions_to_loop:
            # 检查是否已终止操作
            if not self.is_running:
                return False
                
            # 检查暂停状态
            if hasattr(self, 'is_paused') and self.is_paused:
                while hasattr(self, 'is_paused') and self.is_paused and self.is_running:
                    import time
                    time.sleep(0.5)
                    if hasattr(self, 'root'):
                        self.root.update()
                if not self.is_running:
                    return False
            
            # 生成当前订单的XPath
            op_xpath = op.get('xpath', '')
            op_item_xpath = self._generate_xpath_for_item(op_xpath, order_index, xpath_pattern)
            op_copy = op.copy()
            op_copy['xpath'] = op_item_xpath
            
            try:
                result = self._execute_operation(op_copy)
                if result is not None:
                    order_data[op['name']] = result
                    
                # 如果"点击前确认"未勾选，每个操作后添加延迟
                if hasattr(self, 'confirm_click') and not self.confirm_click.get():
                    import time
                    time.sleep(self.auto_action_interval if hasattr(self, 'auto_action_interval') else 1.0)
                    
            except Exception as e:
                self._log_info(f"执行'{op['name']}'操作失败: {str(e)}", "red")
        
        # 检查是否成功采集了订单数据
        if order_data:
            # 获取订单ID
            current_order_id = order_data.get('订单编号', '')
            
            # 写入订单基础数据到缓存
            if current_order_id:
                self.cache_manager.write_order_data(current_order_id, order_data=order_data)
            
            # 如果包含收货信息，同时写入收货信息
            if '复制完整收货信息' in order_data or '复制完整的收货信息' in order_data:
                if current_order_id and isinstance(current_order_id, str):
                    shipping_info = order_data.get('复制完整收货信息') or order_data.get('复制完整的收货信息')
                    if shipping_info:
                        # 写入到数据缓存
                        self.cache_manager.write_order_data(current_order_id, shipping_info=shipping_info)
                        self._log_info(f"已建立订单ID与收货信息的直接关联: {current_order_id}", "green")
                        
                        # 保持向后兼容性
                        clean_order_id = current_order_id.replace('订单编号：', '')
                        if not hasattr(self, 'order_clipboard_contents'):
                            self.order_clipboard_contents = {}
                        self.order_clipboard_contents[clean_order_id] = shipping_info
            
            # 检查是否是重复订单
            current_order_id = None
            for key in ['订单编号', '订单ID', 'order_id', 'orderid']:
                if key in order_data:
                    current_order_id = order_data[key]
                    break
            
            if current_order_id:
                if current_order_id in self.processed_order_ids:
                    self.consecutive_same_order += 1
                    self._log_info(f"检测到重复订单: {current_order_id}，这是第{self.consecutive_same_order}次重复，跳过此订单", "orange")
                    
                    # 如果连续3次重复，记录警告但继续处理（不停止整个流程）
                    if self.consecutive_same_order >= 3:
                        self._log_info("连续多次重复订单，可能存在页面滚动问题，但继续尝试处理", "orange")
                    # 跳过重复订单，不添加到collected_data，但返回True继续处理下一个
                    return True
                else:
                    self.processed_order_ids.add(current_order_id)
                    self.consecutive_same_order = 0
                    if not hasattr(self, 'collected_data'):
                        self.collected_data = []
                    self.collected_data.append(order_data)
                    self._log_info(f"成功采集订单: {current_order_id}", "green")
            else:
                # 没有找到订单ID，但仍然添加数据
                if not hasattr(self, 'collected_data'):
                    self.collected_data = []
                self.collected_data.append(order_data)
        
        return True
    
    def _execute_page_turn(self):
        """执行翻页操作"""
        try:
            if hasattr(self, 'execute_page_turn'):
                return self.execute_page_turn()  # 调用现有的翻页方法
            else:
                self._log_info("翻页功能未配置", "orange")
                return False
        except Exception as e:
            self._log_info(f"翻页失败: {e}", "red")
            return False
    
    def _sync_ui_modular_paging_state(self):
        """同步UI的模块化翻页状态到实例属性"""
        try:
            # 从UI组件获取模块化翻页状态
            if hasattr(self, 'use_modular_paging_value'):
                self.use_modular_paging = self.use_modular_paging_value
            elif hasattr(self, 'use_modular_paging') and hasattr(self.use_modular_paging, 'get'):
                # 如果use_modular_paging是tkinter变量
                self.use_modular_paging = self.use_modular_paging.get()
            else:
                # 默认值
                self.use_modular_paging = False
                
        except Exception as e:
            self._log_info(f"同步模块化翻页状态失败: {e}", "orange")
            self.use_modular_paging = False
            

    def _check_shipping_info_before_export(self):
        """检查并尝试修复收货信息字段，使用数据缓存管理器（只读权限）"""
        # 从数据缓存读取所有订单数据（只读权限）
        cached_orders = self.cache_manager.read_all_orders()
        
        if not cached_orders:
            self._log_info("数据缓存中没有订单数据可以导出", "red")
            from tkinter import messagebox
            messagebox.showwarning("导出失败", "数据缓存中没有订单数据，请先采集数据。")
            return False
            
        field_name = '复制完整收货信息'  # 收货信息字段名
        fixed_count = 0
        
        # 输出缓存统计信息
        orders_with_shipping = self.cache_manager.get_orders_with_shipping_info()
        self._log_info(f"数据缓存统计: 总订单数={len(cached_orders)}, 包含收货信息的订单数={len(orders_with_shipping)}", "blue")
        
        # 确保collected_data存在并从缓存重建
        if not hasattr(self, 'collected_data'):
            self.collected_data = []
        
        # 从缓存重建collected_data（只读操作）
        self.collected_data.clear()
        for order_id, cached_data in cached_orders.items():
            # 创建订单数据副本，排除系统字段
            system_fields = {"order_id", "created_at", "updated_at", "status", "shipping_info", "shipping_info_updated_at"}
            order_data = {k: v for k, v in cached_data.items() if k not in system_fields}
            
            # 如果缓存中有收货信息，添加到订单数据中
            if "shipping_info" in cached_data and cached_data["shipping_info"]:
                order_data[field_name] = cached_data["shipping_info"]
                fixed_count += 1
            
            if order_data:  # 确保有数据才添加
                self.collected_data.append(order_data)
        
        # 保持向后兼容性，更新order_clipboard_contents
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
        self.order_clipboard_contents.clear()
        
        for order_id, shipping_info in orders_with_shipping.items():
            clean_order_id = order_id.replace('订单编号：', '') if isinstance(order_id, str) else str(order_id)
            self.order_clipboard_contents[clean_order_id] = shipping_info["shipping_info"]
        
        # 检查每个订单的收货信息是否唯一
        shipping_info_set = set()
        duplicate_count = 0
        
        for order_data in self.collected_data:
            # 获取当前订单ID
            order_id = order_data.get('订单编号', '')
            if isinstance(order_id, str) and order_id.startswith('订单编号：'):
                order_id = order_id.replace('订单编号：', '')
            
            self._log_info(f"处理订单: {order_id}", "blue")
            
            # 检查是否存在此字段
            if field_name in order_data:
                # 检查内容是否为布尔值或空
                if order_data[field_name] is True or order_data[field_name] is False or not order_data[field_name]:
                    self._log_info(f"检测到订单 {order_id} 的无效收货信息字段值: {order_data[field_name]}", "orange")
                    
                    # 优先使用订单专属的收货信息
                    if order_id and order_id in self.order_clipboard_contents:
                        order_data[field_name] = self.order_clipboard_contents[order_id]
                        fixed_count += 1
                        self._log_info(f"使用订单专属收货信息修复: '{self.order_clipboard_contents[order_id][:30]}...'", "green")
                        
                        # 检查是否是重复的收货信息
                        if order_data[field_name] in shipping_info_set:
                            duplicate_count += 1
                            self._log_info(f"警告: 订单 {order_id} 的收货信息与其他订单重复", "red")
                        else:
                            shipping_info_set.add(order_data[field_name])
                        
                        continue
                    
                    # 其次尝试使用全局变量
                    if hasattr(self, 'last_clipboard_content') and self.last_clipboard_content:
                        order_data[field_name] = self.last_clipboard_content
                        fixed_count += 1
                        self._log_info(f"使用全局变量修复了收货信息: '{self.last_clipboard_content[:30]}...'", "green")
                        continue
                    
                    # 最后尝试从剪贴板获取最新内容
                    import pyperclip
                    clipboard_content = pyperclip.paste()
                    if clipboard_content and clipboard_content.strip():
                        order_data[field_name] = clipboard_content
                        fixed_count += 1
                        self._log_info(f"使用剪贴板内容修复了收货信息: '{clipboard_content[:30]}...'", "green")
                        continue
                        
                    # 如果都失败了，记录一个明确的错误信息
                    order_data[field_name] = "【收货信息获取失败】"
                    self._log_info(f"无法修复订单 {order_id} 的收货信息字段，已标记为失败", "red")
            else:
                # 如果字段不存在，优先使用订单专属的收货信息
                if order_id and order_id in self.order_clipboard_contents:
                    order_data[field_name] = self.order_clipboard_contents[order_id]
                    fixed_count += 1
                    self._log_info(f"添加了订单 {order_id} 的专属收货信息: '{self.order_clipboard_contents[order_id][:30]}...'", "green")
                    continue
                
                # 其次尝试使用全局变量
                elif hasattr(self, 'last_clipboard_content') and self.last_clipboard_content:
                    order_data[field_name] = self.last_clipboard_content
                    fixed_count += 1
                    self._log_info(f"添加了缺失的收货信息字段: '{self.last_clipboard_content[:30]}...'", "green")
                else:
                    # 最后尝试从剪贴板获取
                    import pyperclip
                    clipboard_content = pyperclip.paste()
                    if clipboard_content and clipboard_content.strip():
                        order_data[field_name] = clipboard_content
                        fixed_count += 1
                        self._log_info(f"添加了缺失的收货信息字段: '{clipboard_content[:30]}...'", "green")
        
        if fixed_count > 0:
            self._log_info(f"导出前共修复了 {fixed_count} 条收货信息记录", "green")
        
        if duplicate_count > 0:
            self._log_info(f"警告: 检测到 {duplicate_count} 个订单的收货信息重复", "red")
            
        # 导出前检查收货信息字段内容
        self._log_info("导出前检查收货信息字段内容:", "blue")
        for i, order_data in enumerate(self.collected_data):
            order_id = order_data.get('订单编号', '')
            if isinstance(order_id, str) and order_id.startswith('订单编号：'):
                order_id = order_id.replace('订单编号：', '')
            else:
                order_id = str(order_id)
                
            shipping_info = order_data.get(field_name, "未设置")
            shipping_info_type = type(shipping_info).__name__
            shipping_info_preview = shipping_info[:30] + "..." if isinstance(shipping_info, str) and shipping_info != "未设置" else str(shipping_info)
            
            print(f"DEBUG-EXPORT-CHECK-{i+1}: 订单ID '{order_id}', {field_name} = {shipping_info_preview} (类型: {shipping_info_type})")
            self._log_info(f"记录 {i+1}: 订单ID={order_id}, {field_name} = {shipping_info_preview} (类型: {shipping_info_type})", "blue")
            
        return fixed_count > 0



    
    def _export_excel(self):
        """导出数据到Excel（正常模式专用）"""
        # 检查pandas依赖
        if pd is None:
            self._log_info("pandas模块未正确导入，请安装: pip install pandas", "red")
            messagebox.showerror("错误", "pandas模块未安装，无法导出Excel文件")
            return
            
        # 确保是在正常模式下使用此功能
        if self.collection_mode.get() != "正常模式":
            messagebox.showinfo("提示", "Excel导出功能仅在正常模式下可用")
            return
            
        if len(self.collected_data) == 0:
            messagebox.showinfo("提示", "没有可导出的数据")
            return
        
        # 确保order_clipboard_contents字典存在
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
            self._log_info("警告: 导出前发现订单ID与收货信息映射字典不存在，已创建空字典", "red")
            print("DEBUG-EXPORT-EXCEL-ERROR: 订单ID与收货信息映射字典不存在，已创建空字典")
            
        # 检查并尝试修复收货信息
        self._check_shipping_info_before_export()
        
        # 导出前显示详细统计信息
        self._log_info(f"=== 导出统计信息 ===", "blue")
        self._log_info(f"总共采集到 {len(self.collected_data)} 个订单记录", "blue")
        
        # 按商品名称统计
        product_stats = {}
        for order_data in self.collected_data:
            product_name = None
            for field_name in ['商品名称', '商品', '产品名称', '产品', '商品信息']:
                if field_name in order_data and order_data[field_name]:
                    product_name = str(order_data[field_name]).strip()
                    break
            if not product_name or product_name == "nan":
                product_name = "未知商品"
            
            if product_name not in product_stats:
                product_stats[product_name] = 0
            product_stats[product_name] += 1
        
        for product_name, count in product_stats.items():
            self._log_info(f"  - {product_name}: {count} 个订单", "blue")
        self._log_info(f"=== 统计信息结束 ===", "blue")
        
        try:
            # 检查是否安装了openpyxl
            try:
                import openpyxl
            except ImportError:
                result = messagebox.askquestion("提示", 
                    "缺少Excel导出所需的openpyxl库。是否安装？\n(需要联网，可能需要几分钟)")
                if result == 'yes':
                    self._log_info("正在安装openpyxl库...", "blue")
                    import subprocess
                    import sys
                    try:
                        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
                        self._log_info("openpyxl库安装成功", "green")
                        import openpyxl
                    except Exception as e:
                        self._log_info(f"安装失败: {str(e)}", "red")
                        messagebox.showerror("错误", f"安装openpyxl失败，请手动安装:\npip install openpyxl\n\n错误信息: {str(e)}")
                        return
                else:
                    return
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not file_path:
                return
            
            # 创建数据框架，支持新的字典格式
            # 从collected_data创建DataFrame，每个订单一行
            df = pd.DataFrame(self.collected_data)
            
            # 检测并去重重复的列（如'复制完整收货信息'和'复制完整的收货信息'）
            duplicate_columns = []
            columns_to_merge = {
                '复制完整收货信息': ['复制完整收货信息', '复制完整的收货信息'],
                # 可以在这里添加其他需要合并的列
            }
            
            for target_col, source_cols in columns_to_merge.items():
                existing_cols = [col for col in source_cols if col in df.columns]
                if len(existing_cols) > 1:
                    self._log_info(f"检测到重复列: {existing_cols}，将合并为: {target_col}", "blue")
                    # 合并列数据，优先使用非空值
                    merged_data = []
                    for index, row in df.iterrows():
                        merged_value = ""
                        for col in existing_cols:
                            value = row.get(col, "")
                            if value and str(value).strip() and str(value).strip() != "nan":
                                merged_value = value
                                break
                        merged_data.append(merged_value)
                    
                    # 添加合并后的列
                    df[target_col] = merged_data
                    
                    # 标记要删除的重复列
                    duplicate_columns.extend(existing_cols)
            
            # 删除重复列
            if duplicate_columns:
                # 保留目标列，删除源列
                cols_to_drop = [col for col in duplicate_columns if col in columns_to_merge.keys() and col in df.columns]
                remaining_cols_to_drop = [col for col in duplicate_columns if col not in columns_to_merge.keys()]
                df = df.drop(columns=remaining_cols_to_drop)
                self._log_info(f"已删除重复列: {remaining_cols_to_drop}", "green")
            
            # 导出到Excel
            df.to_excel(file_path, index=False)
            self._log_info(f"Excel导出成功: {file_path}", "green")
        except Exception as e:
            self._log_info(f"Excel导出失败: {str(e)}", "red")
            import traceback
            self._log_info(traceback.format_exc(), "red")
    
    def _export_word(self):
        """导出数据到Word（正常模式专用）- Markdown格式，按商品名称分组"""
        # 检查docx依赖
        if Document is None:
            self._log_info("python-docx模块未正确导入，请安装: pip install python-docx", "red")
            messagebox.showerror("错误", "python-docx模块未安装，无法导出Word文件")
            return
            
        # 确保是在正常模式下使用此功能
        if self.collection_mode.get() != "正常模式":
            messagebox.showinfo("提示", "Word导出功能仅在正常模式下可用")
            return
            
        if len(self.collected_data) == 0:
            messagebox.showinfo("提示", "没有可导出的数据")
            return
            
        # 确保order_clipboard_contents字典存在
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
            self._log_info("警告: 导出前发现订单ID与收货信息映射字典不存在，已创建空字典", "red")
            print("DEBUG-EXPORT-WORD-ERROR: 订单ID与收货信息映射字典不存在，已创建空字典")
            
        # 检查并尝试修复收货信息
        self._check_shipping_info_before_export()
        
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word files", "*.docx")]
            )
            if not file_path:
                return
                
            doc = Document()
            
            # 添加文档标题
            title = doc.add_heading("收货信息采集报告", level=1)
            title.alignment = 1  # 居中对齐
            
            # 添加时间戳
            from datetime import datetime
            timestamp = doc.add_paragraph(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            timestamp.alignment = 1  # 居中对齐
            
            # 处理重复字段的去重逻辑
            columns_to_merge = {
                '复制完整收货信息': ['复制完整收货信息', '复制完整的收货信息'],
                # 可以在这里添加其他需要合并的列
            }
            
            # 创建处理后的数据副本
            processed_data = []
            for order_data in self.collected_data:
                processed_order = order_data.copy()
                
                # 处理每个需要合并的字段组
                for target_col, source_cols in columns_to_merge.items():
                    existing_cols = [col for col in source_cols if col in processed_order]
                    if len(existing_cols) > 1:
                        # 合并字段数据，优先使用非空值
                        merged_value = ""
                        for col in existing_cols:
                            value = processed_order.get(col, "")
                            if value and str(value).strip() and str(value).strip() != "nan":
                                merged_value = value
                                break
                        
                        # 设置合并后的值
                        processed_order[target_col] = merged_value
                        
                        # 删除重复的源字段（保留目标字段）
                        for col in existing_cols:
                            if col != target_col:
                                processed_order.pop(col, None)
                
                processed_data.append(processed_order)
            
            # 按商品名称分组数据
            grouped_data = {}
            for order_data in processed_data:
                # 尝试从多个可能的字段中获取商品名称
                product_name = None
                for field_name in ['商品名称', '商品', '产品名称', '产品', '商品信息']:
                    if field_name in order_data and order_data[field_name]:
                        product_name = str(order_data[field_name]).strip()
                        break
                
                # 如果没有找到商品名称，使用默认分组
                if not product_name or product_name == "nan":
                    product_name = "未知商品"
                
                if product_name not in grouped_data:
                    grouped_data[product_name] = []
                grouped_data[product_name].append(order_data)
            
            # 按商品名称排序
            sorted_products = sorted(grouped_data.keys())
            
            # 为每个商品组生成内容
            total_order_count = 0  # 全局订单计数器
            for product_name in sorted_products:
                orders = grouped_data[product_name]
                
                # 添加商品名称小标题
                product_heading = doc.add_heading(f"商品名称：{product_name}", level=2)
                
                # 为每个订单添加详细信息
                for i, order_data in enumerate(orders):
                    total_order_count += 1
                    # 始终添加订单序号，使用全局计数器确保唯一性
                    if len(orders) > 1:
                        order_heading = doc.add_heading(f"订单 {total_order_count}", level=3)
                    else:
                        # 即使只有一个订单，也显示序号以保持一致性
                        order_heading = doc.add_heading(f"订单 {total_order_count}", level=3)
                    
                    # 以markdown格式添加订单详细信息
                    for field_name, field_value in order_data.items():
                        if field_value and str(field_value).strip() and str(field_value).strip() != "nan":
                            # 创建字段信息段落
                            p = doc.add_paragraph()
                            # 添加字段名（加粗）
                            field_run = p.add_run(f"{field_name}：")
                            field_run.bold = True
                            # 添加字段值
                            value_run = p.add_run(str(field_value))
                    
                    # 在订单之间添加分隔线（除了最后一个订单）
                    if i < len(orders) - 1:
                        doc.add_paragraph("" + "-" * 50)
                
                # 在商品组之间添加空行
                if product_name != sorted_products[-1]:  # 不是最后一个商品组
                    doc.add_paragraph("")
                    doc.add_paragraph("" + "=" * 80)
                    doc.add_paragraph("")
            
            doc.save(file_path)
            self._log_info(f"Word导出成功: {file_path} (按商品名称分组，共{len(sorted_products)}个商品组)", "green")
        except Exception as e:
            self._log_info(f"Word导出失败: {str(e)}", "red")
            import traceback
            self._log_info(f"详细错误信息: {traceback.format_exc()}", "red")
    
    def _export_json(self):
        """导出数据到JSON（采集模式专用）"""
        # 确保是在采集模式下使用此功能
        if self.collection_mode.get() != "采集模式":
            messagebox.showinfo("提示", "JSON导出功能仅在采集模式下可用")
            return
            
        if not hasattr(self, 'collected_data') or len(self.collected_data) == 0:
            messagebox.showinfo("提示", "没有可导出的数据")
            return
            
        # 确保order_clipboard_contents字典存在
        if not hasattr(self, 'order_clipboard_contents'):
            self.order_clipboard_contents = {}
            self._log_info("警告: 导出前发现订单ID与收货信息映射字典不存在，已创建空字典", "red")
            print("DEBUG-EXPORT-JSON-ERROR: 订单ID与收货信息映射字典不存在，已创建空字典")
            
        # 检查并尝试修复收货信息
        self._check_shipping_info_before_export()
            
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".json",
                filetypes=[("JSON files", "*.json")]
            )
            if not file_path:
                return
                
            # 导出数据，确保JSON友好
            import json
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.collected_data, f, ensure_ascii=False, indent=2)
                
            self._log_info(f"JSON导出成功: {file_path}", "green")
        except Exception as e:
            self._log_info(f"JSON导出失败: {str(e)}", "red")
            import logging
            import traceback
            logging.error(f"JSON导出异常: {traceback.format_exc()}")
    
    def _show_verification_dialog_and_wait(self, remark, xpath, screen_x, screen_y, offset_x, offset_y):
        """弹窗让用户验证鼠标位置，支持WASD微调，保存元素特定偏移量"""
        self._last_offset_x = offset_x  # 初始化为当前元素的偏移量
        self._last_offset_y = offset_y
        result = {'ok': -1}
        
        def show_dialog():
            win = tk.Toplevel(self.root)
            win.title('请验证鼠标位置')
            win.geometry('350x220')
            win.transient(self.root)
            win.grab_set()
            msg = f'已将鼠标移动到 "{remark}" 的目标位置。\n\n请检查屏幕上的鼠标指针是否准确。\n\nA/D键: 左/右移1像素，W/S键: 上/下移1像素。\n点击"位置准确"后将保存此元素的偏移量并执行点击。\n点击"位置不准"将跳过本次点击并记录日志。'
            label = ttk.Label(win, text=msg, wraplength=320, justify=tk.LEFT)
            label.pack(pady=10, padx=10)
            offset_label = ttk.Label(win, text=f'当前偏移: X={self._last_offset_x}, Y={self._last_offset_y}')
            offset_label.pack(pady=2)
            
            def update_offset_label():
                offset_label.config(text=f'当前偏移: X={self._last_offset_x}, Y={self._last_offset_y}')
            
            def on_ok():
                result['ok'] = 1
                
                # 保存当前偏移量为元素特定的偏移量
                if remark not in self.element_offsets:
                    self.element_offsets[remark] = {}
                self.element_offsets[remark]["x"] = self._last_offset_x
                self.element_offsets[remark]["y"] = self._last_offset_y
                
                # 在当前位置执行点击
                # 记录点击时的坐标
                click_pos = pyautogui.position()
                self._log_info(f"[坐标日志] 确认模式 - 元素'{remark}' - 点击时坐标: X={click_pos.x}, Y={click_pos.y}", "green")
                
                pyautogui.click()
                self._log_info(f"已保存元素'{remark}'的偏移量 X={self._last_offset_x}, Y={self._last_offset_y} 并执行点击", "green")
                
                # 保存偏移量到配置文件
                self._save_offset_config()
                win.destroy()
                # --- 焦点恢复 ---
                self._manage_focus()
                self.root.after(100, self._manage_focus)
                self._log_info("已尝试恢复主窗口焦点", "blue")
                
            def on_fail():
                result['ok'] = 0
                win.destroy()
                # --- 焦点恢复 ---
                self._manage_focus()
                self.root.after(100, self._manage_focus)
                self._log_info("已尝试恢复主窗口焦点", "blue")
                
            btn_frame = ttk.Frame(win)
            btn_frame.pack(pady=10)
            ok_btn = ttk.Button(btn_frame, text='位置准确', command=on_ok)
            ok_btn.pack(side=tk.LEFT, padx=10)
            fail_btn = ttk.Button(btn_frame, text='位置不准', command=on_fail)
            fail_btn.pack(side=tk.LEFT, padx=10)
            win.bind('<Return>', lambda event: on_ok())
            win.bind('<Escape>', lambda event: on_fail())
            
            def on_key(event):
                moved = False
                if event.keysym.lower() == 'a':
                    self._last_offset_x -= 1
                    pyautogui.moveRel(-1, 0, duration=0)
                    moved = True
                elif event.keysym.lower() == 'd':
                    self._last_offset_x += 1
                    pyautogui.moveRel(1, 0, duration=0)
                    moved = True
                elif event.keysym.lower() == 'w':
                    self._last_offset_y -= 1
                    pyautogui.moveRel(0, -1, duration=0)
                    moved = True
                elif event.keysym.lower() == 's':
                    self._last_offset_y += 1
                    pyautogui.moveRel(0, 1, duration=0)
                    moved = True
                if moved:
                    update_offset_label()
                    
            win.bind('<Key>', on_key)
            
            # 移动鼠标到初始位置
            # 记录移动前的鼠标位置
            current_pos = pyautogui.position()
            self._log_info(f"[坐标日志] 确认模式 - 元素'{remark}' - 移动前鼠标位置: X={current_pos.x}, Y={current_pos.y}", "cyan")
            
            pyautogui.moveTo(int(screen_x), int(screen_y))
            
            # 记录移动后的鼠标位置
            moved_pos = pyautogui.position()
            self._log_info(f"[坐标日志] 确认模式 - 元素'{remark}' - 移动后鼠标位置: X={moved_pos.x}, Y={moved_pos.y}", "cyan")
            
            win.focus_set()
            
        self.root.after(0, show_dialog)
        # 阻塞等待用户操作
        while result['ok'] == -1:
            self.root.update()
            time.sleep(0.05)
        
        return result['ok'] == 1
    
    def _generate_relative_xpath(self, absolute_xpath):
        """从绝对XPath生成更健壮的相对XPath"""
        try:
            # 提取最后几个节点，通常包含更多特定信息
            parts = absolute_xpath.split('/')
            if len(parts) <= 3:  # XPath太短，无法生成有意义的相对路径
                return None
                
            # 从最后一个节点开始构建
            last_part = parts[-1]
            
            # 尝试使用文本内容、ID、class等属性构建相对XPath
            relative_xpaths = []
            
            # 1. 如果最后一个节点包含文本，使用文本构建
            if '[text()' in last_part or 'contains(text()' in last_part:
                text_match = re.search(r'text\(\)\s*=\s*[\'\"]([^\'\"])+[\'\"]', last_part)
                contains_match = re.search(r'contains\(text\(\),\s*[\'\"]([^\'\"])+[\'\"]', last_part)
                
                if text_match:
                    text = text_match.group(1)
                    relative_xpaths.append(f"//*[text()='{text}']")
                elif contains_match:
                    text = contains_match.group(1)
                    relative_xpaths.append(f"//*[contains(text(),'{text}')]")
            
            # 2. 如果包含ID属性，使用ID构建
            id_match = re.search(r'@id\s*=\s*[\'\"]([^\'\"])+[\'\"]', last_part)
            if id_match:
                id_value = id_match.group(1)
                relative_xpaths.append(f"//*[@id='{id_value}']")
            
            # 3. 如果包含class属性，使用class构建
            class_match = re.search(r'@class\s*=\s*[\'\"]([^\'\"])+[\'\"]', last_part)
            contains_class = re.search(r'contains\(@class,\s*[\'\"]([^\'\"])+[\'\"]', last_part)
            
            if class_match:
                class_value = class_match.group(1)
                relative_xpaths.append(f"//*[@class='{class_value}']")
            elif contains_class:
                class_value = contains_class.group(1)
                relative_xpaths.append(f"//*[contains(@class,'{class_value}')]")
            
            # 4. 如果是特定标签，添加标签限制
            tag_match = re.match(r'([a-z]+)(\[|$)', last_part)
            if tag_match:
                tag = tag_match.group(1)
                # 为每个已生成的XPath添加标签限制
                tagged_xpaths = []
                for xpath in relative_xpaths:
                    tagged_xpath = xpath.replace('*', tag)
                    tagged_xpaths.append(tagged_xpath)
                relative_xpaths.extend(tagged_xpaths)
            
            # 5. 如果上述方法都没有产生有效的相对XPath，尝试使用最后两个节点
            if not relative_xpaths and len(parts) >= 2:
                last_two_parts = '/'.join(parts[-2:])
                relative_xpaths.append(f"//{last_two_parts}")
            
            return relative_xpaths[0] if relative_xpaths else None
            
        except Exception as e:
            self._log_info(f"生成相对XPath失败: {str(e)}", "red")
            return None
            
    def _xpath_to_css(self, xpath):
        """尝试将XPath转换为CSS选择器（简化版本）"""
        try:
            # 这只是一个简化的转换，不能处理所有XPath
            if not xpath or xpath.startswith('//'):
                return None
                
            # 移除绝对路径前缀
            if xpath.startswith('/html/'):
                xpath = xpath[6:]
                
            parts = xpath.split('/')
            css_parts = []
            
            for part in parts:
                if not part:
                    continue
                    
                # 处理ID
                if '@id' in part:
                    match = re.search(r'@id\s*=\s*[\'\"]([^\'\"])+[\'\"]', part)
                    if match:
                        css_parts.append(f"#{match.group(1)}")
                        continue
                
                # 处理class
                if '@class' in part:
                    match = re.search(r'@class\s*=\s*[\'\"]([^\'\"])+[\'\"]', part)
                    if match:
                        classes = match.group(1).split()
                        css_parts.append(f".{'.'.join(classes)}")
                        continue
                
                # 处理基本标签和索引
                tag_match = re.match(r'([a-z]+)(?:\[(\d+)\])?', part)
                if tag_match:
                    tag = tag_match.group(1)
                    index = tag_match.group(2)
                    if index:
                        # CSS选择器索引从1开始，XPath也是从1开始
                        css_parts.append(f"{tag}:nth-child({index})")
                    else:
                        css_parts.append(tag)
                        
            return ' > '.join(css_parts) if css_parts else None
            
        except Exception:
            return None
            
    def _find_by_text_content(self, text):
        """通过文本内容查找元素"""
        try:
            # 尝试多种文本匹配策略
            selectors = [
                f"//*[text()='{text}']",
                f"//*[contains(text(),'{text}')]",
                f"//*[contains(normalize-space(text()),'{text}')]",
                f"//span[contains(text(),'{text}')]",
                f"//div[contains(text(),'{text}')]",
                f"//a[contains(text(),'{text}')]",
                f"//button[contains(text(),'{text}')]",
                f"//label[contains(text(),'{text}')]",
                f"//*[contains(@title,'{text}')]",
                f"//*[contains(@aria-label,'{text}')]"
            ]
            
            for selector in selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    if elements:
                        # 优先选择文本长度接近的元素
                        elements.sort(key=lambda e: abs(len(e.text) - len(text)))
                        return elements[0]
                except:
                    continue
                    
            return None
            
        except Exception as e:
            self._log_info(f"通过文本内容查找元素失败: {str(e)}", "red")
            return None
    
