#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
元素采集相关

从原始main.py文件中拆分出来的模块
"""

from utils import *
from operation_sequence_dialog import OperationSequenceDialog

class ElementCollector:
    """元素采集相关"""
    
    def __init__(self):
        """初始化元素采集器"""
        pass

    def _get_hovered_xpath_recursive(self):
        """递归获取当前悬停元素的XPath，支持嵌套iframe"""
        if not self.driver:
            self._log_info('无法获取XPath：浏览器未连接', 'red')
            return None
        get_xpath_script = '''
        function getXPath(element) {
            if (element && element.id) return '//*[@id="' + element.id + '"]';
            var paths = [];
            for (; element && element.nodeType == 1; element = element.parentNode) {
                var index = 0;
                var hasSimilarSibling = false;
                for (var sibling = element.previousSibling; sibling; sibling = sibling.previousSibling) {
                    if (sibling.nodeType == 1 && sibling.tagName === element.tagName) {
                        index++;
                    }
                }
                for (var sibling = element.nextSibling; sibling; sibling = sibling.nextSibling) {
                    if (sibling.nodeType == 1 && sibling.tagName === element.tagName) {
                        hasSimilarSibling = true;
                        break;
                    }
                }
                var tagName = element.tagName.toLowerCase();
                var pathIndex = (index > 0 || hasSimilarSibling) ? `[${index + 1}]` : '';
                paths.splice(0, 0, tagName + pathIndex);
            }
            return paths.length ? '/' + paths.join('/') : null;
        }
        if (!window.lastHoveredElement) return null;
        var xpath = getXPath(window.lastHoveredElement);
        window.lastHoveredElement = null;
        return xpath;
        '''
        try:
            xpath = self.driver.execute_script(get_xpath_script)
            if xpath:
                return xpath
            iframes = self.driver.find_elements(By.TAG_NAME, 'iframe')
            for i in range(len(iframes)):
                try:
                    self.driver.switch_to.frame(i)
                    xpath_in_frame = self._get_hovered_xpath_recursive()
                    if xpath_in_frame:
                        return xpath_in_frame
                except Exception:
                    pass
                finally:
                    self.driver.switch_to.parent_frame()
        except Exception as e:
            self._log_info(f'通过JS获取XPath时出错: {e}', 'red')
        return None
    
    def _get_element_at_cursor(self):
        """获取当前鼠标位置下的元素及其位置信息"""
        if not self.driver:
            self._log_info("浏览器未连接，无法获取元素", "red")
            return None
            
        try:
            # 获取当前鼠标屏幕位置
            import ctypes
            from ctypes import wintypes
            
            class POINT(ctypes.Structure):
                _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]
            
            pt = POINT()
            ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
            mouse_x, mouse_y = pt.x, pt.y
            
            # 获取浏览器窗口位置和大小
            window_pos = self.driver.get_window_position()
            window_size = self.driver.get_window_size()
            browser_left = window_pos['x']
            browser_top = window_pos['y']
            browser_right = browser_left + window_size['width']
            browser_bottom = browser_top + window_size['height']
            
            if not (browser_left <= mouse_x <= browser_right and browser_top <= mouse_y <= browser_bottom):
                self._log_info("鼠标不在浏览器窗口内", "red")
                return None
            
            # 动态获取内容区顶部偏移
            js = "return {screenY: window.screenY, outerHeight: window.outerHeight, innerHeight: window.innerHeight};"
            win_metrics = self.driver.execute_script(js)
            content_top = win_metrics['screenY'] + (win_metrics['outerHeight'] - win_metrics['innerHeight'])
            viewport_x = mouse_x - browser_left
            viewport_y = mouse_y - content_top
            
            # JavaScript脚本获取元素信息
            script = '''
            function getElementInfoWithPosition(x, y) {
                let el = document.elementFromPoint(x, y);
                
                if (!el) {
                    return { found: false, reason: "未找到元素" };
                }
                
                // 检查元素是否可见
                const rect = el.getBoundingClientRect();
                const isVisible = !!(rect.width && rect.height && 
                    window.getComputedStyle(el).visibility !== 'hidden' && 
                    window.getComputedStyle(el).display !== 'none');
                
                if (!isVisible) {
                    return { found: true, visible: false, reason: "元素存在但不可见" };
                }
                
                // 获取文本内容
                const text = el.innerText || el.textContent || "";
                
                // 生成XPath
                function getElementXPath(element) {
                    if (!element) return '';
                    
                    // 处理document节点
                    if (element.nodeType === 9) return '/';
                    
                    // 处理普通节点
                    var path = [];
                    var current = element;
                    
                    while (current && current.nodeType === 1) {
                        var index = 0;
                        var found = false;
                        
                        for (var i = 0; i < current.parentNode.childNodes.length; i++) {
                            var node = current.parentNode.childNodes[i];
                            if (node.nodeType === 1 && node.tagName === current.tagName) {
                                index++;
                            }
                            if (node === current) {
                                found = true;
                                break;
                            }
                        }
                        
                        var tagName = current.tagName.toLowerCase();
                        var pathIndex = (index > 1 || !found) ? '[' + index + ']' : '';
                        
                        if (current.id) {
                            path.unshift(tagName + '[@id="' + current.id + '"]');
                            break;
                        } else {
                            path.unshift(tagName + pathIndex);
                        }
                        
                        current = current.parentNode;
                    }
                    
                    return '/' + path.join('/');
                }
                
                // 获取元素位置和属性信息
                return {
                    found: true,
                    visible: true,
                    text: text,
                    xpath: getElementXPath(el),
                    y: rect.top,
                    x: rect.left,
                    height: rect.height,
                    width: rect.width,
                    tag: el.tagName,
                    class: el.className || ""
                };
            }
            
            return getElementInfoWithPosition(arguments[0], arguments[1]);
            '''
            
            # 执行脚本获取元素信息
            element_info = self.driver.execute_script(script, viewport_x, viewport_y)
            
            # 验证返回结果
            if not element_info:
                self._log_info("无法获取元素信息", "red")
                return None
                
            if not element_info.get('found', False):
                reason = element_info.get('reason', '未知原因')
                self._log_info(f"未找到元素: {reason}", "red")
                return None
                
            if not element_info.get('visible', False):
                self._log_info("元素存在但不可见", "red")
                return None
                
            return element_info
            
        except Exception as e:
            self._log_info(f"获取元素信息失败: {str(e)}", "red")
            import traceback
            self.logger.error(f"获取元素信息异常详情: {traceback.format_exc()}")
            return None
    
    def _start_distance_learning(self):
        """开始辅助定位模式"""
        try:
            # 重置状态变量
            self.is_distance_learning = True
            self.distance_learning_step = 0
            self.first_element_position = None
            self.second_element_position = None
            
            # 禁用开始和停止按钮，防止在学习过程中误操作
            self.start_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.DISABLED)
            
            self._log_info("开始辅助定位模式", "blue")
            self._log_info("请按 F1 键采集第一个订单的元素（建议选择订单号或其他唯一标识）", "green")
            self._log_info("然后按 F2 键采集第二个订单的相同位置元素", "green")
            self._log_info("系统将自动计算两个订单之间的滚动距离", "green")
            
        except Exception as e:
            self._log_info(f"启动辅助定位失败: {str(e)}", "red")
            import traceback
            self.logger.error(f"启动辅助定位异常: {traceback.format_exc()}")
    
    def _process_distance_learning(self, step):
        """处理辅助定位的步骤"""
        try:
            if not self.is_distance_learning:
                return
                
            if step == 0:  # F1键 - 采集第一个元素
                self._log_info("正在采集第一个订单的元素位置...", "blue")
                
                # 获取当前鼠标位置的元素
                element_info = self._get_element_at_cursor()
                if element_info:
                    xpath = element_info.get('xpath')
                    x_position = element_info.get('x')
                    y_position = element_info.get('y')
                    text = element_info.get('text', '')
                    tag_name = element_info.get('tag', '')
                    class_name = element_info.get('class', '')
                    
                    if xpath and x_position is not None and y_position is not None:
                        self.first_element_position = {
                            'xpath': xpath,
                            'x': x_position,
                            'y': y_position,
                            'text': text,
                            'tag': tag_name,
                            'class': class_name
                        }
                        
                        self.distance_learning_step = 1
                        self._log_info(f"第一个元素采集成功！", "green")
                        self._log_info(f"元素信息: 标签={tag_name}, 位置=({x_position}, {y_position}), 文本='{text[:30]}...'", "blue")
                        self._log_info("现在请滚动到下一个订单，然后按 F2 键采集第二个订单的相同位置元素", "green")
                    else:
                        self._log_info("无法获取元素的完整位置信息，请重试", "red")
                else:
                    self._log_info("无法获取当前鼠标位置的元素，请确保鼠标悬停在目标元素上", "red")
                    
            elif step == 1:  # F2键 - 采集第二个元素
                self._log_info("正在采集第二个订单的元素位置...", "blue")
                
                # 获取当前鼠标位置的元素
                element_info = self._get_element_at_cursor()
                if element_info:
                    xpath = element_info.get('xpath')
                    x_position = element_info.get('x')
                    y_position = element_info.get('y')
                    text = element_info.get('text', '')
                    tag_name = element_info.get('tag', '')
                    class_name = element_info.get('class', '')
                    
                    if xpath and x_position is not None and y_position is not None:
                        self.second_element_position = {
                            'xpath': xpath,
                            'x': x_position,
                            'y': y_position,
                            'text': text,
                            'tag': tag_name,
                            'class': class_name
                        }
                        
                        # 验证元素类型一致性
                        if self.first_element_position and self.second_element_position and self.first_element_position.get('tag') != self.second_element_position.get('tag'):
                            self._log_info(f"警告：两次采集的元素类型不同 (第一次: {self.first_element_position.get('tag')}, 第二次: {self.second_element_position.get('tag')})", "orange")
                            self._log_info("建议重新开始辅助定位，确保选择相同类型的元素", "orange")
                        
                        # 确保两个元素位置都有效
                        if self.first_element_position and self.second_element_position and 'y' in self.first_element_position and 'y' in self.second_element_position:
                            # 计算距离
                            distance = self.second_element_position['y'] - self.first_element_position['y']
                            
                            # 验证距离合理性
                            if distance <= 0:
                                self._log_info(f"警告：计算的滚动距离为负值或零 ({distance}px)，这可能不正确", "orange")
                                self._log_info("请确保第二个元素在第一个元素下方，建议重新开始辅助定位", "orange")
                                # 不保存这个不合理的距离值
                            elif distance > 1000:
                                self._log_info(f"警告：计算的滚动距离过大 ({distance}px)，这可能不正确", "orange")
                                self._log_info("请确保选择了相邻两个订单的元素，建议重新开始辅助定位", "orange")
                                # 不保存这个不合理的距离值
                            else:
                                self.scroll_distance = distance
                                self._log_info(f"辅助定位成功完成！两个订单之间的滚动距离为: {distance}px", "green")
                                self._log_info("现在您可以开始正常采集流程，系统将使用这个距离自动滚动到下一个订单", "green")
                        else:
                            self._log_info("错误：无法计算滚动距离，元素位置信息不完整", "red")
                            self._log_info("请重新开始辅助定位", "red")
                        
                        # 完成学习并恢复状态
                        self.is_distance_learning = False
                        self.distance_learning_step = 2
                        
                        # 恢复按钮状态
                        self.start_button.config(state=tk.NORMAL)
                        self.stop_button.config(state=tk.NORMAL)
                    else:
                        self._log_info("无法获取元素的完整位置信息，请重试", "red")
                else:
                    self._log_info("无法获取当前鼠标位置的元素，请确保鼠标悬停在目标元素上", "red")
                    
        except Exception as e:
            self._log_info(f"辅助定位过程中出错: {str(e)}", "red")
            import traceback
            self.logger.error(f"辅助定位异常: {traceback.format_exc()}")
            
            # 出错时重置状态
            self.is_distance_learning = False
            self.distance_learning_step = 0
            self.first_element_position = None
            self.second_element_position = None
            
            # 恢复按钮状态
            self.start_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.NORMAL)
        

    def _load_elements_from_json(self, file_path=None):
        """从JSON文件加载元素数据"""
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="选择元素配置文件",
                filetypes=[("JSON文件", "*.json")],
                initialdir=os.path.dirname(os.path.abspath(__file__))
            )
            
        if not file_path:
            return None
            
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                elements_data = json.load(f)
                
            # 处理并验证JSON数据
            processed_data = []
            for i, elem in enumerate(elements_data):
                # 转换为内部数据结构，添加默认顺序和启用状态
                processed_data.append({
                    "element_id": i,
                    "name": elem.get("custom_name", f"元素{i+1}"),
                    "action": elem.get("action", "getText"),
                    "xpath": elem.get("xpath", ""),
                    "order": i+1,  # 默认顺序为加载顺序
                    "enabled": True  # 默认启用
                })
            return processed_data
        except Exception as e:
            self._log_info(f"加载元素数据失败: {str(e)}", "red")
            return None
    

    def _configure_operations(self):
        """显示操作选择与排序对话框"""
        if not self.is_browser_connected:
            self._log_info("浏览器未连接，无法配置操作", "red")
            return
            
        elements_data = self._load_elements_from_json()
        if not elements_data:
            self._log_info("未找到有效的元素数据，请先使用采集模式收集元素", "red")
            return
            
        dialog = OperationSequenceDialog(self.root, elements_data)
        if dialog.result:
            self.operation_sequence = dialog.result
            # 按顺序排序
            self.operation_sequence.sort(key=lambda x: x["order"])
            # 仅保留启用的操作
            self.operation_sequence = [op for op in self.operation_sequence if op["enabled"]]
            self._log_info(f"已配置{len(self.operation_sequence)}个操作，准备就绪", "green")
            # 启用开始按钮
            self.start_button.config(state=tk.NORMAL)
            

    def _preview_element(self, xpath):
        """高亮显示指定XPath的元素，用于预览"""
        if not self.driver:
            self._log_info("浏览器未连接，无法预览元素", "red")
            return False
            
        try:
            # 查找元素
            element = self.driver.find_element(By.XPATH, xpath)
            
            # 使用JavaScript高亮元素
            script = """
            function highlightElement(element) {
                var originalStyle = element.getAttribute('style');
                element.setAttribute('style', originalStyle + '; border: 2px solid red !important; background-color: yellow !important;');
                setTimeout(function() {
                    element.setAttribute('style', originalStyle);
                }, 2000);
            }
            arguments[0].scrollIntoView({behavior: "smooth", block: "center"});
            highlightElement(arguments[0]);
            """
            self.driver.execute_script(script, element)
            return True
        except Exception as e:
            self._log_info(f"预览元素失败: {str(e)}", "red")
            return False


    def _collect_ref1_xpath(self):
        self._log_info("请将鼠标悬停在第1个订单的参照元素上，然后按'.'键采集", "blue")
        self._pending_collect = 'ref1'


    def _collect_ref2_xpath(self):
        self._log_info("请将鼠标悬停在第2个订单的参照元素上，然后按'.'键采集", "blue")
        self._pending_collect = 'ref2'


    def _collect_scroll_container_xpath(self):
        self._log_info("请将鼠标悬停在滚动容器上，然后按'.'键采集", "blue")
        self._pending_collect = 'scroll_container'


    def _start_collection(self):
        # 启动验证码检测（如果已配置）
        if hasattr(self, 'template_images') and (self.template_images or self.use_mask_detection):
            if hasattr(self, 'target_window_handle') and self.target_window_handle:
                self._start_captcha_detection()
                self._log_info("验证码检测已自动启动", "green")
            else:
                self._log_info("提示：未选择验证码检测目标窗口，验证码检测未启动", "orange")
        
        mode = self.collection_mode.get()
        if mode == "正常模式":
            if not hasattr(self, 'operation_sequence') or not self.operation_sequence:
                self._configure_operations()
                return
                
            # 确保order_clipboard_contents字典已初始化
            if not hasattr(self, 'order_clipboard_contents'):
                self.order_clipboard_contents = {}
                self._log_info("初始化订单ID与收货信息映射字典", "blue")
                print("DEBUG-START-COLLECTION: 初始化订单ID与收货信息映射字典")
            
            # 程序启动时已清空所有历史数据，无需再次清理
                
            # 清空并记录初始剪贴板状态
            self.initial_clipboard_content = self._get_clipboard_content()
            self._log_info(f"记录初始剪贴板状态，长度: {len(self.initial_clipboard_content) if self.initial_clipboard_content else 0}", "blue")
            
            # 启动剪贴板监听
            self._start_clipboard_monitor()
            
            # 重置订单ID缓存
            self.current_order_id = None
            
            # 在主线程中预先获取订单数量（如果需要手动输入）
            order_count_elements = [op for op in self.operation_sequence if op.get("is_order_count", False)]
            manual_order_count = None
            
            if not order_count_elements:
                # 没有订单数量元素，需要手动输入订单数量
                self._log_info("未选择订单数量元素，将手动输入订单数量", "blue")
                try:
                    manual_order_count = self._get_order_count()
                    if manual_order_count == 0:
                        self._log_info('订单数量为0，无需执行。')
                        return
                    self._log_info(f"手动输入的订单数量: {manual_order_count}", "blue")
                except Exception as e:
                    self._log_info(f"获取订单数量失败: {str(e)}", "red")
                    return
            
            self._log_info("开始自动采集流程...", "green")
            
            # 缩小窗口到左上角
            self._shrink_window_to_corner()
            
            self.is_running = True
            self.is_paused = False
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL)
            self.continue_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.configure_button.config(state=tk.DISABLED)
            threading.Thread(target=self.run_actions_loop, args=(manual_order_count,), daemon=True).start()
        elif mode == "采集模式":
            # 确保order_clipboard_contents字典已初始化
            if not hasattr(self, 'order_clipboard_contents'):
                self.order_clipboard_contents = {}
                self._log_info("初始化订单ID与收货信息映射字典", "blue")
                print("DEBUG-START-COLLECTION-MODE: 初始化订单ID与收货信息映射字典")
            
            # 启动剪贴板监听
            self._start_clipboard_monitor()
            
            # 重置订单ID缓存
            self.current_order_id = None
            
            self._log_info("进入采集准备状态，请按'.'键采集当前元素...", "green")
            
            # 缩小窗口到左上角
            self._shrink_window_to_corner()
            
            self.is_running = True
            self.is_paused = False
            self.start_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL)
            self.continue_button.config(state=tk.DISABLED)
            self.stop_button.config(state=tk.NORMAL)
            self.configure_button.config(state=tk.DISABLED)
    
    def _pause_collection(self):
        """暂停采集流程"""
        self._log_info("已暂停采集", "orange")
        self.is_paused = True
        
        # 更新按钮状态
        self.pause_button.config(state=tk.DISABLED)
        self.continue_button.config(state=tk.NORMAL)
    
    def _continue_collection(self):
        """继续采集流程"""
        self._log_info("继续采集", "green")
        self.is_paused = False
        
        # 更新按钮状态
        self.pause_button.config(state=tk.NORMAL)
        self.continue_button.config(state=tk.DISABLED)
    
    def _stop_collection(self):
        if not self.is_running:
            return
        self._log_info("已终止采集", "red")
        self.is_running = False
        self.is_paused = False
        
        # 停止验证码检测
        if hasattr(self, 'captcha_running') and self.captcha_running:
            self._stop_captcha_detection()
            self._log_info("验证码检测已停止", "orange")
        
        # 停止剪贴板监听
        self._stop_clipboard_monitor()
        
        # 保存当前的剪贴板映射
        self._save_clipboard_mappings()
        
        # 删除辅助定位相关状态重置
        # 更新按钮状态
        self.start_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.continue_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.DISABLED)
        self.configure_button.config(state=tk.NORMAL if self.is_browser_connected else tk.DISABLED)
        # 如果有数据，根据模式启用不同导出按钮
        if hasattr(self, 'collected_data') and len(self.collected_data) > 0:
            if self.collection_mode.get() == "正常模式":
                self.excel_button.config(state=tk.NORMAL)
                self.word_button.config(state=tk.NORMAL)
            else:
                self.json_button.config(state=tk.NORMAL)

    def _learn_xpath_pattern(self, xpath1, xpath2):
        """参考脚本的循环XPath推算：找到唯一递增索引位置，允许中间层级差异"""
        import re
        
        # 完全相同的XPath，无法学习
        if xpath1 == xpath2:
            self._log_info('两个参照XPath完全相同，无法学习循环模式。', 'red')
            return None
            
        parts1 = xpath1.split('/')
        parts2 = xpath2.split('/')
        
        # 处理不同长度的XPath
        if len(parts1) != len(parts2):
            self._log_info(f'参照XPath层级不同: 第1个有{len(parts1)}层，第2个有{len(parts2)}层', 'orange')
            # 尝试从后向前匹配，找到共同部分
            min_len = min(len(parts1), len(parts2))
            parts1 = parts1[-min_len:]
            parts2 = parts2[-min_len:]
            self._log_info(f'尝试使用后{min_len}层进行匹配', 'blue')
            
        # 查找所有不同的部分
        diff_segments = []
        for i in range(len(parts1)):
            if parts1[i] != parts2[i]:
                diff_segments.append((i, parts1[i], parts2[i]))
                
        # 没有差异，不应该发生，因为之前已经检查了完全相同的情况
        if not diff_segments:
            self._log_info('截取后的XPath没有差异，无法学习。', 'red')
            return None
            
        # 筛选出包含数字索引的差异部分
        index_diffs = []
        for idx, part1, part2 in diff_segments:
            match1 = re.search(r'\[(\d+)\]', part1)
            match2 = re.search(r'\[(\d+)\]', part2)
            if match1 and match2:
                index1 = int(match1.group(1))
                index2 = int(match2.group(1))
                if index2 > index1:  # 确保索引是递增的
                    index_diffs.append((idx, index1, index2, part1, part2))
                    
        # 如果没有找到递增的数字索引差异
        if not index_diffs:
            # 尝试找出任何包含数字索引的部分
            for idx, part1, part2 in diff_segments:
                match1 = re.search(r'\[(\d+)\]', part1)
                match2 = re.search(r'\[(\d+)\]', part2)
                if match1 or match2:  # 只要有一个包含索引
                    if match1 and not match2:
                        self._log_info(f'XPath差异: 第1个有索引[{match1.group(1)}]，第2个没有索引', 'orange')
                    elif match2 and not match1:
                        self._log_info(f'XPath差异: 第1个没有索引，第2个有索引[{match2.group(1)}]', 'orange')
                    else:
                        self._log_info(f'XPath索引不是递增的: {match1.group(1)} -> {match2.group(1)}', 'orange')
            
            self._log_info('未找到有效的递增索引模式，尝试使用最后一个数字索引作为模式。', 'orange')
            
            # 查找最后一个包含数字索引的部分
            last_index_part = None
            last_index_value = None
            last_index_idx = -1
            
            for i, part in enumerate(parts1):
                match = re.search(r'\[(\d+)\]', part)
                if match:
                    last_index_part = part
                    last_index_value = int(match.group(1))
                    last_index_idx = i
                    
            if last_index_part:
                self._log_info(f'使用最后一个索引 [{last_index_value}] 作为基准', 'blue')
                template = re.sub(r'\[\d+\]', '[{}]', last_index_part, count=1)
                return {'diff_segment_index': last_index_idx, 'template': template, 'start_index': last_index_value}
            else:
                self._log_info('XPath中没有找到任何数字索引，无法学习。', 'red')
                return None
                
        # 如果找到了多个递增索引差异，使用最可能的那个（通常是第一个）
        if len(index_diffs) > 1:
            self._log_info(f'发现{len(index_diffs)}处递增索引差异，使用第一处作为模式。', 'orange')
            
        # 使用找到的差异生成模式
        diff_idx, start_index, next_index, part1, part2 = index_diffs[0]
        template = re.sub(r'\[\d+\]', '[{}]', part1, count=1)
        self._log_info(f'已学习到XPath模式: 在第{diff_idx}段，索引从{start_index}递增到{next_index}', 'green')
        
        return {'diff_segment_index': diff_idx, 'template': template, 'start_index': start_index}


    def _generate_xpath_for_item(self, base_xpath, loop_counter, pattern):
        """参考脚本的循环XPath生成"""
        import re
        parts = base_xpath.split('/')
        if pattern and 'diff_segment_index' in pattern:
            idx = pattern['diff_segment_index']
            if len(parts) > idx:
                actual_dom_index = pattern['start_index'] + loop_counter - 1
                original_segment = parts[idx]
                if re.search(r'\[\d+\]', original_segment):
                    parts[idx] = re.sub(r'\[\d+\]', f'[{actual_dom_index}]', original_segment, count=1)
                    return '/'.join(parts)
        return base_xpath

