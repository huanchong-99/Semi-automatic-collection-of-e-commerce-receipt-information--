#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
翻页功能相关

从原始main.py文件中拆分出来的模块
"""

from utils import *

class PageTurner:
    """翻页功能相关"""
    
    def __init__(self):
        pass

    def collect_page_turn_element(self):
        """采集翻页元素"""
        try:
            if not self.driver:
                self._log_info("请先连接浏览器", "error")
                return
            
            # 设置翻页采集模式标志
            self.collecting_page_turn = True
            
            self._log_info("请将鼠标悬停在翻页按钮上，然后按\".\"键进行采集...", "blue")
            
            # 更新按钮状态
            if hasattr(self, 'collect_page_btn'):
                self.collect_page_btn.config(text="等待采集中...", state=tk.DISABLED)
                
        except Exception as e:
            self._log_info(f"启动翻页元素采集出错: {e}", "error")
            messagebox.showerror("错误", f"启动翻页元素采集出错: {e}")
    

    def _handle_page_turn_collection(self):
        """处理翻页元素采集 - 使用text.py中成功的悬停监听机制"""
        try:
            if not self.driver:
                self._log_info("浏览器未连接，无法采集", "red")
                messagebox.showerror("错误", "浏览器未连接，无法采集")
                # 恢复按钮状态
                if hasattr(self, 'collect_page_btn'):
                    self.collect_page_btn.config(text="采集翻页元素", state=tk.NORMAL)
                # 清除翻页采集模式标志
                self.collecting_page_turn = False
                return
            
            self._log_info("正在采集翻页按钮XPath...", "blue")
            
            # 使用悬停监听获取元素XPath
            self._log_info("正在通过悬停监听获取元素XPath...", "blue")
            xpath = self._get_hovered_xpath_recursive()
            
            if not xpath:
                self._log_info("✗ 未检测到悬停元素", "red")
                messagebox.showerror("采集失败", "未能获取到元素信息，请确保鼠标悬停在浏览器页面内的有效元素上。")
                # 恢复按钮状态
                if hasattr(self, 'collect_page_btn'):
                    self.collect_page_btn.config(text="采集翻页元素", state=tk.NORMAL)
                # 清除翻页采集模式标志
                self.collecting_page_turn = False
                return
                
            self._log_info("✓ 成功通过悬停监听方式定位到元素", "green")
            
            # 获取元素文本和元素对象
            element = None
            element_text = "未知元素"
            
            try:
                element = self.driver.find_element(By.XPATH, xpath)
                element_text = element.text or element.get_attribute('title') or element.get_attribute('aria-label') or element.tag_name
                self._log_info(f"成功获取元素: {element_text}")
                                    
            except Exception as e:
                self._log_info(f"无法通过XPath找到元素: {str(e)}", "orange")
                # 即使找不到元素，我们仍然保存XPath，因为它可能在点击时有效
                element_text = "SVG元素或动态元素"
            
            self._log_info(f"最终XPath: {xpath}")
            
            # 设置为翻页按钮XPath
            self.next_page_xpath = xpath
            self.next_page_collected = True
            
            # 高亮显示采集到的元素
            if element:
                try:
                    self._highlight_element_with_script(element)
                    self._log_info("元素高亮显示成功", "green")
                except Exception as e:
                    self._log_info(f"无法高亮显示元素: {str(e)}", "orange")
            else:
                self._log_info("跳过高亮显示：元素可能是SVG或动态元素", "orange")
            
            self._log_info(f"翻页元素采集成功: {xpath}", "green")
            
            # 更新状态显示并保存配置
            if hasattr(self, '_update_element_status_display'):
                self._update_element_status_display()
            if hasattr(self, '_save_element_config'):
                self._save_element_config()
            
            # 恢复按钮状态
            if hasattr(self, 'collect_page_btn'):
                self.collect_page_btn.config(text="重新采集翻页元素", state=tk.NORMAL)
            
            messagebox.showinfo("成功", f"翻页元素采集成功！\n元素: {element_text}")
            
            # 清除翻页采集模式标志
            self.collecting_page_turn = False
                
        except Exception as e:
            self._log_info(f"采集翻页元素出错: {e}", "error")
            messagebox.showerror("错误", f"采集翻页元素出错: {e}")
            
            # 恢复按钮状态
            if hasattr(self, 'collect_page_btn'):
                self.collect_page_btn.config(text="采集翻页元素", state=tk.NORMAL)
            
            # 清除翻页采集模式标志
            self.collecting_page_turn = False
    

    def _highlight_element_with_script(self, element):
        """使用JavaScript高亮显示元素"""
        try:
            script = """
            var element = arguments[0];
            var originalStyle = element.getAttribute('style') || '';
            element.setAttribute('style', originalStyle + '; border: 3px solid red !important; background-color: yellow !important;');
            setTimeout(function() {
                element.setAttribute('style', originalStyle);
            }, 3000);
            """
            self.driver.execute_script(script, element)
        except Exception as e:
            self._log_info(f"高亮显示元素失败: {str(e)}", "orange")
    

    def _get_hovered_xpath(self, mouse_x, mouse_y):
        """获取鼠标悬停位置的元素XPath"""
        try:
            # 将屏幕坐标转换为浏览器内坐标
            browser_x, browser_y = self._screen_to_browser_coords(mouse_x, mouse_y)
            
            if browser_x is None or browser_y is None:
                return None
            
            # 执行JavaScript获取元素
            script = f"""
            var element = document.elementFromPoint({browser_x}, {browser_y});
            if (!element) return null;
            
            function getXPath(element) {{
                if (element.id !== '') {{
                    return '//*[@id="' + element.id + '"]';
                }}
                if (element === document.body) {{
                    return '/html/body';
                }}
                
                var ix = 0;
                var siblings = element.parentNode.childNodes;
                for (var i = 0; i < siblings.length; i++) {{
                    var sibling = siblings[i];
                    if (sibling === element) {{
                        return getXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                    }}
                    if (sibling.nodeType === 1 && sibling.tagName === element.tagName) {{
                        ix++;
                    }}
                }}
            }}
            
            return getXPath(element);
            """
            
            xpath = self.driver.execute_script(script)
            return xpath
            
        except Exception as e:
            self._log_info(f"获取元素XPath失败: {e}", "error")
            return None
    

    def _screen_to_browser_coords(self, screen_x, screen_y):
        """将屏幕坐标转换为浏览器内坐标"""
        try:
            if not self.target_window_handle:
                return None, None
            
            # 获取窗口位置
            rect = win32gui.GetWindowRect(self.target_window_handle)
            window_x, window_y, window_right, window_bottom = rect
            
            # 计算浏览器内坐标
            browser_x = screen_x - window_x
            browser_y = screen_y - window_y
            
            # 考虑浏览器标题栏和工具栏的偏移
            browser_y -= 100  # 大概的标题栏和工具栏高度
            
            return browser_x, browser_y
            
        except Exception as e:
            self._log_info(f"坐标转换失败: {e}", "error")
            return None, None
    

    def _highlight_element(self, xpath):
        """高亮显示元素"""
        try:
            script = f"""
            var element = document.evaluate('{xpath}', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
            if (element) {{
                element.style.border = '3px solid red';
                element.style.backgroundColor = 'yellow';
                element.style.opacity = '0.8';
                
                setTimeout(function() {{
                    element.style.border = '';
                    element.style.backgroundColor = '';
                    element.style.opacity = '';
                }}, 3000);
            }}
            """
            
            self.driver.execute_script(script)
            
        except Exception as e:
            self._log_info(f"高亮元素失败: {e}", "error")
    

    def check_page_turn_needed(self):
        """检查是否需要翻页"""
        try:
            # 如果启用了模块化翻页，禁用原有翻页检查
            if hasattr(self, 'use_modular_paging') and self.use_modular_paging:
                return False  # 模块化翻页模式下，翻页由外层循环控制
            
            # 获取当前翻页页数设置
            if hasattr(self, 'page_count_var'):
                self.target_page_count = int(self.page_count_var.get())
            
            # 检查是否达到翻页条件
            if hasattr(self, 'current_order_index') and self.current_order_index > 0:
                if (self.current_order_index + 1) % self.target_page_count == 0:
                    return True
            
            return False
            
        except Exception as e:
            self._log_info(f"检查翻页条件失败: {e}", "error")
            return False
    

    def execute_page_turn(self):
        """执行翻页操作"""
        try:
            if not self.next_page_collected or not self.next_page_xpath:
                self._log_info("请先采集翻页元素", "error")
                return False
            
            self._log_info("开始执行翻页操作...", "blue")
            
            # 翻页前截图
            before_screenshot = self._save_page_screenshot("before")
            if not before_screenshot:
                self._log_info("翻页前截图失败", "error")
                return False
            
            # 检查翻页按钮是否存在
            try:
                element = self.driver.find_element(By.XPATH, self.next_page_xpath)
                if not element.is_enabled():
                    self._log_info("翻页按钮不可用", "error")
                    return False
            except Exception as e:
                self._log_info(f"找不到翻页按钮: {e}", "error")
                return False
            
            # 点击翻页按钮
            try:
                element.click()
                self._log_info("翻页按钮点击成功", "green")
            except Exception as e:
                self._log_info(f"点击翻页按钮失败: {e}", "error")
                return False
            
            # 等待页面加载
            time.sleep(3)
            
            # 翻页后截图
            after_screenshot = self._save_page_screenshot("after")
            if not after_screenshot:
                self._log_info("翻页后截图失败", "error")
                return False
            
            # 比较截图验证翻页是否成功
            if self._compare_screenshots(before_screenshot, after_screenshot):
                self.page_turn_count += 1
                self._log_info(f"翻页成功！当前已翻页 {self.page_turn_count} 次", "green")
                
                # 清理截图文件
                self._clear_screenshots()
                
                return True
            else:
                self._log_info("翻页失败：页面内容未发生变化", "error")
                return False
                
        except Exception as e:
            self._log_info(f"翻页操作失败: {e}", "error")
            return False
    

    def _save_page_screenshot(self, prefix):
        """保存页面截图"""
        try:
            if not self.target_window_handle:
                return None
            
            screenshot = self._capture_window_screenshot()
            if screenshot is None:
                return None
            
            timestamp = int(time.time() * 1000)
            filename = f"{prefix}_{timestamp}.png"
            filepath = os.path.join(self.screenshot_dir, filename)
            
            cv2.imwrite(filepath, screenshot)
            return filepath
            
        except Exception as e:
            self._log_info(f"保存截图失败: {e}", "error")
            return None
    

    def _compare_screenshots(self, before_path, after_path):
        """比较两张截图是否不同"""
        try:
            # 读取图片
            before_img = cv2.imread(before_path)
            after_img = cv2.imread(after_path)
            
            if before_img is None or after_img is None:
                return False
            
            # 确保图片尺寸相同
            if before_img.shape != after_img.shape:
                return True  # 尺寸不同说明页面发生了变化
            
            # 计算图片差异
            diff = cv2.absdiff(before_img, after_img)
            gray_diff = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
            
            # 计算非零像素的比例
            non_zero_count = cv2.countNonZero(gray_diff)
            total_pixels = gray_diff.shape[0] * gray_diff.shape[1]
            diff_ratio = non_zero_count / total_pixels
            
            # 如果差异比例大于阈值，认为翻页成功
            threshold = 0.1  # 10%的像素发生变化
            return diff_ratio > threshold
            
        except Exception as e:
            self._log_info(f"比较截图失败: {e}", "error")
            return False
    

    def _capture_window_screenshot(self):
        """截取目标窗口截图"""
        try:
            import win32gui, win32ui, win32con
            from PIL import Image
            import numpy as np
            
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
    
    def _clear_screenshots(self):
        """清理截图文件"""
        try:
            for filename in os.listdir(self.screenshot_dir):
                if filename.endswith('.png'):
                    filepath = os.path.join(self.screenshot_dir, filename)
                    os.remove(filepath)
        except Exception as e:
            self._log_info(f"清理截图文件失败: {e}", "error")
    

    def on_page_count_changed(self, *args):
        """翻页页数改变时的回调"""
        try:
            if hasattr(self, 'page_count_var'):
                self.target_page_count = int(self.page_count_var.get())
                self._log_info(f"翻页页数设置为: {self.target_page_count}", "blue")
        except Exception as e:
            self._log_info(f"设置翻页页数失败: {e}", "error")

