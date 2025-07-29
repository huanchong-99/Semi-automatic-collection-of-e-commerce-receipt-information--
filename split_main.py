#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
插桩脚本：将超大的main.py文件按模块拆分

这个脚本会将7119行的main.py文件拆分成以下模块：
1. main.py - 主应用程序入口
2. ui_components.py - UI组件和界面相关
3. browser_controller.py - 浏览器控制相关
4. element_collector.py - 元素采集相关
5. data_processor.py - 数据处理和导出相关
6. clipboard_manager.py - 剪贴板管理相关
7. captcha_detector.py - 验证码检测相关
8. page_turner.py - 翻页功能相关
9. config_manager.py - 配置管理相关
10. utils.py - 工具函数和常量
"""

import os
import re
import shutil
from datetime import datetime

class MainSplitter:
    def __init__(self, source_file="f:\\shoujifahuoxinxi\\main_original.py"):
        self.source_file = source_file
        self.output_dir = os.path.dirname(source_file)
        self.backup_dir = os.path.join(self.output_dir, "backup")
        
        # 读取源文件内容
        with open(source_file, 'r', encoding='utf-8') as f:
            self.content = f.read()
            self.lines = self.content.split('\n')
        
        print(f"源文件总行数: {len(self.lines)}")
    
    def find_function_boundaries(self, function_name):
        """查找函数的边界"""
        start_line = None
        end_line = None
        indent_level = None
        
        for i, line in enumerate(self.lines):
            # 查找函数定义
            if re.match(rf'\s*def\s+{re.escape(function_name)}\s*\(', line):
                start_line = i
                indent_level = len(line) - len(line.lstrip())
                continue
            
            # 如果找到了函数开始，查找结束
            if start_line is not None:
                current_indent = len(line) - len(line.lstrip()) if line.strip() else float('inf')
                
                # 如果遇到同级或更低级的缩进（且不是空行），函数结束
                if line.strip() and current_indent <= indent_level:
                    end_line = i - 1
                    break
        
        # 如果没找到结束，说明是最后一个函数
        if start_line is not None and end_line is None:
            end_line = len(self.lines) - 1
        
        return start_line, end_line
    
    def find_class_boundaries(self):
        """查找ShippingInfoCollector类的边界"""
        start_line = None
        end_line = None
        
        for i, line in enumerate(self.lines):
            if re.match(r'^class\s+ShippingInfoCollector\s*[\(:]', line):
                start_line = i
                break
        
        if start_line is not None:
            # 查找类的结束（下一个同级定义或main函数）
            for i in range(start_line + 1, len(self.lines)):
                line = self.lines[i]
                if line.strip() and not line.startswith((' ', '\t')):
                    if line.startswith('def main()'):
                        end_line = i - 1
                        break
                    elif line.startswith('class ') or line.startswith('def '):
                        end_line = i - 1
                        break
            
            if end_line is None:
                # 查找main函数的位置
                for i in range(start_line + 1, len(self.lines)):
                    if self.lines[i].startswith('def main()'):
                        end_line = i - 1
                        break
                
                if end_line is None:
                    end_line = len(self.lines) - 1
        
        return start_line, end_line
    
    def extract_imports_and_constants(self):
        """提取导入语句和常量定义"""
        imports = []
        constants = []
        functions = []
        
        # 提取文件开头的导入和函数
        in_imports = True
        i = 0
        
        while i < len(self.lines) and i < 500:  # 只检查前500行
            line = self.lines[i].strip()
            
            # 跳过注释和空行
            if line.startswith('#') or line == '' or line.startswith('"""') or line.startswith("'''"):
                i += 1
                continue
            
            # 导入语句
            if line.startswith(('import ', 'from ')) and in_imports:
                imports.append(self.lines[i])
                i += 1
                continue
            
            # 函数定义
            if line.startswith('def '):
                func_start = i
                func_name = re.match(r'def\s+(\w+)\s*\(', line)
                if func_name:
                    func_name = func_name.group(1)
                    # 查找函数结束
                    func_end = self.find_function_end(i)
                    if func_end:
                        functions.append({
                            'name': func_name,
                            'start': func_start,
                            'end': func_end,
                            'code': '\n'.join(self.lines[func_start:func_end+1])
                        })
                        i = func_end + 1
                        continue
            
            # 常量定义（全大写变量或特殊变量）
            if '=' in line and not line.startswith(('import ', 'from ', 'class ', 'def ')):
                var_name = line.split('=')[0].strip()
                if var_name.isupper() or 'dependencies_ok' in var_name:
                    constants.append(self.lines[i])
                    i += 1
                    continue
            
            # 如果遇到class定义，停止
            if line.startswith('class '):
                break
            
            in_imports = False
            i += 1
        
        return imports, constants, functions
    
    def find_function_end(self, start_line):
        """查找函数的结束行"""
        if start_line >= len(self.lines):
            return None
        
        start_indent = len(self.lines[start_line]) - len(self.lines[start_line].lstrip())
        
        for i in range(start_line + 1, len(self.lines)):
            line = self.lines[i]
            if line.strip() == '':
                continue
            
            current_indent = len(line) - len(line.lstrip())
            if current_indent <= start_indent:
                return i - 1
        
        return len(self.lines) - 1
    
    def extract_class_methods(self):
        """提取类中的所有方法"""
        class_start, class_end = self.find_class_boundaries()
        if class_start is None:
            return {}
        
        methods = {}
        i = class_start + 1
        
        while i <= class_end:
            line = self.lines[i].strip()
            
            # 查找方法定义
            if re.match(r'def\s+\w+\s*\(', line):
                method_match = re.match(r'def\s+(\w+)\s*\(', line)
                if method_match:
                    method_name = method_match.group(1)
                    method_start = i
                    method_end = self.find_method_end(i, class_end)
                    
                    if method_end:
                        methods[method_name] = {
                            'start': method_start,
                            'end': method_end,
                            'code': '\n'.join(self.lines[method_start:method_end+1])
                        }
                        i = method_end + 1
                    else:
                        i += 1
                else:
                    i += 1
            else:
                i += 1
        
        return methods
    
    def find_method_end(self, start_line, class_end):
        """查找方法的结束行"""
        if start_line >= len(self.lines):
            return None
        
        start_indent = len(self.lines[start_line]) - len(self.lines[start_line].lstrip())
        
        for i in range(start_line + 1, min(class_end + 1, len(self.lines))):
            line = self.lines[i]
            if line.strip() == '':
                continue
            
            current_indent = len(line) - len(line.lstrip())
            if current_indent <= start_indent:
                return i - 1
        
        return class_end
    
    def create_utils_module(self):
        """创建utils.py模块"""
        imports, constants, functions = self.extract_imports_and_constants()
        
        content = []
        content.append('#!/usr/bin/env python3')
        content.append('# -*- coding: utf-8 -*-')
        content.append('"""')
        content.append('工具函数和常量')
        content.append('')
        content.append('从原始main.py文件中拆分出来的模块')
        content.append('"""')
        content.append('')
        
        # 添加导入语句
        content.extend(imports)
        content.append('')
        
        # 添加工具函数
        for func in functions:
            if func['name'] in ['init_tcl_compatibility', 'install_package', 'auto_install_dependencies', 'check_and_install_dependencies']:
                content.append(func['code'])
                content.append('')
        
        # 添加常量
        content.extend(constants)
        content.append('')
        
        return '\n'.join(content)
    
    def create_module_with_methods(self, module_name, method_names, description):
        """创建包含指定方法的模块"""
        methods = self.extract_class_methods()
        
        content = []
        content.append('#!/usr/bin/env python3')
        content.append('# -*- coding: utf-8 -*-')
        content.append('"""')
        content.append(description)
        content.append('')
        content.append('从原始main.py文件中拆分出来的模块')
        content.append('"""')
        content.append('')
        
        # 添加导入
        content.append('from utils import *')
        content.append('')
        
        # 添加类定义
        class_name = self.get_class_name_for_module(module_name)
        content.append(f'class {class_name}:')
        content.append(f'    """{description}"""')
        content.append('')
        
        # 添加初始化方法
        if '__init__' in method_names and '__init__' in methods:
            # 修改__init__方法为模块特定的初始化
            init_code = methods['__init__']['code']
            # 简化初始化方法
            content.append('    def __init__(self, *args, **kwargs):')
            content.append(f'        """初始化{description}"""')
            content.append('        pass')
            content.append('')
        
        # 添加指定的方法
        for method_name in method_names:
            if method_name in methods and method_name != '__init__':
                method_code = methods[method_name]['code']
                content.append(method_code)
                content.append('')
        
        return '\n'.join(content)
    
    def get_class_name_for_module(self, module_name):
        """根据模块名生成类名"""
        name_map = {
            'ui_components.py': 'UIComponents',
            'browser_controller.py': 'BrowserController',
            'element_collector.py': 'ElementCollector',
            'data_processor.py': 'DataProcessor',
            'clipboard_manager.py': 'ClipboardManager',
            'captcha_detector.py': 'CaptchaDetector',
            'page_turner.py': 'PageTurner',
            'config_manager.py': 'ConfigManager'
        }
        return name_map.get(module_name, 'BaseComponent')
    
    def create_main_module(self):
        """创建主模块"""
        content = []
        content.append('#!/usr/bin/env python3')
        content.append('# -*- coding: utf-8 -*-')
        content.append('"""')
        content.append('主应用程序入口')
        content.append('')
        content.append('整合所有模块的主应用程序')
        content.append('"""')
        content.append('')
        
        # 导入所有模块
        content.append('from utils import *')
        content.append('from ui_components import UIComponents')
        content.append('from browser_controller import BrowserController')
        content.append('from element_collector import ElementCollector')
        content.append('from data_processor import DataProcessor')
        content.append('from clipboard_manager import ClipboardManager')
        content.append('from captcha_detector import CaptchaDetector')
        content.append('from page_turner import PageTurner')
        content.append('from config_manager import ConfigManager')
        content.append('')
        
        # 主类定义
        content.append('class ShippingInfoCollector(')
        content.append('    UIComponents,')
        content.append('    BrowserController,')
        content.append('    ElementCollector,')
        content.append('    DataProcessor,')
        content.append('    ClipboardManager,')
        content.append('    CaptchaDetector,')
        content.append('    PageTurner,')
        content.append('    ConfigManager')
        content.append('):')
        content.append('    """收货信息采集工具主类"""')
        content.append('    ')
        content.append('    def __init__(self, root):')
        content.append('        """初始化应用程序"""')
        content.append('        self.root = root')
        content.append('        ')
        content.append('        # 初始化基本属性')
        content.append('        self._init_basic_attributes()')
        content.append('        ')
        content.append('        # 初始化所有组件')
        content.append('        UIComponents.__init__(self, root)')
        content.append('        BrowserController.__init__(self)')
        content.append('        ElementCollector.__init__(self)')
        content.append('        DataProcessor.__init__(self)')
        content.append('        ClipboardManager.__init__(self)')
        content.append('        CaptchaDetector.__init__(self)')
        content.append('        PageTurner.__init__(self)')
        content.append('        ConfigManager.__init__(self)')
        content.append('')
        content.append('    def _init_basic_attributes(self):')
        content.append('        """初始化基本属性"""')
        content.append('        # 从原始__init__方法中提取的基本属性初始化')
        
        # 提取原始__init__方法中的属性初始化
        methods = self.extract_class_methods()
        if '__init__' in methods:
            init_lines = methods['__init__']['code'].split('\n')
            for line in init_lines:
                if 'self.' in line and '=' in line and not line.strip().startswith('#'):
                    content.append('        ' + line.strip())
        
        content.append('')
        
        # 添加main函数
        main_start, main_end = self.find_function_boundaries('main')
        if main_start is not None and main_end is not None:
            main_code = '\n'.join(self.lines[main_start:main_end+1])
            content.append(main_code)
        
        content.append('')
        content.append('if __name__ == "__main__":')
        content.append('    try:')
        content.append('        main()')
        content.append('    except Exception as e:')
        content.append('        # 捕获并记录未处理的异常')
        content.append('        error_msg = f"程序发生未处理的异常: {str(e)}\\n{traceback.format_exc()}"')
        content.append('        logging.error(error_msg)')
        content.append('        try:')
        content.append('            messagebox.showerror("错误", f"程序发生错误:\\n{str(e)}\\n\\n详细信息已记录到日志文件中。")')
        content.append('        except Exception:')
        content.append('            # 如果messagebox也失败，至少确保错误被记录')
        content.append('            print(f"程序发生严重错误: {str(e)}\\n详细信息已记录到日志文件中。")')
        
        return '\n'.join(content)
    
    def split_modules(self):
        """执行模块拆分"""
        print("开始重新拆分main.py文件...")
        print(f"源文件: {self.source_file}")
        print(f"输出目录: {self.output_dir}")
        print()
        
        # 定义模块和对应的方法
        modules = {
            'utils.py': {
                'description': '工具函数和常量',
                'methods': [],
                'create_func': self.create_utils_module
            },
            'ui_components.py': {
                'description': 'UI组件和界面相关',
                'methods': [
                    '__init__', '_create_ui', '_create_menu', '_create_main_frame',
                    '_create_browser_frame', '_create_collection_frame', '_create_operation_frame',
                    '_create_log_frame', '_create_captcha_frame', '_create_page_turn_frame',
                    '_log_info', '_update_always_on_top', '_show_about', '_show_help'
                ]
            },
            'browser_controller.py': {
                'description': '浏览器控制相关',
                'methods': [
                    'connect_browser', 'disconnect_browser', '_get_chrome_debugger_url',
                    '_test_browser_connection', '_switch_focus_to_browser',
                    '_switch_focus_between_browser_and_tool', '_manage_focus',
                    '_ensure_focus_for_clipboard'
                ]
            },
            'element_collector.py': {
                'description': '元素采集相关',
                'methods': [
                    'start_collection', 'stop_collection', '_collection_loop',
                    '_get_hovered_xpath_recursive', '_get_hovered_element_info',
                    '_capture_element_info', '_save_elements_to_json', '_load_elements_from_json',
                    '_configure_operations', '_preview_element', '_collect_ref1_xpath',
                    '_collect_ref2_xpath', '_collect_scroll_container_xpath',
                    '_learn_xpath_pattern', '_generate_xpath_for_item'
                ]
            },
            'data_processor.py': {
                'description': '数据处理和导出相关',
                'methods': [
                    'run_actions_loop', '_execute_operation', '_find_element_smart',
                    '_get_order_count', '_extract_current_order_id', '_scroll_to_next_order',
                    '_scroll_with_javascript', '_swipe_with_pyautogui', '_scroll_with_keys',
                    '_click_next_page', 'export_to_excel', '_check_shipping_info_before_export',
                    '_save_screenshot'
                ]
            },
            'clipboard_manager.py': {
                'description': '剪贴板管理相关',
                'methods': [
                    '_get_clipboard_content', '_wait_for_clipboard_content',
                    '_store_clipboard_content', '_save_clipboard_mappings',
                    '_load_clipboard_mappings', '_clean_existing_clipboard_mappings',
                    '_start_clipboard_monitor', '_stop_clipboard_monitor',
                    '_manually_associate_clipboard_with_order_id', '_is_valid_shipping_info',
                    '_batch_review_shipping_info', '_confirm_valid_content', '_delete_mapping'
                ]
            },
            'captcha_detector.py': {
                'description': '验证码检测相关',
                'methods': [
                    '_show_captcha_manager', '_update_template_list',
                    '_select_captcha_target_window', '_add_captcha_template',
                    '_clear_captcha_templates', '_auto_load_captcha_templates',
                    '_capture_window_as_template', '_start_captcha_detection',
                    '_stop_captcha_detection', '_captcha_detection_loop',
                    '_capture_target_window', '_template_match_detection',
                    '_mask_layer_detection', '_update_captcha_status_display',
                    '_on_captcha_detected', '_on_captcha_disappeared'
                ]
            },
            'page_turner.py': {
                'description': '翻页功能相关',
                'methods': [
                    'collect_page_turn_element', '_handle_page_turn_collection',
                    '_highlight_element_with_script', '_get_hovered_xpath',
                    '_screen_to_browser_coords', '_highlight_element',
                    'check_page_turn_needed', 'execute_page_turn',
                    '_save_page_screenshot', '_compare_screenshots',
                    '_clear_screenshots', 'on_page_count_changed'
                ]
            },
            'config_manager.py': {
                'description': '配置管理相关',
                'methods': [
                    '_load_offset_config', '_save_offset_config', '_reset_offsets',
                    '_show_offset_manager', '_export_offset_config', '_import_offset_config'
                ]
            }
        }
        
        # 生成各个模块文件
        for module_name, module_info in modules.items():
            if module_name == 'utils.py':
                content = module_info['create_func']()
            else:
                content = self.create_module_with_methods(
                    module_name, 
                    module_info['methods'], 
                    module_info['description']
                )
            
            # 写入文件
            output_file = os.path.join(self.output_dir, module_name)
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(content)
            
            print(f"✓ 已生成: {module_name} ({module_info['description']})")
        
        # 生成主模块
        main_content = self.create_main_module()
        main_file = os.path.join(self.output_dir, "main.py")
        with open(main_file, 'w', encoding='utf-8') as f:
            f.write(main_content)
        
        print(f"✓ 已生成: main.py (主应用程序入口)")
        
        print()
        print("✓ 重新拆分完成！")
        print()
        print("生成的模块文件:")
        for module_name, module_info in modules.items():
            print(f"  - {module_name} ({module_info['description']})")
        print(f"  - main.py (主应用程序入口)")

def main():
    """主函数"""
    print("=" * 60)
    print("main.py 模块重新拆分工具")
    print("=" * 60)
    print()
    
    # 检查源文件是否存在
    source_file = "f:\\shoujifahuoxinxi\\main_original.py"
    if not os.path.exists(source_file):
        print(f"错误: 源文件不存在 - {source_file}")
        input("按回车键退出...")
        return
    
    # 确认操作
    print(f"即将重新拆分文件: {source_file}")
    print("这将重新生成所有模块文件。")
    print()
    confirm = input("确认继续？(y/N): ")
    if confirm.lower() != 'y':
        print("操作已取消。")
        return
    
    try:
        # 执行拆分
        splitter = MainSplitter(source_file)
        splitter.split_modules()
        
        print()
        print("重新拆分成功完成！")
        print("现在可以运行新的main.py文件来测试应用程序。")
        
    except Exception as e:
        print(f"拆分过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
    
    input("\n按回车键退出...")

if __name__ == "__main__":
    main()