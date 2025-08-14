#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
工具函数和常量

从原始main.py文件中拆分出来的模块
"""

# 标准库导入
import os
import sys
import time
import json
import logging
import threading
import subprocess
import tempfile
import traceback
import ctypes
import re
from datetime import datetime

# GUI相关导入
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

# 第三方库导入（这些会在运行时检查和安装）
try:
    import websocket
    import requests
    import cv2
    import numpy as np
    import pyautogui
    from PIL import Image, ImageTk
    import win32gui
    import win32ui
    import win32con
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    import pandas as pd
    from docx import Document
    from docx.shared import Inches
    import pyperclip
except ImportError:
    # 这些导入失败是正常的，会在依赖检查时处理
    pass

# 配置日志记录
log_filename = "收货信息自动采集工具.log"
# 如果日志文件已存在，则删除它（每次启动都使用新的日志文件）
if os.path.exists(log_filename):
    try:
        os.remove(log_filename)
    except:
        pass  # 如果删除失败，继续执行

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_filename, encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)


def init_tcl_compatibility():
    """初始化TCL兼容性设置，解决版本冲突问题"""
    try:
        # 获取当前Python环境的TCL路径
        import sys
        python_dir = os.path.dirname(sys.executable)
        
        # 尝试多个可能的TCL路径
        possible_tcl_paths = [
            os.path.join(python_dir, 'tcl', 'tcl8.6'),
            os.path.join(python_dir, 'lib', 'tcl8.6'),
            r"C:\Users\Administrator\AppData\Local\Programs\Python\Python311\tcl\tcl8.6",
            # Conda环境路径
            os.path.join(python_dir, '..', 'lib', 'tcl8.6'),
            os.path.join(python_dir, '..', 'Library', 'lib', 'tcl8.6')
        ]
        
        possible_tk_paths = [
            os.path.join(python_dir, 'tcl', 'tk8.6'),
            os.path.join(python_dir, 'lib', 'tk8.6'),
            r"C:\Users\Administrator\AppData\Local\Programs\Python\Python311\tcl\tk8.6",
            # Conda环境路径
            os.path.join(python_dir, '..', 'lib', 'tk8.6'),
            os.path.join(python_dir, '..', 'Library', 'lib', 'tk8.6')
        ]
        
        # 查找存在的TCL路径
        tcl_path = None
        for path in possible_tcl_paths:
            if os.path.exists(path):
                tcl_path = path
                break
        
        tk_path = None
        for path in possible_tk_paths:
            if os.path.exists(path):
                tk_path = path
                break
        
        if tcl_path and tk_path:
            # 设置环境变量
            os.environ['TCL_LIBRARY'] = tcl_path
            os.environ['TK_LIBRARY'] = tk_path
            
            # 禁用精确版本检查，允许版本不匹配
            os.environ['TCL_EXACT_VERSION'] = '0'
            os.environ['TK_EXACT_VERSION'] = '0'
            
            print(f"TCL兼容性设置成功: TCL={tcl_path}, TK={tk_path}")
            return True
        else:
            print(f"未找到有效的TCL/TK路径")
            return False
            
    except Exception as e:
        print(f"TCL兼容性设置失败: {e}")
        return False


def install_package(package):
    """安装单个包"""
    try:
        print(f"正在安装 {package}...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        print(f"✓ {package} 安装成功")
        return True
    except subprocess.CalledProcessError as e:
        print(f"✗ {package} 安装失败: {e}")
        return False


def auto_install_dependencies(missing_packages):
    """自动安装缺失的依赖包"""
    print("收货信息自动采集工具 - 依赖安装")
    print("=" * 50)
    print(f"检测到缺失的依赖包: {', '.join(missing_packages)}")
    print("正在自动安装...")
    print("-" * 30)
    
    # 安装每个依赖包
    success_count = 0
    failed_packages = []
    
    for package in missing_packages:
        if install_package(package):
            success_count += 1
        else:
            failed_packages.append(package)
        print()
    
    # 显示安装结果
    print("=" * 50)
    print(f"安装完成！成功: {success_count}/{len(missing_packages)}")
    
    if failed_packages:
        print(f"\n安装失败的包: {', '.join(failed_packages)}")
        print("请手动安装这些包或检查网络连接。")
        return False
    else:
        print("\n所有依赖包安装成功！程序将重新启动...")
        return True


def check_and_install_dependencies():
    """检查并安装缺失的依赖包，避免重复安装"""
    # 检测是否是打包环境
    if getattr(sys, 'frozen', False):
        print("打包环境运行，跳过依赖自动安装")
        return True
        
    import importlib
    try:
        from importlib.metadata import distribution
    except ImportError:
        # Python < 3.8 fallback
        from importlib_metadata import distribution
    
    # 定义依赖包映射（模块名 -> 安装名）
    dependencies = {
        'websocket': 'websocket-client',
        'requests': 'requests',
        'cv2': 'opencv-python',
        'numpy': 'numpy',
        'pyautogui': 'pyautogui',
        'PIL': 'pillow',
        'win32gui': 'pywin32',
        'selenium': 'selenium',
        'pandas': 'pandas',
        'docx': 'python-docx',
        'pyperclip': 'pyperclip'
    }
    
    missing_packages = []
    
    # 检查每个依赖
    for module_name, install_name in dependencies.items():
        try:
            # 使用importlib.metadata检查包是否已安装
            distribution(install_name)
        except Exception:
            missing_packages.append(install_name)
    
    # 如果有缺失的包，提供自动安装选项
    if missing_packages:
        print(f"检测到缺失的依赖包: {', '.join(missing_packages)}")
        
        # 尝试自动安装
        if auto_install_dependencies(missing_packages):
            # 安装成功，重新启动程序
            print("\n正在重新启动程序...")
            try:
                import time
                time.sleep(2)
                os.execv(sys.executable, [sys.executable] + sys.argv)
            except Exception as e:
                print(f"重新启动失败: {e}")
                print("请手动重新运行程序。")
        
        return False
    
    return True


dependencies_ok = check_and_install_dependencies()


def log_retry_event(self, event_type: str, element_name: str, details: dict = None):
    """记录重试事件日志 - 新增方法"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_entry = {
        "timestamp": timestamp,
        "event_type": event_type,  # 'retry_start', 'retry_success', 'retry_failed', 'retry_abandoned'
        "element_name": element_name,
        "details": details or {}
    }
    
    # 写入重试日志文件
    retry_log_file = "retry_events.log"
    try:
        with open(retry_log_file, 'a', encoding='utf-8') as f:
            f.write(f"{json.dumps(log_entry, ensure_ascii=False)}\n")
    except Exception as e:
        print(f"写入重试日志失败: {e}")
    
    # 同时输出到控制台
    color_map = {
        'retry_start': 'orange',
        'retry_success': 'green',
        'retry_failed': 'red',
        'retry_abandoned': 'purple'
    }
    
    if hasattr(self, '_log_info'):
        self._log_info(f"[重试] {event_type}: {element_name}", color_map.get(event_type, 'blue'))
    else:
        print(f"[重试] {event_type}: {element_name}")


def create_retry_log_entry(event_type: str, element_name: str, details: dict = None) -> dict:
    """创建重试日志条目 - 独立函数版本"""
    return {
        "timestamp": datetime.now().isoformat(),
        "event_type": event_type,
        "element_name": element_name,
        "details": details or {}
    }


def write_retry_log(log_entry: dict, log_file: str = "retry_events.log") -> bool:
    """写入重试日志到文件"""
    try:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{json.dumps(log_entry, ensure_ascii=False)}\n")
        return True
    except Exception as e:
        print(f"写入重试日志失败: {e}")
        return False
