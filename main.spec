# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    # 主程序入口
    ['main.py'],
    
    # 项目路径
    pathex=[
        'C:\\Users\\Administrator\\AppData\\Local\\Programs\\Python\\Python313\\Lib\\site-packages'
    ],
    
    # 二进制文件（包含Edge驱动）
    binaries=[
        ('msedgedriver.exe', '.'),
    ],
    
    # 数据文件（配置文件、模板图片等）
    datas=[
        # 配置文件
        ('captcha_config.json', '.'),
        ('element_config.json', '.'),
        ('offset_config.json', '.'),
        ('retry_config.json', '.'),
        ('采集到的元素.json', '.'),
        
        # 验证码模板图片目录
        ('captcha_templates', 'captcha_templates'),
        
        # 页面截图目录（如果存在）
        ('page_screenshots', 'page_screenshots'),
    ],
    
    # 隐式导入的模块
    hiddenimports=[
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.edge',
        'selenium.webdriver.edge.options',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.common.keys',
        'selenium.common.exceptions',
        'cv2',
        'numpy',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'pyautogui',
        'pyperclip',
        'pandas',
        'docx',
        'websocket',
        'requests',
        'pywin32.win32gui',
        'win32ui',
        'win32con',
        'win32api',
        'winsound',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.simpledialog',
        'tkinter.scrolledtext',
        'pywintypes',
        'pywin32',
    ],
    
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    
    # 排除备份文件和不需要的文件
    excludes=[
        'backup',
        'main_original',
        'main_backup',
        'data_processor_backup_20250729_002158',
        'page_turner_backup',
        '__pycache__',
        '.vercel',
        '.git',
        '.gitignore',
        'README.md',
        '文件作用.md',
        '问题修复方案.md',
        '项目开发与问题解决综合文档.md',
    ],
    
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
