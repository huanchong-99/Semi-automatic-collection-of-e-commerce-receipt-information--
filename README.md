# 收货信息自动采集工具

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://python.org)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)]()
[![Edge Browser](https://img.shields.io/badge/browser-Microsoft%20Edge-blue.svg)]()

一个专为电商平台订单收货信息自动采集而设计的Python桌面应用程序。采用模块化架构设计，通过浏览器调试模式和智能元素定位技术，实现批量订单信息的自动化采集、处理和导出。支持验证码智能检测、坐标缓存重试、多格式数据导出等高级功能。

## ✨ 主要特性

### 🏗️ 模块化架构设计
- **分层架构**：UI层、业务逻辑层、服务层、基础设施层清晰分离
- **模块解耦**：15个独立模块，职责明确，便于维护和扩展
- **配置管理**：统一的配置管理系统，支持配置持久化和导入导出
- **错误恢复**：完善的异常处理和自动恢复机制

### 🌐 智能浏览器控制
- **调试模式连接**：通过Edge浏览器远程调试端口连接，保留用户登录状态和个人设置
- **自动启动管理**：一键启动浏览器并建立调试连接，支持本地和系统驱动
- **断线重连机制**：支持连接中断后的自动重连和状态恢复
- **用户数据持久化**：独立的浏览器配置文件目录，保持登录状态和个人设置
- **焦点管理**：智能焦点切换，确保操作的准确性

### 🎯 精确元素采集
- **悬停采集技术**：鼠标悬停目标元素，按`.`键快速采集XPath
- **跨框架支持**：递归注入JavaScript监听器，自动处理嵌套iframe中的元素定位
- **多种采集模式**：
  - 📄 获取文本内容 (getText)
  - 🖱️ 点击元素 (click)
  - 📋 点击并获取剪贴板内容 (clickAndGetClipboard)
- **智能XPath生成**：生成稳定、唯一的元素定位符，适应页面结构变化
- **自定义命名**：为每个采集元素设置有意义的名称和操作类型
- **实时预览验证**：支持元素高亮预览、定位验证和可见性检查

### 🔄 智能批量处理
- **模块化翻页**：双层循环架构，页面级别和页内订单独立处理
- **自动数量检测**：根据"待发货"等关键元素自动确定处理次数
- **可视化序列配置**：直观配置操作顺序和执行模式
- **灵活循环控制**：
  - 🔁 始终循环：每个订单都执行该操作
  - 1️⃣ 单次循环：仅在第一个订单执行
- **智能自动滚动**：基于元素位置智能计算滚动距离，支持距离学习
- **状态重置机制**：每页处理完成后自动重置缓存和计数器
- **实时进度反馈**：详细的执行进度和状态信息

### 🎮 精确控制系统
- **WASD微调**：点击前支持WASD键精确微调鼠标位置
- **偏移量记忆**：自动保存和应用每个元素的位置偏移，支持元素级别配置
- **坐标缓存系统**：智能缓存元素坐标，支持重试模式和缓存验证
- **人工确认模式**：可选择每次操作前弹窗确认
- **快捷键控制**：
  - `.`键：暂停当前操作
  - `-`键：继续执行
  - `*`键：终止并导出数据

### 🔄 智能重试机制
- **多策略重试**：缓存坐标、智能搜索、备用XPath等多种重试策略
- **配置化管理**：可配置重试次数、延迟时间、超时设置等参数
- **坐标验证**：智能验证缓存坐标的有效性和时效性
- **统计分析**：详细的重试统计和成功率分析
- **手动干预**：支持重试失败时的手动干预选项

### 📊 多格式数据导出
- **Excel表格** (.xlsx)：结构化数据表格，便于数据分析
- **Word文档** (.docx)：格式化文档，便于打印和分享
- **JSON数据** (.json)：原始数据结构，便于程序处理
- **智能去重**：自动处理重复字段（如"复制完整收货信息"）
- **自定义保存**：用户自定义文件名和保存位置
- **数据清洗**：自动验证和清洗采集的数据

### 🛡️ 验证码智能检测
- **实时监控**：后台异步检测页面验证码出现，不影响主要功能
- **模板匹配**：支持自定义验证码模板图片，自动加载和管理
- **遮罩层检测**：检测页面遮罩层变化和窗口状态
- **自动暂停**：检测到验证码时自动暂停操作并提醒用户
- **声音提醒**：验证码出现时播放系统提示音
- **窗口管理**：智能选择目标窗口，支持多窗口环境

### ⚙️ 丰富辅助功能
- **窗口置顶**：工具窗口可选择始终置顶显示
- **距离学习**：智能学习两个元素间的滚动距离
- **配置持久化**：自动保存偏移量、操作序列等所有配置
- **详细日志记录**：完整的操作日志和错误记录，支持日志级别控制
- **依赖自动管理**：首次运行自动检测并安装所需依赖包
- **配置导入导出**：支持配置文件的备份、导入和导出
- **统计分析**：提供详细的操作统计和性能分析

## 🚀 快速开始

### 系统要求

- **操作系统**：Windows 10/11 (64位)
- **Python版本**：3.8 或更高版本
- **浏览器**：Microsoft Edge（最新版本）
- **内存**：建议 4GB 以上
- **网络**：需要联网（用于依赖包安装）
- **权限**：建议以管理员身份运行（用于系统API调用）

### ⚠️ 重要：浏览器安装要求

**Microsoft Edge 必须安装在以下路径之一：**

```
C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
```

**如果您的 Edge 浏览器安装在其他位置，程序将无法正常启动浏览器。**

**检查方法：**
1. 打开文件资源管理器
2. 导航到 `C:\Program Files (x86)\Microsoft\Edge\Application\`
3. 确认存在 `msedge.exe` 文件

**如果路径不正确：**
- 卸载当前的 Edge 浏览器
- 从 [Microsoft 官网](https://www.microsoft.com/edge) 重新下载并安装
- 或者修改 `browser_controller.py` 文件中的 `BROWSER_PATH` 变量为您的实际安装路径

### 安装步骤

#### 方法一：克隆仓库（推荐）

```bash
# 1. 克隆项目
git clone https://github.com/huanchong-99/Semi-automatic-collection-of-e-commerce-receipt-information-.git
cd Semi-automatic-collection-of-e-commerce-receipt-information-

# 2. 创建虚拟环境（推荐）
python -m venv .venv
.venv\Scripts\activate  # Windows

# 3. 安装依赖
pip install -r requirements.txt

# 4. 运行程序
python main.py
```

#### 方法二：直接下载

1. 下载项目压缩包并解压到本地目录
2. 确认项目目录包含以下关键文件：
   - `main.py`（主程序入口）
   - `requirements.txt`（依赖列表）
   - `msedgedriver.exe`（Edge驱动程序）
3. 打开命令提示符或PowerShell，进入项目目录
4. 执行以下命令：

```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

### 首次运行

程序首次运行时会自动检测并安装缺失的依赖包。如果自动安装失败，请手动安装：

```bash
pip install websocket-client requests opencv-python numpy pyautogui pillow pywin32 selenium pandas python-docx pyperclip
```

**首次运行检查清单：**
1. ✅ Python 3.8+ 已安装
2. ✅ Edge 浏览器安装在正确路径
3. ✅ 所有依赖包安装成功
4. ✅ 项目目录包含 `msedgedriver.exe`
5. ✅ 以管理员权限运行（推荐）

## 📖 使用指南

### 1. 环境准备

- 确保已安装 Microsoft Edge 浏览器
- 确保 Python 环境正常
- 首次运行会自动安装所需依赖包

### 2. 启动和连接

1. 运行 `python main.py` 启动程序
2. 点击"打开浏览器"按钮
3. 等待浏览器启动并建立调试连接
4. 在浏览器中登录电商平台并导航到订单页面

### 3. 元素采集配置

1. 切换到"采集模式"
2. 将鼠标悬停在需要采集的元素上
3. 按 `.` 键采集元素XPath
4. 在弹出对话框中选择操作类型并设置元素名称
5. 重复以上步骤采集所有需要的元素

### 4. 操作序列配置

1. 点击"配置操作"按钮
2. 设置每个元素的执行顺序
3. 选择循环模式（始终循环/单次循环）
4. 指定订单数量检测元素（可选）
5. 确认配置

### 5. 开始采集

1. 切换回"正常模式"
2. 点击"开始"按钮
3. 程序将自动执行配置的操作序列
4. 可随时使用暂停/继续功能

### 6. 数据导出

1. 采集完成后点击相应的导出按钮
2. 选择保存位置和文件名
3. 数据将以选定格式保存

## 🏗️ 技术架构

### 模块化架构设计

```
┌─────────────────────────────────────────────────────────────┐
│                    UI层 (Presentation Layer)                │
├─────────────────────────────────────────────────────────────┤
│  main.py  │  ui_components.py  │  operation_sequence_dialog.py │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                 业务逻辑层 (Business Logic Layer)            │
├─────────────────────────────────────────────────────────────┤
│  data_processor.py  │  element_collector.py  │  page_turner.py │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   服务层 (Service Layer)                    │
├─────────────────────────────────────────────────────────────┤
│  browser_controller.py  │  captcha_detector.py  │  retry_manager.py │
│  clipboard_manager.py   │  coordinate_cache.py   │  config_manager.py │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   基础设施层 (Infrastructure)                │
├─────────────────────────────────────────────────────────────┤
│  Selenium WebDriver  │  OpenCV  │  PyAutoGUI  │  文件系统   │
└─────────────────────────────────────────────────────────────┘
```

### 核心技术栈

| 技术组件 | 版本要求 | 用途说明 |
|----------|----------|----------|
| **Python** | 3.8+ | 主要开发语言 |
| **Tkinter** | 8.6+ | GUI界面框架 |
| **Selenium** | 4.0+ | 浏览器自动化控制 |
| **OpenCV** | 4.5+ | 图像处理和验证码检测 |
| **PyAutoGUI** | 0.9+ | 鼠标键盘自动化操作 |
| **Pandas** | 1.3+ | 数据处理和Excel导出 |
| **python-docx** | 0.8+ | Word文档生成 |
| **pywin32** | 300+ | Windows系统API调用 |
| **Pillow** | 8.0+ | 图像处理支持 |
| **pyperclip** | 1.8+ | 剪贴板操作 |

### 关键实现特性

- **模块化架构**：15个独立模块，职责明确，便于维护和扩展
- **浏览器调试连接**：通过 remote-debugging-port 方式连接 Edge 浏览器
- **递归元素定位**：注入 JavaScript 脚本实现跨 iframe 元素定位
- **智能滚动算法**：基于元素位置计算智能滚动距离
- **坐标缓存系统**：智能缓存和验证元素坐标，支持重试机制
- **数据持久化**：JSON 格式保存配置和采集数据
- **完善异常处理**：多层次错误处理和自动恢复机制
- **异步验证码检测**：后台实时监控，不影响主要功能

## 📁 项目结构

```
shoujifahuoxinxi/
├── 📁 核心模块
│   ├── main.py                      # 主程序入口
│   ├── ui_components.py             # UI界面组件
│   ├── browser_controller.py        # 浏览器控制模块
│   ├── element_collector.py         # 元素采集模块
│   ├── data_processor.py            # 数据处理模块
│   ├── captcha_detector.py          # 验证码检测模块
│   ├── page_turner.py               # 翻页功能模块
│   ├── clipboard_manager.py         # 剪贴板管理模块
│   ├── coordinate_cache.py          # 坐标缓存模块
│   ├── config_manager.py            # 配置管理模块
│   ├── retry_manager.py             # 重试管理模块
│   ├── operation_sequence_dialog.py # 操作序列配置对话框
│   └── utils.py                     # 工具函数模块
│
├── 📁 配置文件
│   ├── captcha_config.json          # 验证码检测配置
│   ├── element_config.json          # 元素配置文件
│   ├── offset_config.json           # 偏移量配置文件
│   ├── retry_config.json            # 重试功能配置文件
│   ├── 采集到的元素.json             # 采集元素数据存储
│   ├── clipboard_mappings.json      # 订单ID与收货信息映射（运行时生成）
│   └── coordinate_cache.json        # 坐标缓存数据（运行时生成）
│
├── 📁 资源文件
│   ├── captcha_templates/           # 验证码模板图片目录
│   ├── msedgedriver.exe             # Edge浏览器驱动程序
│   └── requirements.txt             # Python依赖包列表
│
├── 📁 备份目录
│   └── backup/                      # 代码备份文件
│
├── 📁 运行时目录（自动生成）
│   ├── pdd_browser_profile/         # 浏览器用户数据目录
│   ├── captcha_debug/               # 验证码调试信息目录
│   ├── data_backups/                # 数据备份存储目录
│   ├── page_screenshots/            # 页面截图临时存储目录
│   └── *.log                        # 程序运行日志文件
│
└── 📁 文档
    ├── README.md                    # 项目说明文档
    ├── 文件作用.md                   # 文件说明文档
    └── 项目开发与问题解决综合文档.md   # 开发文档
```

## ⚙️ 配置文件说明

### 静态配置文件

| 文件名 | 用途 | 格式 | 是否必需 |
|--------|------|------|----------|
| `采集到的元素.json` | 保存采集的元素XPath和操作配置 | JSON | ✅ 必需 |
| `element_config.json` | 元素引用配置（ref1_xpath, ref2_xpath等） | JSON | ✅ 必需 |
| `offset_config.json` | 保存WASD微调的偏移量配置 | JSON | ✅ 必需 |
| `captcha_config.json` | 验证码检测相关配置 | JSON | ✅ 必需 |
| `retry_config.json` | 重试机制配置参数 | JSON | ✅ 必需 |

### 运行时生成文件

| 文件名 | 用途 | 格式 | 自动生成 |
|--------|------|------|----------|
| `clipboard_mappings.json` | 订单ID与收货信息的映射关系 | JSON | 🔄 运行时 |
| `coordinate_cache.json` | 元素坐标缓存数据 | JSON | 🔄 运行时 |
| `*.log` | 程序运行的详细日志记录 | 文本 | 🔄 运行时 |

### 配置文件详细说明

**采集到的元素.json**
- 存储通过悬停采集的元素信息
- 包含XPath、操作类型、自定义名称等
- 支持getText、click、clickAndGetClipboard操作

**element_config.json**
- 存储参考元素的XPath配置
- 用于智能循环和翻页功能
- 包含ref1_xpath、ref2_xpath、next_page_xpath等

**offset_config.json**
- 存储每个元素的WASD微调偏移量
- 支持元素级别的精确位置调整
- 自动保存用户的微调操作

**retry_config.json**
- 重试机制的详细配置
- 包含重试次数、延迟时间、策略选择等
- 支持坐标缓存和验证设置

## 🔧 故障排除

### 常见问题

#### 1. 浏览器连接失败
- **症状**：提示"无法连接到浏览器"或"连接失败"
- **检查项**：
  - Edge浏览器是否安装在 `C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe`
  - 调试端口是否被占用
  - 防火墙是否阻止连接
- **解决方案**：
  - 确认浏览器安装路径正确
  - 重启程序或更换调试端口
  - 临时关闭防火墙测试
  - 更新到最新版本的Edge浏览器

#### 2. 元素定位失败
- **症状**：提示"未找到元素"或"元素不可见"
- **检查项**：
  - 页面结构是否发生变化
  - 元素是否在iframe中
  - 页面是否完全加载
- **解决方案**：
  - 重新采集元素XPath
  - 检查页面是否有动态加载内容
  - 增加等待时间
  - 使用重试机制和坐标缓存

#### 3. 依赖包安装失败
- **症状**：ModuleNotFoundError或导入错误
- **检查项**：
  - 网络连接是否正常
  - Python版本是否符合要求
  - pip版本是否最新
- **解决方案**：
  - 使用国内镜像源安装
  - 升级pip到最新版本
  - 手动安装失败的包

```bash
# 使用清华镜像源
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple/

# 升级pip
python -m pip install --upgrade pip

# 手动安装关键依赖
pip install selenium==4.15.0
pip install opencv-python==4.8.1.78
```

#### 4. 验证码检测异常
- **症状**：验证码检测不准确或频繁误报
- **检查项**：
  - 目标窗口选择是否正确
  - 验证码模板是否匹配
  - 检测区域设置是否合适
- **解决方案**：
  - 重新选择目标窗口（选择包含"Microsoft Edge"的窗口）
  - 更新验证码模板图片
  - 调整检测灵敏度
  - 清除旧的模板文件

#### 5. 权限相关问题
- **症状**：无法操作鼠标键盘或访问系统API
- **检查项**：
  - 程序是否以管理员身份运行
  - 杀毒软件是否阻止操作
  - UAC设置是否过于严格
- **解决方案**：
  - 右键选择"以管理员身份运行"
  - 将程序添加到杀毒软件白名单
  - 临时降低UAC级别

#### 6. 坐标偏移问题
- **症状**：点击位置不准确
- **检查项**：
  - 显示器缩放设置
  - 浏览器缩放级别
  - 多显示器配置
- **解决方案**：
  - 使用WASD微调功能精确调整
  - 检查显示器缩放设置（建议100%）
  - 重置浏览器缩放到100%
  - 重新采集元素坐标

#### 7. 内存占用过高
- **症状**：程序运行缓慢或系统卡顿
- **检查项**：
  - 处理的数据量是否过大
  - 是否有内存泄漏
  - 验证码检测频率是否过高
- **解决方案**：
  - 分批处理大量数据
  - 定期重启程序
  - 降低验证码检测频率
  - 清理临时文件和缓存

### 日志查看

程序运行过程中的详细信息记录在多个日志文件中，遇到问题时可查看相应日志获取详细错误信息：

**主要日志文件：**
- `收货信息自动采集工具.log` - 主程序运行日志
- `captcha_detector.log` - 验证码检测日志
- `retry_events.log` - 重试机制日志

**Windows查看日志命令：**
```powershell
# 查看最新日志（PowerShell）
Get-Content "收货信息自动采集工具.log" -Tail 50

# 搜索错误信息
Select-String "ERROR" "收货信息自动采集工具.log"

# 实时监控日志
Get-Content "收货信息自动采集工具.log" -Wait -Tail 10
```

**日志级别说明：**
- `INFO` - 一般信息
- `WARNING` - 警告信息
- `ERROR` - 错误信息
- `DEBUG` - 调试信息

## 🤝 贡献指南

我们欢迎所有形式的贡献！请查看 [贡献指南](CONTRIBUTING.md) 了解详细信息。

### 如何贡献

1. Fork 本仓库
2. 创建您的特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交您的更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 打开一个 Pull Request

### 报告问题

如果您发现了bug或有功能建议，请在 [Issues](https://github.com/your-username/shipping-info-collector/issues) 页面创建新的issue。

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详细信息。

## ⚠️ 免责声明

- 本工具仅供学习和合法用途使用
- 请遵守相关网站的使用条款和法律法规
- 使用本工具产生的任何后果由用户自行承担
- 请确保您有权限访问和采集相关数据

## 🙏 致谢

感谢以下开源项目的支持：

- [Selenium](https://selenium.dev/) - 浏览器自动化框架
- [OpenCV](https://opencv.org/) - 计算机视觉库
- [PyAutoGUI](https://pyautogui.readthedocs.io/) - 自动化操作库
- [Pandas](https://pandas.pydata.org/) - 数据处理库
- [Tkinter](https://docs.python.org/3/library/tkinter.html) - Python GUI框架

## 📞 联系方式

- 项目主页：[GitHub Repository](https://github.com/huanchong-99/Semi-automatic-collection-of-e-commerce-receipt-information-)
- 问题反馈：[Issues](https://github.com/huanchong-99/Semi-automatic-collection-of-e-commerce-receipt-information-/issues)
- 技术文档：查看项目中的 `项目开发与问题解决综合文档.md`
- 文件说明：查看项目中的 `文件作用.md`

## 🔄 版本更新

**当前版本特性：**
- ✅ 模块化架构重构完成
- ✅ 智能重试机制
- ✅ 坐标缓存系统
- ✅ 异步验证码检测
- ✅ 双层循环翻页
- ✅ 配置管理系统

**后续计划：**
- 🔄 支持更多浏览器
- 🔄 增加更多数据导出格式
- 🔄 优化性能和稳定性
- 🔄 增加更多自动化功能

---

⭐ 如果这个项目对您有帮助，请给我们一个星标！

💡 **提示**：首次使用建议仔细阅读本文档，特别是浏览器安装要求和故障排除部分。