
# 工程量计算书

## 已知问题

-  当打开多个AUTOCAD实例时,反查功能默认与最开始连接的进程通信。

## 项目概述

工程量计算书是一款专门为建筑、安装等工程领域设计的专业软件，集成了AutoCAD和Microsoft Excel的强大功能，实现了从CAD图纸中快速提取工程量数据、自动计算和智能汇总的一体化解决方案。

**核心价值：**
- 大幅提升工程量计算效率
- 减少人工计算错误
- 实现数据的可视化管理
- 提供便捷的追溯和反查功能

## 技术栈选型与架构设计

### 开发技术
- **开发语言**: Visual Basic .NET (VB.NET)
- **目标框架**: .NET Framework 4.8（ClassLibraryAutocad2025 为 .NET 8 / net8.0-windows）
- **开发环境**: Visual Studio 2022 / 2026
- **版本控制**: Git

### 第三方依赖
- **AutoCAD API**: 支持AutoCAD 2012~2025各版本
- **Excel VSTO**: Excel COM自动化接口
- **Windows API**: 系统级功能调用

### 系统架构

项目采用模块化架构设计，主要包括以下组件：

```
工程量计算书/
├── HomeStart/                # 启动器模块（单实例 + 插件自注册）
├── 工程量计算书(EXCEL)/     # Excel VSTO 插件核心
├── ClassLibraryAutocad2012~2025/  # AutoCAD 各版本插件（代码集中于 SharedAuxiliary + 2016 窗体源）
├── SharedAuxiliary/         # AutoCAD 共享代码库（共享项目，编译进每个版本插件）
├── AuxiliaryClassLibrary/   # 公共辅助类库（注册表/设置/共享内存/通知/COM 辅助）
├── Lisp/                    # AutoLISP 脚本（插件启动时加载）
├── SettingsManagerTests/    # 单元测试项目（MSTest，23 个用例）
├── scripts/                 # 工程脚本（编码检查/批量构建/版本管理）
├── Setup/                   # Inno Setup 安装脚本
└── dist/                    # 发布产物（安装包）
```

## 核心功能模块说明

### 1. 启动器模块 (HomeStart)
- 负责程序的启动与初始化
- 自动检测并注册Excel和AutoCAD插件
- 管理备份文件目录
- 提供启动动画界面
- 支持通过命令行参数直接打开计算书文件

### 2. Excel插件模块 (工程量计算书)
**主要功能：**
- **工程量计算表**: 提供标准化的计算表模板
- **表达式计算**: 支持复杂的数学表达式计算
- **数据汇总**: 自动汇总计算结果生成汇总表
- **数据查询与筛选**: 快速查找和筛选工程量数据
- **批注管理**: 支持图片批注和文字批注
- **变量表**: 支持自定义变量，便于统一调整参数
- **辅助工具**:
  - 型钢计算器
  - 钢管计算器
  - 防腐绝热计算
- **备份管理**: 自动备份文件，防止数据丢失
- **截屏功能**: 直接截取屏幕并插入到单元格批注中

### 3. AutoCAD插件模块 (ClassLibraryAutocad*)
支持AutoCAD 2012至2025各版本，提供以下功能：

**核心命令：**
- **CC (测量长度)**: 在CAD中绘制多段线测量长度，自动将结果传输到Excel
- **GG (过滤提取)**: 根据条件筛选图元并提取工程量
- **TQ (提取统计)**: 统计同类图元数量（点、线、圆、块、文本等）
- **BZLG (布置立管)**: 快速布置和修改立管
- **CX (查询对象)**: 查询选中图元信息
- **QC (清除标记)**: 清除图元的提取标记（原 QCBZ 命令已合并至此）
- **BX (标注长度)**: 在图纸上标注线段长度
- **RX (刷新属性)**: 刷新图元扩展属性
- **FX (反查Excel)**: 选中CAD图元后反查引用该图元的Excel单元格
- **GETTEXT (提取文本)**: 提取文本内容到Excel

**其他功能：**
- 图层控制：支持对选中图元所在图层进行冻结/隐藏等操作
- 自定义工具栏：提供直观的工具按钮
- Ribbon界面：现代化的功能区界面
- 实体过滤：根据属性条件筛选图元
- 图元扩展数据：在CAD图元上保存扩展信息，支持追溯

### 4. 共享辅助模块 (SharedAuxiliary)
提供AutoCAD各版本插件共享的功能代码：
- 命令系统 (Command)
- 实体信息提取 (EntInfo)
- 选择集增强 (SelectionEx)
- 自定义面板 (CustomPalette)
- 绘制覆盖 (DrawOverrule)
- 临时图形 (Transient)
- 多重匹配 (MultipleMatching)
- Ribbon界面管理 (Ribbon)
- 静态工具模块 (StaticModule)

### 5. 公共辅助类库 (AuxiliaryClassLibrary)
提供通用工具类：
- 注册表操作 (RegditEx, RegCurrentUser)
- 设置管理 (SettingsManager)
- 共享内存管理 (SharedMemoryManager)
- Excel互操作辅助 (ExcelInteropHelper)
- Windows API包装
- 文件自动备份 (FileAutoBaker)
- 通知助手 (NotificationHelper)

## 环境配置与安装步骤

### 系统要求
- **操作系统**: Windows 10 及以上版本（安装包强制 64 位）
- **Excel**: Microsoft Excel 2010 及以上版本 (32位或64位)
- **AutoCAD**: AutoCAD 2012 - 2025（64 位）
- **.NET Framework**: .NET Framework 4.8 或更高版本（AutoCAD 2025 插件依赖 .NET 8，由 CAD 自带）

### 安装步骤
1. 下载安装包
2. 运行安装程序
3. 安装完成后，双击桌面快捷方式"工程量计算书"启动程序
4. 程序会自动注册Excel和AutoCAD插件

### 首次使用配置
1. 首次启动会自动创建"工程备份"文件夹
2. 在AutoCAD中输入命令加载插件
3. 首次使用Excel插件时，允许加载VSTO插件

## 使用指南与操作示例

### 基本工作流程

```
1. 启动工程量计算书
   ↓
2. 在Excel中创建新的计算书或打开已有文件
   ↓
3. 打开AutoCAD图纸
   ↓
4. 使用CC、GG等命令提取工程量
   ↓
5. 在Excel中进行计算和调整
   ↓
6. 生成汇总表
   ↓
7. 保存文件
```

### 常用操作示例

#### 示例1: 使用CC命令测量长度
1. 在AutoCAD中输入 `CC` 命令
2. 依次点击要测量的线段端点
3. 完成后，长度数据会自动传输到当前Excel单元格
4. 支持连续测量，结果会自动累加

#### 示例2: 使用GG命令过滤提取
1. 在Excel中选中目标单元格
2. 在AutoCAD中输入 `GG` 命令
3. 选择要提取的图元对象
4. 系统会自动统计同类图元并发送到Excel

#### 示例3: 反查功能
1. 在Excel中选中包含工程量数据的单元格
2. 使用"反查CAD"功能
3. AutoCAD中会自动选中并高亮对应的图元对象
4. 支持多选单元格批量反查

#### 示例4: 使用变量表
1. 打开变量表窗口
2. 添加变量名和对应的值
3. 在计算式中引用变量，如 `L1*数量`
4. 修改变量值时，相关计算会自动更新

### 快捷键说明
- `Ctrl+E`: 清除批注
- `Ctrl+方向键`: 控制分级显示

## 项目目录结构说明

```
工程量计算书/
├── HomeStart/                        # 启动程序（单实例 + WebView2 启动动画 + 插件自注册）
│   ├── SplashScreenStart.vb          # 主启动窗体（DevExpress Splash + WebView2 HTML 动画）
│   ├── ApplicationEvents.vb          # 单实例/未处理异常处理
│   ├── AppLogger.vb                  # 日志模块（%APPDATA%\工程量计算书\HomeStart.log）
│   ├── SingleInstancePipe.vb         # 单实例管道（命令行文件转发）
│   └── HomeStart.vbproj
│
├── 工程量计算书(EXCEL)/             # Excel VSTO插件
│   ├── ThisAddIn.vb                 # 插件入口（启动/热键/定时保存/命令分发）
│   ├── Ribbon/xlRibbonNew.vb        # Ribbon界面（XML 嵌入资源）
│   ├── CustomInfoHelper/            # 自定义信息/批注持久化（CustomXmlPart）
│   ├── HotKey/                      # 全局热键处理
│   ├── HotTracking/                 # 阅读模式高亮跟踪
│   ├── Screenshot/                  # 屏幕捕获
│   ├── ToolsForm/                   # 工具窗体（汇总/变量表/计算器/查找/自检等）
│   │   ├── FormFind/                # 查找窗体（数据筛选/反查）
│   │   ├── FormDataChecker/         # 工程数据自检工具（规则引擎 R-001~R-009）
│   │   └── FastSearch/              # 快速检索面板
│   ├── Variable/                    # 变量表
│   └── 工程量计算书.vbproj
│
├── ClassLibraryAutocad2012/ ~ /ClassLibraryAutocad2025/  # AutoCAD各版本插件
│   ├── Form/                        # 窗体源码集中存放于 ClassLibraryAutocad2016，
│   │                                #   其余版本通过 MSBuild Link 引用同一批文件
│   └── ClassLibraryAutocad*.vbproj  # 仅工程配置：版本化 ObjectARX 引用 + VER 条件编译常量
│
├── SharedAuxiliary/                 # AutoCAD共享代码库（共享项目，编译进每个版本插件）
│   ├── Command/                     # 命令注册（MCommand）与提取/过滤实现
│   ├── Main/                        # 插件生命周期（TlsApplication）与主类（ClassAutoCad）
│   ├── StaticModule/                # 静态工具（ProcomToExcel 通信/Xls/SelEntity/Xdata 等）
│   ├── MultipleMatching/            # 多对象组合匹配（空间匹配 + 块转换）
│   ├── Ribbon/                      # Ribbon 界面（GCLRibbon）
│   ├── CustomPalette/               # 自定义面板（PaletteTreeListView）
│   ├── Transient/                   # 瞬态图形（呼吸灯/长度标注）
│   ├── DrawOverrule/                # 绘制覆盖（WorldDraw 重写）
│   ├── SelectionEx/                 # 选择集增强
│   ├── ToolBar/                     # 自定义工具栏
│   ├── ReFind/                      # 反查（FX 命令）
│   └── SharedAuxiliary.projitems
│
├── AuxiliaryClassLibrary/           # 公共类库（被全部项目引用）
│   ├── SettingsManager/             # 设置管理（XML 持久化，单例）
│   ├── SharedMemoryManager/         # 共享内存（进程间通信）
│   ├── ComHelper/                   # Excel COM 辅助（ExcelInteropHelper）
│   ├── Notifications/               # 消息框/通知（GclMessageBox/GclNotify）
│   ├── RegditEx/                    # 注册表操作（插件注册/路径检测）
│   ├── TestLog/                     # 诊断日志
│   └── AuxiliaryClassLibrary.vbproj
│
├── SettingsManagerTests/            # 单元测试项目（MSTest：设置/自检规则/排序/计时）
├── scripts/                         # 工程脚本
│   ├── build-all.ps1                # 14 个 CAD 版本批量构建（结果汇总）
│   ├── set-version.ps1              # 版本号统一管理（校验/更新）
│   └── check-encoding.ps1           # 编码一致性检查（UTF-8 BOM）
├── Lisp/                            # AutoLISP脚本（Main.lsp / Reactor.lsp）
├── Setup/                           # Inno Setup 安装脚本（工程量计算书.iss）
├── dist/                            # 发布产物（安装包）
├── 工程量计算书.sln                 # 解决方案文件
├── CLAUDE.md                        # 代码规范指南
└── README.md                        # 本文档
```

## 开发指南

### 解决方案配置
- 使用 `工程量计算书.sln` 打开解决方案
- 支持 Debug 和 Release 两种配置
- 支持 x86 和 x64 平台编译

### 编译顺序
1. 首先编译 `AuxiliaryClassLibrary`
2. 然后编译各 `ClassLibraryAutocad*` 项目
3. 最后编译 `HomeStart` 和 `工程量计算书(EXCEL)`

### 调试技巧
- 使用 HomeStart 作为启动项目
- 调试Excel插件时，附加到Excel进程
- 调试AutoCAD插件时，附加到acad.exe进程
- 查看输出窗口获取调试信息

### 工程脚本
```powershell
# 批量编译全部 14 个 AutoCAD 版本插件（共享代码改动后必跑，防漏编）
.\scripts\build-all.ps1

# 版本号统一管理（发版前 -CheckOnly 校验，发版时 -Version 更新）
.\scripts\set-version.ps1 -CheckOnly
.\scripts\set-version.ps1 -Version 3.0.1.0

# 编码一致性检查（源文件必须 UTF-8 BOM，防止 VS 乱码）
.\scripts\check-encoding.ps1

# 运行单元测试（SettingsManagerTests 项目）
dotnet test SettingsManagerTests\SettingsManagerTests.vbproj
```

### 代码规范
参考项目根目录下的 `CLAUDE.md` 文件



## 常见问题解答

### Q: 支持哪些AutoCAD版本？
A: 支持AutoCAD 2012至2025的所有版本，安装包会自动适配您的CAD版本。

### Q: 同时支持32位和64位系统吗？
A: 是的，程序会自动检测并适配32位或64位的Excel和AutoCAD。

### Q: 数据安全吗？
A: 程序会自动备份文件到"工程备份"文件夹，可以通过设置调整备份策略。

### Q: 可以导入其他格式的计算书吗？
A: 当前版本主要支持本软件创建的Excel格式文件，未来会考虑增加更多格式支持。

### Q: 如何反馈问题或建议？
A: 可以通过以下方式联系：
- QQ: 54034869
- E-mail: goddcw@163.com

## 许可证信息

本项目版权所有，保留所有权利。

## 版本历史

### 3.0.1.0 - 3.0.11.0 (当前最新版本 3.0.11.0)

本版本线包含自 3.0.0.0 以来的历次迭代，按功能主题归类如下。版本号由 `scripts\set-version.ps1` 在发版时统一同步至 17 个 `AssemblyInfo.vb` + `ClassLibraryAutocad2025.vbproj` + 3 处 `ApplicationVersion` + `Setup\工程量计算书.iss` 的 `#define AppVersion`。

**新增「建工计算器」独立项目（ConstructionCalculator）**
- 新增独立计价工具，Ribbon 新增「造价工具」按钮，Excel 加载项定位并启动/前置该程序；Inno 打包直连同 AutoCAD 做法，补齐 DevExpress 依赖 DLL
- 配置驱动通用计算框架：XML 架构 + XSD 强校验 + 公式引擎 + 表格计算页；新增风管导流叶片、防腐保温（阀门/法兰绝热、给水/排水铸铁管、UPVC 管等）、桥架支架、型钢/钢管共多类计算页
- 计算页锁定/增删行控制、可编辑单元格浅黄高亮、批量输入右键菜单、镀锌附加系数、定位单价（元/吨→元/米）等
- 设置窗口并入 `FormConfig`（PropertyGridControl 绑定），数据持久化到程序同目录 `UserData.xml`；导出 Excel、导航字体、说明页自适应等

**DevExpress 迁移与界面统一**
- DevExpress 库版本由 26.1 降级为 23.2，更新 `licenses.licx` 授权文件；部分项目纳入 licenses.licx
- 设置窗体、消息框、关于对话框等迁移至 DevExpress 原生控件；消息框语义规范化（成功/确认绿色配色、标准警示红）
- 新增 GclWaitForm 等待动画、SuperTip 超级提示工具（HTML/CSS 统一样式，绿色圆角边框 + 阴影，修复 Tooltip 图标嵌入与 Ribbon/控件 HTML 渲染）
- 导航分组增加 SVG 图标；启动页颜色/亮度调整；皮肤运行时统一 Office 2019 Colorful + Forest

**Excel 插件核心**
- 新增查找替换/批量操作防抖机制：单格编辑立即同步处理（批注 / `=`/`+` 剥离 / 逐行计算）、批量替换停顿后整表重算 + 刷新序号，事件侧零开销、防死循环；`RefreshSheet` 并行度限制为 CPU 核数
- 阅读模式高亮、快速检索（FastSearch：汉字拼音首字母匹配 / 原生附加过滤 / 关键字高亮 / 防抖）、查找窗体（右键按活动单元格值筛选、空值查找、隐藏空白列）
- 表达式计算器优化、Find、汇总、反查等功能稳定性改进
- 新增 UpdateService 后台更新服务（无窗体系统通知 / 静默升级 / 重复命令去重 / IPC 命名管道 / 本地 SHA256 校验）；编码守卫 `ensure-utf8-bom`

**AutoCAD 模块**
- 瞬态呼吸灯重构、多对象匹配（空间匹配 + 块转换）、过滤选择（GL/GLS 命令）、快速选择、图层控制、ReadOverrule 重绘
- 大量性能 / 稳定性 / 资源泄漏 / 崩溃修复；`ConvertToBlock` 批量事务；`SpatialHash` 几何防护
- `ProcomToExcel` 平台兼容；`EntityComparer` 性能与健壮性优化

**工程 / 发布**
- 新增打包发布方案与 Inno Setup 安装脚本；发布流程重构（编译 → VSTO 发布 → 组装 → 打包 → 校验）；发布产物直连打包
- 单元测试项目与工程脚本；版本号同步机制纳入 ISS 安装脚本；编码一致性检查（UTF-8 BOM）
- 修复：GclMessageBox 弹窗无 owner 时非模态可致关闭宿主崩溃；VSTO 发布签名证书指纹恢复

### 3.0.0.0

**主要更新：**
- 全项目引入 DevExpress 控件库引用，主要窗体与对话框完成 DevExpress 迁移（XtraForm、XtraMessageBox、SimpleButton、TextEdit、CheckEdit、SpinEdit 等）
- 启动器重构：改用 DevExpress SplashScreenStart + WebView2 渲染 HTML5 启动动画，新增主标题流光/呼吸灯特效
- 大量 AutoCAD 模块稳定性、性能修复与代码重构

**AutoCAD模块：**
- 修复文档切换时瞬态图形导致 AutoCAD 崩溃的 Bug
- 新增图层反转命令 GCL_LayerInvert（测试版），优化嵌套实体容器追溯支持
- 修复多段线弧形段标注按直线距离计算的问题，瞬态图形增加对弧形段、圆、圆弧的支持
- 修复 WorldDraw 重绘中 LayerId 赋值偶发 eInvalidInput 崩溃
- 修复瞬态线宽状态不一致问题；优化调色板 UI 文本
- 修复图层0状态下 CC 画线不显示的问题；修复重绘逻辑漏洞；优化图层操作逻辑，对图层0进行智能操作
- 优化瞬态显示性能，修复 SelObjFormIds 空值检查与锁定图层处理逻辑
- 重构 DrawOverrule：修复性能瓶颈、清理死代码与注册冲突
- 重构 GCLRibbon：修复单例模式、防止重复初始化并清理命名
- 修复 EntityComparer 哈希码一致性问题，优化比较逻辑与属性读取方式
- 修复多对象匹配中实体跨组重复的 Bug，添加性能分析报告
- 修复自定义面板中对象列表逆序排列的问题
- 简化 GCL_Purge 方法，改用 SendStringToExecute 发送原生 PURGE 命令替代事务处理
- 修复 TreeListView 空引用错误与 SetVariableLWDISPLAY 失效问题；优化 TreeListView 性能
- 简化编辑立管流程（InputBox 直接输入保存），修复 TreeListView 刷新问题
- TreeListView 右键菜单新增"编辑立管"功能，支持批量编辑点对象扩展数据
- 修复批量编辑立管时因 DataGridView 排序导致修改错误行的 Bug
- 修复 SetPointXdataBatch 缺少文档锁导致 eLockViolation 异常
- 修复 BlockName 静态字段生命周期问题：每次 SMatch 调用生成新的块名前缀
- 将 FrmPoint 改为单实例模式，新增 DocumentToBeActivated 事件处理清理瞬态
- 修复启动时字段初始化空引用问题，新增字体文件存在性检查与回退机制
- 调整解锁全部图层代码，注释不必要的操作以提高性能
- 将查找精确匹配默认值改为 False
- 重构 MText 显示配置，提取 ConfigureMTextDisplay 公共方法消除重复代码
- 调整默认设置值：启用标注编号和字体路径功能，优化匹配容差和刷新间隔参数
- 修复自定义字体文件未能正确加载的问题
- 修复 ModifyFonts 字体替换逻辑：移除 CheckFontExists 误判、清空 TypeFace、BigFont 不再无条件覆盖
- 新增 DeepSeek 图标资源，更新 AutoCAD 2025 程序集引用
- 修复反查 CAD 按钮事件绑定；COM 错误消息明确程序名；清理测试代码
- ClassLibraryAutocad2016 全部 7 个窗体迁移至 DevExpress 控件 + XtraForm
- 查找窗体迁移至 DevExpress RibbonForm
- 移除 7 个未被引用的残留窗体（FormMsgBox/FormSet/FrmEditXdata/FormScreen/Form型钢计算/Form钢管计算/Form防腐绝热），随后恢复被 Command.vb 引用的 FrmEditXdata 窗体
- 修复窗体迁移过程中的各类编译错误（Imports、控件属性兼容性等）

**Excel与公共辅助库：**
- 刷新工作表并发计算以提高效率
- 重构热键管理并增加焦点检查，防止非 Excel 窗口误触发热键
- 重构 ThisAddIn 命名规范，统一采用 VB.NET 标准命名约定
- 将 MessageBox/MsgBox 对话框替换为 DevExpress XtraMessageBox（含 AuxiliaryClassLibrary 全局 MsgBox）；移除无用的 ThisAddIn.cs
- 主要窗体迁移至 DevExpress 控件：Form清理样式、FormInsertStr、AboutBox、FormScreen、Form计算器、FormScreenSet、Form设置、Form汇总、Form属性、Form查找、FastSelectForm、FrmToast 等
- Form清理样式进度条改用 DevExpress 原生 ShowTitle 显示百分比
- 调整 Excel 插件代码文件路径结构

**HomeStart 启动器：**
- 使用 DevExpress SplashScreenStart 替代原 HomeStart 主窗体
- 启动窗体改用 WebView2 渲染 HTML5 动画
- 修复启动状态文本仅显示文件名而非完整路径的问题
- 修复启动窗体 WebView2 初始化时序及备份路径注册表写入逻辑
- 修复启动白屏（Opacity + Environment 预加载 + NavigationCompleted）
- 主标题"工程量计算书"增加呼吸灯发光与渐变流光特效

**安装程序与工程：**
- 修复安装程序多项问题（发布者名称、安装路径、自修复、SDK文件泄漏、卸载清理）
- 修复安装后 DevExpress 界面无法汉化，新增 AssemblyResolve 卫星程序集解析
- 移除 QuantityCalculation 项目
- WPF 程序集引用标准化，移除冗余项目引用，新增资源

### 2.0.3.23
- 更新版本号
- 为反查窗口增加插入功能，可以将选择集的对象写回Excel
- 优化反查窗口，增加对象隐藏时的处理
- 优化事件处理过程
- 修复快速选择命令的错误（空对象引用导致主程序崩溃）
- 为反查窗口添加右键菜单（定位、写入）
- 修复对象不存在时反查出错的BUG
- 切换文件时清空反查窗口
- 调整选择范围，只选择属性
- 修复Zoom函数的一处错误
- 重写重置视图方法

### 2.0.3.22
- 全面代码审查修复
- 修复30+严重/一般错误
- 清理冗余文件
- 优化性能
- 移除DevExpress相关配置代码
- 添加SettingsManager自定义设置管理器
- 添加ObjectListView包
- 添加新的Palette组件
- 修复自定义设置管理类的若干问题

### 2.0.3.21
- 解决Excel与AutoCAD不在同一权限下无法获取主程序实例的问题（解决失败，暂时不考虑）
- 添加弹出式渐变窗口以替代部分消息框
- 仅向用户展示结果信息
- 调整图片显示窗口效果
- 优化闪烁显示性能
- 增加重复对象检测

### 2.0.3.20
- 重写全局快捷键
- 取消钩子函数
- 删除防重入标志
- 重写图片批注功能
- 重写自定义信息功能
- 一些代码清理和优化
- 添加版本控制功能

### 2.0.3.19
- 使用全局热键完全替代钩子函数
- 清理部分无用代码
- 解决了Ctrl+E按键组合会在单元格留下E字符的问题
- 解决了按下自定义快捷键之后先检查Excel状态是否繁忙

### 2.0.3.18 (SP2)
- 使用全局热键替代部分原钩子函数定义的按钮
- 解决了Ctrl+E按键组合会在单元格留下E字符的问题

### 2.0.3.18 (SP1)
- 清理无效引用代码

### 2.0.3.18
- 重构动态显示功能
- 重构向Excel单元格发送数据的代码
- 清理Excel Addin中无效的代码
- 优化多重匹配代码

### 2.0.3.17
- 对主启动器进行了一些优化
- 对Excel安装路径检测进行了一些优化
- 完全重写多重匹配功能
- 生成安装包

### 2.0.3.16 (多个SP版本)
**主要更新：**
- 添加对AutoCad2022~2024的支持
- 添加对AutoCad2025的支持（虽然AutoCad2025基于.NET Core平台，暂时不支持与Excel交互）
- 重新整理了代码结构
- 修复一些bug

**AutoCAD模块：**
- 重新设计设置窗口界面布局
- 完善多对象匹配功能，可以设置在匹配完成后原地插入块
- 优化重绘功能的代码
- 修复当反查结果为空、开启闪烁效果导致程序崩溃的错误
- 调整闪烁功能
- 修复重绘过程中的错误
- 优化快速选择代码，向用户展示可视化进度条并可中途取消操作
- 修改过滤选择命令设置，使其可对锁定图层对象进行操作
- 修正多重匹配逻辑错误
- 优化闪烁显示效果
- 重写多重匹配，可实现在匹配对象处原地生成块
- 新增GLS命令，可让用户选择单个或多个源对象过滤，并返回不重复过滤结果，添加进度条并监控用户按键，按ESC可终止过程
- 修改并优化EntityComparer比较器，增加对Leader的支持，重新启用单行/多行文本的支持
- 添加一个命令可一次性解锁全部图层
- 修复立管编辑窗口中批量修改功能BUG
- 修复一处逻辑错误（意外释放xlapp对象导致FX命令无效）
- 完善打开文件缺失字体功能
- 修复重绘过程中可能导致程序出现致命错误的问题（但牺牲一定性能）
- 修复特定情况下无法引用Excel.Application对象的问题
- 添加对属性定义对象(AcDbAttributeDefinition)的支持，支持TQ、GG、GL、GLS命令
- 优化GL命令，现在可选择多个源对象过滤
- 添加共享项目SharedAuxiliary，逐步将复用代码转换至该项目
- 完善重绘过程WorldDraw代码，修正由于重绘导致AUTOCAD主程序崩溃问题
- 用纯代码实现功能区自定义，防止因CUIX文件处理不当导致程序加载失败
- 将初始化共享内存移至启动Excel之前，且在退出程序之后清理共享内存文件

**Excel模块：**
- 添加反查快捷键CTRL+R
- 修复在活动单元格变化之后不会更新活动单元格信息的问题
- 修改新建工程逻辑，当版本号为6时采用新的列表样式
- 完善文件自动备份逻辑
- 重写定时保存代码
- 在重建索引和重算工作表过程中加入函数防护处理
- 修改单实例窗口入口函数
- 修改提示框，由msgbox改为messagebox.show函数
- 修复在汇总窗体第一次打开时在计算表中操作穿透问题

**HomeStart：**
- 关闭主窗口后自动清理共享内存文件

### 2.0.3.15 (SP13)
- AutoCAD：重写快速选择功能，使其可适应多种组合情况下的筛选
- AutoCAD：计划增加多图元识别功能（从最简单的块、文本标识开始）
- AutoCAD：在新建或打开图形文件时更新XdataId池，避免生成的XdataId与已打开的工作表中出现重复元素
- AutoCAD：修复立管编辑窗口中批量修改功能BUG
- AutoCAD：修复一处逻辑错误（意外释放xlapp对象导致FX命令无效）
- AutoCAD：完善当打开文件缺失字体的功能
- AutoCAD：修复重绘过程中的一处错误
- Excel：在重建索引和重算工作表过程中加入函数防护处理
- 通用：修改单实例窗口入口函数
- 通用：修改提示框，由msgbox改为messagebox.show函数
- AutoCad：新增GLS命令

### 2.0.3.14
- 界面：调整部分窗口上控件的TAB键顺序
- 计算器：修复四则运算计算器的一些BUG
- Excel：修改定时保存逻辑，定时保存仅对工程量计算书文件生效
- Excel：将重建列表行索引改为异步方式执行，以提高界面流畅度（后取消异步方式，但在原方法基础上添加防重入设置，避免短时间内多次执行该过程以提高性能）
- Excel：删除行支持多个不连续的区域
- AutoCad：文本提取(GetText、GetTexts)可支持多重引线
- AutoCad：执行GG、TQ命令时，当对象处于锁定图层内时，现在可先保存图层锁定状态、解锁图层后执行操作，完成之后恢复图层锁定状态
- AutoCad：GG命令增加多重引线(AcDbMLeader)对象的支持
- AutoCad：新增GetTexts命令，可对多个选择的文本按坐标排序后连接成一个字符串写入Excel单元格
- AutoCad：优化GL命令查询结果显示
- AutoCad：打开图形文件时自动检测缺失的字体并替换为默认的字体文件

### 2.0.3.13
- QC命令优化
- 临时隐藏/显示按钮图标异常修复
- 定时保存提醒逻辑修正
- GL命令选择天正文本缩放修复
- GetCrossingPolygon函数逻辑错误修复
- 高亮选择单元格所在行优化

### 2.0.3.12
- 重写备份文件过程
- 修改向Excel发送数据方法，将字符集改为GB18030
- 修复备份文件夹不存在时点击管理备份按钮无反应的问题
- 禁用MoveTops对象置顶功能
- 新增QCBZ命令，清除选定图元的提取标记
- AutoCad选项卡增加立管尺寸设置
- AutoCad选项卡增加恢复本页设置按钮
- 修正反查CAD对象在选择集内逆序排列逻辑
- 重写RandHexs函数
- 添加公共辅助类库AuxiliaryClassLibrary
- 调整批注写入逻辑
- 添加NotificationHelper模块
- 添加EntExtents3d类
- 激活文档时检查视图是否为正视图
- 重写TQ命令
- 添加图层控制命令组
- 对CUI外观重新设计排版
- 添加过滤选择(GL)命令
- 重写GETTEXT命令

### 2.0.3.11
- 重写RandHexs函数
- 调整批注写入逻辑
- 添加NotificationHelper模块
- 激活文档时检查视图是否为正视图

### 2.0.3.9
- 调整Excel功能区布局
- 调整截图批注部分显示效果
- 修正TQ命令筛选锁定图层时的错误
- 修正TQ命令筛选多行文本时Contents属性处理
- 增强GetText命令功能
- 新增FX命令，选中CAD图元后反查Excel中引用该图元的单元格
- 修改BZLG命令输入法处理
- 修改清除批注过程，数量大于1时提示用户
- 修改HomeStart监视方法
- 自定义工具栏改为加载CUIX文件
- 重写向Excel发送数据方法

### 更早版本
- 支持多版本AutoCAD
- 完善反查功能
- 增加图片批注功能
- 添加型钢、钢管、防腐绝热计算器
- 实现表达式计算引擎

---

**感谢您使用工程量计算书！**
