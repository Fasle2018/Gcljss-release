# 工程量计算书 · 发布版本

本仓库（`Fasle2018/Gcljss-release`）仅存放 **工程量计算书** 的正式发布安装包，供用户下载与自动更新。**不含源码**；源码与开发信息请访问主仓库，如需反馈问题请通过 README 底部的联系方式。

---

## 立即下载最新版

[![下载最新版](https://img.shields.io/badge/安装包-最新版-blue)](https://github.com/Fasle2018/Gcljss-release/releases/latest/download/Gcljss_Setup_3.0.20.0.exe)
[![最新版本](https://img.shields.io/github/v/release/Fasle2018/Gcljss-release?label=最新版本)](https://github.com/Fasle2018/Gcljss-release/releases/latest)
[![下载总量](https://img.shields.io/github/downloads/Fasle2018/Gcljss-release/total?label=下载总量)](https://github.com/Fasle2018/Gcljss-release/releases)

> 上面第一个链接使用 GitHub 的 `releases/latest/download/<文件名>` 稳定地址，**永远指向最新正式版安装包**，无需记住版本号。版本徽章与会自动更新，无需手动维护。

- 安装包文件名：`Gcljss_Setup_3.0.20.0.exe`（约 44 MB，单文件安装包）
- 版本：**3.0.20.0**（当前最新正式版）
- 下载：点击上方按钮，或访问 [Releases 页面](https://github.com/Fasle2018/Gcljss-release/releases)
- **自动下载页**：[https://fasle2018.github.io/Gcljss-release/](https://fasle2018.github.io/Gcljss-release/)（GitHub Pages 着陆页，自动读取最新版本并提供一键下载）

---

## 系统要求

| 项 | 要求 |
|---|---|
| 操作系统 | Windows 10 及以上（安装包强制 64 位） |
| Excel | Microsoft Excel 2010 及以上（32 位或 64 位） |
| AutoCAD | AutoCAD 2012 - 2025（64 位） |
| .NET Framework | .NET Framework 4.8 或更高（AutoCAD 2025 插件依赖 .NET 8，由 CAD 自带） |

## 安装方法

1. 下载 `Gcljss_Setup_<版本>.exe`。
2. 运行安装程序（安装到 `%ProgramFiles%\工程量计算书`，自动创建桌面/开始菜单快捷方式）。
3. 安装完成后双击桌面快捷方式"工程量计算书"启动，程序会自动注册 Excel 与 AutoCAD 插件。
4. 首次使用在 Excel 中允许加载 VSTO 插件；在 AutoCAD 中输入命令加载插件。

## 更新

- 程序内置「检查更新」功能：Excel 功能区的「检查更新」按钮 → 自动从本仓库 `releases/latest` 拉取新版并下载安装。
- 安装包支持**覆盖安装**：新版本号 > 当前版本即可升级；重复安装同一版本会自动卸载旧版再安装，无文件/注册表残留。

---

## 版本历史（近期）

### 3.0.20.0（当前最新）
- 统一「关于」窗体：以 `UpdateService/UpdateForm` 窗体的模板为样式基准，新增 `AboutForm`（DevExpress DirectXForm + HTML/CSS 重绘），Excel 与 AutoCAD 共用统一关于对话框。
- 查找窗体数据区右键菜单：删除「筛选{值}」，新增「限定:{值}」与「反查」（功能与功能区反查按钮一致）。
- 启动检测到待安装更新时改为弹窗询问（是否立即更新），不再静默安装造成困惑。

### 3.0.19.0（近期）
- **VSTO Excel 加载项改「扁平注册表式部署」**：去掉 ClickOnce 的 `Application Files\*.deploy` 冗余结构，`.vsto`/`.manifest`/加载项 DLL 与依赖 DLL 同装程序根目录、与清单同目录就地解析；安装包更精简、加载更稳。
- 发布资源目录调整（静态资源独立为 `publish\ReleaseProduct`）；移除无用数据文件 `工程量计算书数据库.xml`。

### 3.0.18.0（近期）
- **新增查找替换/批量操作防抖机制**：单格编辑立即同步处理（批注 / `=`、`+` 剥离 / 逐行计算），批量替换停顿后整表重算 + 刷新序号；事件侧零开销、防死循环；`RefreshSheet` 并行度限制为 CPU 核数。
- 插入/删除行后直接编辑单元格不再报错；`InsertRow` 偶发 COM 失效容错。
- 修复查找替换批量处理时反复重算导致的严重卡顿。

### 3.0.9.0 - 3.0.10.0（基础）
- 版本号同步机制纳入安装脚本；新增 UpdateService 后台更新服务（无窗体系统通知、静默升级、重复命令去重、IPC 命名管道、本地 SHA256 校验）。

### 3.0.1.0 - 3.0.8.0（系列迭代）
- 新增「建工计算器」独立项目（风管导流叶片、防腐保温、桥架支架、型钢/钢管等多类计算页），Ribbon 新增「造价工具」按钮。
- DevExpress 迁移与界面统一（库版本降级 23.2、控件/消息框迁移、GclWaitForm、SuperTip 超级提示、SVG 图标）。
- README 细节优化：阅读模式、快速检索、表达式计算器、查找窗体等增强。
- 发布流程重构（编译 → VSTO 发布 → 组装 → 打包 → 校验），Inno Setup 单文件安装包，单元测试与工程脚本。

> 完整的版本历史（含 3.0.0.0 及更早，以及各功能模块明细）见主仓库 `README.md` 的「版本历史」章节。

---

## 联系方式 / 反馈

- 反馈问题或建议，请提供：Excel 版本、AutoCAD 版本、操作步骤、复现现象。
- 我们会在新版本中优先处理影响面大的问题。

> 本仓库仅作安装包分发。软件版权归开发者所有，保留所有权利。
