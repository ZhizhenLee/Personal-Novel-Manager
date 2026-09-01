# Personal Novel Manager（个人小说管理 · 通天录）

> 免安装、单文件的 Windows 桌面小说编写管理工具。
> Portable, single-file Windows desktop tool for organizing and writing novels.

[简体中文](README.md) | [English](README_EN.md)

## 项目描述 (Project Description)

**Personal Novel Manager**（软件名「通天录」）是一款面向网络小说作者的**本地小说编写与组织工具**。它把「一本小说 = 一个主目录」作为核心理念，通过三栏式界面帮助你高效管理一本或多本小说的草稿、正文、角色库、世界观、章节标题与简介，并支持 RAR 包整体导入 / 导出，实现小说的便捷备份与迁移。

特点：

- **免安装、单文件**：双击 `通天录.exe` 即可运行，无任何多余文件、无需环境配置。
- **纯本地**：所有数据以明文文件形式保存在你自己的目录里，随时可用任何编辑器打开，不会有「被软件锁死」的风险。
- **结构清晰**：正文（每章一个文件）、草稿、角色库、世界观、章节标题/简介一目了然。
- **自动保存**：可开关，编辑后切换/关闭时自动落盘。
- **RAR 往返**：整本小说可一键导出为 RAR、也可用 RAR 一键覆盖还原某个小说。
- **颜值在线**：封面图片缩略图、卡片式小说列表、xlsx/图片文件预览。

---
<img width="1202" height="812" alt="ScreenShot_2026-09-01_105627_288" src="https://github.com/user-attachments/assets/2ca9a824-c623-4d70-9477-5a7aa8d79ece" />
<img width="1202" height="812" alt="ScreenShot_2026-09-01_105830_262" src="https://github.com/user-attachments/assets/bdae7bcb-5536-43f3-b504-cb9146688379" />

## 功能

- **小说管理**：每个小说一个主目录，目录名即小说名。主界面以「封面图片 + 小说名」卡片展示小说列表，点击进入内容管理。
- **内容管理界面（三栏）**：左侧「主要文件夹」（正文/草稿/角色库/世界观/章节标题列表/章节简介汇总/其他文件），中间「文件列表」，右侧正文编辑器。
- **标准目录结构**（新建小说时自动创建）：

  ```
  小说名/
    草稿/          多个草稿文件，新建默认命名 yyyyMMddHHMM，可重命名
    封面/          封面图片，界面显示缩略图，可上传更换
    角色库/        多个以角色名为文件名的 .md 文件，可新建/重命名/删除
    正文/          多个 txt/md 文件，每文件一章，命名 chapter_003.txt
    世界观.md      世界观设定
    章节标题列表.md  每章标题
    章节简介汇总.md  每章简介
    _备份/        版本备份（覆盖前自动备份，导出时可排除）
    其他文件       主目录下未分类的其他文件
  ```

- **编写与查看**：编辑草稿、正文、角色、世界观、章节标题与简介，实时字数统计，UTF-8 读写（兼容 GBK 读取）。
- **一键新建章节**：自动识别最大编号并创建 `chapter_NNN` 下一章。
- **可配置小说路径**：通过「设置路径」选择小说根目录，自动记忆，下次启动沿用。
- **导入RAR覆盖小说目录**：选择一个 RAR 压缩包，用其内容覆盖某个小说的整个目录（解压→替换→刷新）。支持本机 WinRAR / 7-Zip 解压，解压软件路径可配置并记忆。
- **导出为RAR文件**：将整个小说打包为 `小说名.rar`（可选排除 `_备份`），需本机 WinRAR 创建 RAR。
- **封面上传**：从本地选择图片作为小说封面，管理界面显示缩略图。
- **文件预览**：直接预览 txt/md 正文、图片（封面/草稿）、以及带样式（背景色/边框/字体色/合并单元格/对齐）的 xlsx 表格。

---

## 技术栈

- **语言/界面**：Python 3.13 + Tkinter（标准库）
- **表格预览**：openpyxl（可选，缺失时回退到标准库 zipfile+xml 极简解析）
- **图片**：Pillow（可选，用于封面缩放）
- **打包**：PyInstaller（`--onefile --windowed`）
- **运行环境**：Windows 10/11（x64）

> 运行时除 Python 自带 Tkinter 外**无需任何第三方依赖**；Pillow / openpyxl 为可选增强。

---

## 安装与使用

### 方式一：直接运行（推荐）

1. 从 [Releases](../../releases) 下载 `通天录.exe` 到任意目录（或双击项目内 `dist\通天录.exe`）。
2. 双击运行。
3. 首次使用：菜单「文件 → 设置小说根目录」，选择一个存放小说的目录。
4. 菜单「文件 → 新建小说」输入名称，自动生成完整结构。

### 方式二：从源码运行

```bash
# 安装依赖（可选；仅 Tkinter 必需）
python -m pip install -r requirements.txt

# 运行
python novel_manager.py
```

### 打包为单文件 exe

在 PowerShell 中执行：

```powershell
./build_exe.ps1
```

产物生成在 `dist\通天录.exe`（单文件、免安装）。

---

## 目录结构

```
personal_novel_manager/
├── novel_manager.py      # 主程序（全部逻辑与 GUI）
├── build_exe.ps1         # 打包脚本
├── requirements.txt      # 打包/运行依赖（pyinstaller、pillow、openpyxl）
├── app.ico               # 程序图标
├── dist/                 # 打包产物 通天录.exe（默认由 .gitignore 忽略）
├── README.md             # 简体中文说明
├── README_EN.md          # 英文说明
└── LICENSE               # MIT 许可证
```

---

## 许可协议

[MIT License](LICENSE)

---

## English

Full English documentation is available at **[README_EN.md](README_EN.md)**.
