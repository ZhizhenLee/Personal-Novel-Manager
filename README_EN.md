# Personal Novel Manager（个人小说管理 · 通天录）

> Portable, single-file Windows desktop tool for organizing and writing novels.
> 免安装、单文件的 Windows 桌面小说编写管理工具。

[English](README_EN.md) | [简体中文](README.md)

## Project Description

**Personal Novel Manager** (software name “通天录”/Tongtianlu) is a **local novel-writing and organization tool** built for web-novel authors. Its core idea is **“one novel = one main folder.”** A three-pane interface helps you efficiently manage drafts, chapters (body text), character profiles, worldbuilding, and per-chapter titles/summaries for one or many novels — plus RAR import/export for easy backup and migration.

Highlights:

- **Portable & single-file**: double-click `通天录.exe` and run — no install, no environment setup, no extra files.
- **100% local**: all data is stored as plain files in folders you control. Open them with any editor anytime; nothing is “locked away” by the app.
- **Clear structure**: body text (one file per chapter), drafts, characters, worldbuilding, chapter titles/summaries — all visible at a glance.
- **Auto-save**: optional toggle; saves automatically on switch/close.
- **RAR round-trip**: export a whole novel to a RAR in one click, and restore/replace a novel from a RAR in one click.
- **Good looks**: cover thumbnails, card-style novel list, and previews for images and styled xlsx files.

---

## Features

- **Novel management**: each novel is a main folder whose name is the novel name. The home screen shows a card list of novels (cover thumbnail + novel name); click one to enter its content management.
- **Three-pane content UI**: left = main folders (body/drafts/characters/worldbuilding/chapter-title-list/chapter-summary-list/other files); middle = file list of the selected folder; right = text editor.
- **Standard folder layout** (auto-created when you create a novel):

  ```
  NovelName/
    drafts/           Multiple draft files, default name yyyyMMddHHMM, renamable
    cover/            Cover image, shown as thumbnail, can be changed
    characters/       One .md file per character (name = filename)
    body/             One txt/md file per chapter, e.g. chapter_003.txt
    worldview.md      Worldbuilding notes
    chapter-titles.md     Per-chapter titles
    chapter-summaries.md  Per-chapter summaries
    _backup/          Version backup (auto before overwrite; excluded on export)
    other/            Unclassified loose files in the novel folder
  ```

  > Actual on-disk folder names use Chinese labels (草稿/封面/角色库/正文/世界观.md/章节标题列表.md/章节简介汇总.md/_备份/其他文件); the layout above is the equivalent structure.

- **Write & preview**: edit drafts, body text, characters, worldbuilding, chapter titles and summaries; live word count; UTF-8 read/write (with GBK fallback for reading).
- **One-click next chapter**: auto-detects the highest chapter number and creates the next `chapter_NNN`.
- **Configurable novel root path**: pick the folder via “Set Path”; it is remembered and restored on next launch.
- **Import RAR to replace a novel folder**: pick a `.rar`, replace an entire existing novel with its contents (extract → replace → refresh). Uses the local WinRAR / 7-Zip; the extractor path is configurable and remembered.
- **Export novel as RAR**: pack a whole novel into `NovelName.rar` (optionally excluding `_backup`). Requires WinRAR to create RAR (7-Zip cannot create RAR).
- **Cover upload**: choose a local image as the novel cover; thumbnails shown in the UI.
- **File preview**: preview txt/md text, images (covers/drafts), and styled xlsx sheets (background color, borders, font color, merged cells, alignment).

---

## Tech Stack

- **Language / UI**: Python 3.13 + Tkinter (standard library)
- **Spreadsheet preview**: openpyxl (optional; falls back to a minimal standard-library zipfile+xml parser if absent)
- **Images**: Pillow (optional; used for cover scaling)
- **Packaging**: PyInstaller (`--onefile --windowed`)
- **Platform**: Windows 10/11 (x64)

> At runtime, **no third-party dependency is required** beyond Python's built-in Tkinter. Pillow / openpyxl are optional enhancements.

---

## Install & Usage

### Option 1 — Run the portable exe (recommended)

1. Download `通天录.exe` from [Releases](../../releases) to any folder (or run `dist\通天录.exe` in this repo).
2. Double-click to run.
3. First use: menu “File → Set Novel Root Folder” and choose a folder to hold your novels.
4. Menu “File → New Novel”, enter a name; the full structure is created automatically.

### Option 2 — Run from source

```bash
# Install dependencies (optional; only Tkinter is required)
python -m pip install -r requirements.txt

# Run
python novel_manager.py
```

### Build a single-file exe

Run in PowerShell:

```powershell
./build_exe.ps1
```

Output: `dist\通天录.exe` (single file, no install).

---

## Repository Layout

```
personal_novel_manager/
├── novel_manager.py      # Main program (all logic & GUI)
├── build_exe.ps1         # Build script
├── requirements.txt      # Dependencies (pyinstaller, pillow, openpyxl)
├── app.ico               # App icon
├── dist/                 # Built 通天录.exe (ignored by .gitignore by default)
├── README.md             # Simplified Chinese readme
├── README_EN.md          # English readme (this file)
└── LICENSE               # MIT License
```

---

## Release Checklist

- [x] License chosen (`LICENSE`, MIT)
- [ ] Add **screenshots** (novel list home screen, three-pane editor) under `docs/screenshots/` and reference them in the README
- [ ] Set the **GitHub repository description** (see suggested description below)
- [ ] Add release / issue / PR templates (optional)
- [ ] Decide whether to commit `dist/` (recommended: publish the exe as a GitHub Release rather than committing it)

### Suggested GitHub repo description

> Portable single-file Windows tool for organizing & writing novels — drafts, chapters, characters, worldbuilding, RAR import/export.

---

## License

[MIT License](LICENSE)
