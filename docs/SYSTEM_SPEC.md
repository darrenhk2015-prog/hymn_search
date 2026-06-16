# 詩歌冊搜索系統 — 系統規格書

**版本：** Phase 1  
**主程式：** `hymn_search.py` / `HymnSearch.exe`  
**最後更新：** 2026-06

---

## 1. 系統概述

| 項目 | 說明 |
|------|------|
| 名稱 | 詩歌冊搜索系統（Hymn Search） |
| 類型 | 桌面應用程式（單機、本地檔案） |
| 技術棧 | Python 3.10+、PyQt6、PyMuPDF、python-docx |
| 目標平台 | Windows 10/11（完整功能）；macOS/Linux（基本開檔） |
| 預設資料路徑 | `%USERPROFILE%\Desktop\神家詩歌集new` |

### 1.1 目的

在本地詩歌資料夾中快速搜尋、瀏覽及開啟 Word / PDF 詩歌檔案，支援聚會投影場景（書名提示、同步畫面、Word 視圖快捷鍵）。

### 1.2 使用者角色

- **司琴 / 投影操作員**：搜尋詩歌、開啟檔案、辨識目前書冊
- **管理員**：維護詩歌資料夾結構、建立全文索引

---

## 2. 功能需求

### 2.1 書冊掃描

- 掃描 `HYMN_FOLDER` 下所有子資料夾與單一檔案（PDF / DOC / DOCX / ODT）
- 每個項目為一本「詩歌冊」，記錄：名稱、檔案列表、PDF 頁數、TOC 書籤（若 TOC 項目 > 5）
- 左側列表顯示書冊名、首數、類型標記（W=Word、P=PDF、🔖=PDF+書籤）

### 2.2 搜尋模式

| 模式 | 說明 | 輸入 | 結果 |
|------|------|------|------|
| 書本 | 過濾左側書冊 + 右側檔案/書籤 | 書冊名、檔案關鍵字 | 當前書冊檔案與 PDF 書籤 |
| 全局 | 跨所有書冊 | 歌名/關鍵字 | 各書冊匹配的書籤與檔名 |
| 關鍵字 | 全文搜索 | 任意文字 | PDF/Word 內容命中片段 |

### 2.3 開啟檔案

- **PDF**：優先 Sumatra PDF → Foxit → Adobe → 系統預設；書籤可跳至指定頁
- **Word**：`ShellExecute` 開啟；自動送出 Enter → Alt+W → Alt+O（投影視圖）
- 單一搜索結果時：按鈕顯示「Open」，Enter 直接開啟

### 2.4 書名置頂提示（Overlay）

- 開啟 Word/PDF 時於所有螢幕顯示黃色浮層（書名 + 檔案/書籤名）
- 提示位置：右下角 / 置中→右下角（含縮小動畫）
- 提示顯示：1–10 秒 / **一直顯示**
- 關閉 Word/PDF 視窗時**自動隱藏**
- 可選：開啟後切換延伸桌面為同步畫面（Windows）

### 2.5 全文索引

- 快取路徑：`{HYMN_FOLDER}/content_index.json.gz`
- 版本號 `INDEX_VERSION = 1`；依檔案 `mtime` 判斷過期
- 背景執行緒建立，支援取消

---

## 3. 非功能需求

| 項目 | 規格 |
|------|------|
| 視窗最小尺寸 | 900 × 580 px |
| 左側書冊欄 | 固定 240px，不可拖曳 |
| 主題 | dark / light |
| 字體大小 | 10–18 px |
| 響應 | 關鍵字即時搜索防抖 400ms |
| 打包體積 | exe 約 50–80 MB（含 PyQt6） |

---

## 4. 外部依賴

### 4.1 Python 套件

```
PyQt6
PyMuPDF
python-docx
```

打包額外需要：`pyinstaller`

### 4.2 系統程式

| 功能 | Windows | 其他平台 |
|------|---------|----------|
| PDF 閱讀器 | Sumatra / Foxit / Adobe / 預設 | 系統預設 |
| Word 開啟 | Microsoft Word | LibreOffice / open |
| 顯示同步 | DisplaySwitch.exe | 不支援 |
| 文件關閉監聽 | PowerShell + 行程標題 | 不隱藏 overlay |

---

## 5. 資料結構

### 5.1 書冊物件 `book`

```python
{
    'name': str,           # 書冊名稱
    'pdf': str | None,     # 主 PDF 路徑
    'docx': str | None,    # 主 Word 路徑
    'toc': [(lvl, title, page), ...],
    'bookmark': bool,      # 是否有有效書籤索引
    'pages': int,
    'hymn_count': int,
    'files': [{'name', 'path', 'ext'}, ...],
}
```

### 5.2 列表項目 payload（UserRole）

| kind | 用途 |
|------|------|
| `file` | 一般檔案 |
| `bookmark` | PDF 書籤 |
| `content_pdf` | 關鍵字命中 PDF |
| `content_word` | 關鍵字命中 Word |

---

## 6. 設定常數

| 常數 | 預設 | 說明 |
|------|------|------|
| `HYMN_FOLDER` | Desktop/神家詩歌集new | 資料根目錄 |
| `DEFAULT_OVERLAY_SECS` | 10 | 提示預設秒數 |
| `OVERLAY_DURATION_ALWAYS` | -1 | 一直顯示 |
| `CENTER_OVERLAY_SECS` | 5 | 置中階段秒數 |

---

## 7. 檔案清單

```
hymn_search/
├── hymn_search.py          # 主程式
├── assets/                 # 圖示 app.ico / app.png
├── scripts/make_icon.py    # 產生 ico
├── build_exe.ps1           # 打包腳本
├── HymnSearch.spec         # PyInstaller 規格
├── requirements.txt
└── docs/                   # 文檔
```

---

## 8. 限制與已知問題

- `HYMN_FOLDER` 需手動修改程式碼或日後改為可設定
- `.doc` 全文搜索需 Windows + Word COM（可選 pywin32）
- PyInstaller exe 首次啟動較慢；防毒可能誤報
- 同步畫面切換後視窗位置可能略為移動（Windows 行為）
