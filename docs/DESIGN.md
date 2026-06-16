# 詩歌冊搜索系統 — 設計文檔

## 1. 架構總覽

```
┌─────────────────────────────────────────────────────────┐
│                    QApplication                          │
│  ┌───────────────────────────────────────────────────┐ │
│  │                  MainWindow                          │ │
│  │  ┌──────────┐  ┌────────────────────────────────┐ │ │
│  │  │ Sidebar  │  │ Search Area + Settings         │ │ │
│  │  │ book_list│  │ File List (results)            │ │ │
│  │  │ 240px    │  │                                │ │ │
│  │  └──────────┘  └────────────────────────────────┘ │ │
│  └───────────────────────────────────────────────────┘ │
│                                                          │
│  浮動元件（無父視窗）：                                    │
│  • DropdownList — 書名建議                               │
│  • BookNameOverlay — 每螢幕一個 _BookNameOverlayPanel      │
└─────────────────────────────────────────────────────────┘

背景執行緒：
  ContentSearchWorker  — 關鍵字搜索
  IndexBuildWorker     — 建立索引
  threading            — Word 快捷鍵、文件監聽、顯示切換
```

---

## 2. 模組設計

### 2.1 分層說明

| 層級 | 模組 | 職責 |
|------|------|------|
| UI | MainWindow, DropdownList | 佈局、事件、狀態 |
| 呈現 | THEMES, build_*_style | 視覺主題 |
| 業務 | scan_books, search_*, open_* | 掃描、搜索、開檔 |
| 基礎設施 | Index 快取、Workers、Overlay | 效能與系統整合 |

### 2.2 搜尋模式狀態機

```
search_mode ∈ { 'book', 'global', 'keyword' }

書本模式 → 顯示 book_row，右欄顯示當前書冊檔案
全局模式 → 顯示 global_row，右欄顯示跨書冊結果
關鍵字模式 → 顯示 keyword_row + index_row，右欄顯示全文命中
```

切換模式時停止進行中的搜索/索引 worker，避免競態。

### 2.3 開檔流程（Word + 可選同步畫面）

```
使用者點擊結果
    │
    ▼
_show_open_book_popup()     ← 1. 顯示書名浮層（各螢幕）
    │
    ▼
open_docx() / open_pdf()   ← 2. 開啟閱讀器
    │
    ▼
_schedule_word_focus_keys() ← 3. Word：Enter → Alt+W → Alt+O
    │
    ▼
_maybe_duplicate_displays() ← 4. 若勾選：切換同步畫面
    │                         （先 snap overlay，再 DisplaySwitch）
    ▼
BookNameOverlay 監聽         ← 5. 每 1.5s 檢查檔案是否仍開啟
    │
    ▼
關閉 Word/PDF → hide_overlay()
```

---

## 3. UI 設計

### 3.1 佈局原則

- **左固定、右彈性**：側欄不可拖曳，避免誤操作
- **設定可摺疊**：預設隱藏，減少非全螢幕時垂直空間壓力
- **設定可捲動**：`QScrollArea` + 動態 `maximumHeight`
- **雙欄設定**：「外觀」與「開啟 Word/PDF 時」並排

### 3.2 主題系統

- `THEMES['dark'|'light']` 色票字典
- `_apply_theme()` 為單一入口，更新 palette + 各元件 stylesheet
- 模式按鈕、操作按鈕、GroupBox 分開定義樣式

### 3.3 書名浮層（Overlay）

```
BookNameOverlay（管理類）
    └── _BookNameOverlayPanel × N（每個 QScreen 一個）

顯示模式：
  corner              — 直接右下角
  center_then_corner  — 置中 5s → 動畫縮小移至右下角

動畫：QPropertyAnimation(geometry) + 插值字體/邊距/圓角
```

**多螢幕策略：**

- 延伸模式：每螢幕各一浮層
- 切換同步後：僅主螢幕保留一個浮層（避免重疊錯位）

---

## 4. 索引設計

```
build_content_index()
    └── 對每個可搜索檔案 extract_file_chunks()
            ├── PDF: 每頁文字
            ├── DOCX: 每段落
            ├── ODT: content.xml 段落
            └── DOC: Word COM 段落（可選）

儲存：gzip JSON
    version, folder, built_at, files{ path → {book, file, ext, mtime, chunks} }

搜索：search_content_index() 線性掃描 chunks（已壓縮在記憶體）
```

---

## 5. 執行緒與信號

| 元件 | 機制 | 說明 |
|------|------|------|
| ContentSearchWorker | QThread + pyqtSignal | 搜索完成回傳主執行緒 |
| IndexBuildWorker | QThread + pyqtSignal | 索引進度與完成 |
| display_duplicate_requested | pyqtSignal | Word 快捷鍵後安全切換顯示 |
| document watch | threading + QTimer.singleShot(0) | 背景檢查、主執行緒隱藏 |

---

## 6. 打包設計

- **PyInstaller onefile**：單一 `HymnSearch.exe`
- **資源**：`--add-data "assets;assets"` 供 `app_resource_path()` 解析
- **圖示**：`assets/app.ico` 嵌入 exe 與視窗
- **`sys.frozen` / `_MEIPASS`**：運行時區分開發與打包路徑

詳見 [BUILD.md](BUILD.md)。

---

## 7. 擴展建議（未實作）

- 設定檔（JSON）取代硬編碼 `HYMN_FOLDER`
- 系統匣常駐與全域快捷鍵
- 自訂 PDF 閱讀器路徑
- 關閉檔案後自動恢復延伸桌面
