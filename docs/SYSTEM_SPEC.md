# 詩歌冊搜索系統 — 系統規格書

**版本：** Phase 1 + 增強功能（排程、Mobile 遙控、Session）  
**主程式：** `hymn_search.py` / `HymnSearch.exe`  
**設定版本：** `SETTINGS_VERSION = 2`  
**最後更新：** 2026-06-18

---

## 1. 系統概述

| 項目 | 說明 |
|------|------|
| 名稱 | 詩歌冊搜索系統（Hymn Search） |
| 類型 | 桌面應用程式（單機、本地檔案）+ 可選 LAN HTTP API |
| 技術棧 | Python 3.10+、PyQt6、PyMuPDF、python-docx、FastAPI、uvicorn |
| 目標平台 | Windows 10/11（完整功能）；macOS/Linux（基本開檔，無顯示同步） |
| 預設資料路徑 | `{exe 或專案目錄}/神家詩歌集`（可在設定中更改） |
| 設定檔 | `{exe 或專案目錄}/settings.json` |
| 版本號 | `version.py` → `APP_VERSION`、`BUILD_DATE`（打包時自動遞增） |

### 1.1 目的

在本地詩歌資料夾中快速搜尋、瀏覽及開啟 Word / PDF 詩歌檔案，支援聚會投影場景（書名提示、同步畫面、Word 視圖快捷鍵），並可透過手機瀏覽器遙控開檔與排程播放。

### 1.2 使用者角色

| 角色 | 職責 |
|------|------|
| **司琴 / 投影操作員** | 搜尋詩歌、開啟檔案、管理排程清單、辨識目前書冊 |
| **Mobile 使用者** | 同 Wi‑Fi 下以瀏覽器開檔、瀏覽排程、下一首 |
| **管理員** | 維護詩歌資料夾、建立全文索引、設定 API Token / 開檔策略 |

---

## 2. 架構總覽

```
┌─────────────────────────────────────────────────────────────┐
│  hymn_search.py — MainWindow (PyQt6)                          │
│  ├── 書冊掃描、三/四種搜尋模式、開檔、Overlay、索引 Worker    │
│  └── EnhancementMixin (hymn_features/mixin.py)              │
│       Session、排程、釘選、Tray、設定匯入/匯出、Remote QR     │
└──────────────────────────┬──────────────────────────────────┘
                           │ pyqtSignal（主執行緒）
                           │ _invoke_on_main（API 執行緒 → GUI）
┌──────────────────────────▼──────────────────────────────────┐
│  hymn_remote/                                                 │
│  ├── server.py    — uvicorn 背景 Thread                       │
│  ├── api.py       — FastAPI 路由                              │
│  ├── state.py     — 執行緒安全狀態、開檔策略、待確認佇列       │
│  ├── resolver.py  — 書冊/詩歌解析（無 Qt 依賴）               │
│  ├── operation_log.py — 遙控操作記錄                          │
│  └── static/mobile.html — Mobile 分頁 UI                      │
└─────────────────────────────────────────────────────────────┘

hymn_features/
├── session_store.py  — settings 正規化、排程、Session 歷史、匯出
├── fuzzy.py          — 模糊標題匹配（含可選 pypinyin）
├── preview.py        — 列表選取預覽文字
└── mixin.py          — MainWindow 增強行為 Mixin
```

### 2.1 執行緒模型

| 元件 | 執行緒 | 說明 |
|------|--------|------|
| Qt GUI | 主執行緒 | 所有 UI、開檔、QTimer |
| ContentSearchWorker | QThread | 關鍵字全文搜索 |
| IndexBuildWorker | QThread | 建立/刷新索引 |
| RemoteServer (uvicorn) | 背景 Thread | HTTP API；**不得**直接操作 Qt |
| API → 桌面 | Signal + `_invoke_on_main` | `open`/`populate`/`pending` 用 Signal；`next`/`setlist/open` 用阻塞式主執行緒調度 |

---

## 3. 桌面功能規格

### 3.1 書冊掃描

- 掃描 `hymn_folder` 下所有子資料夾與單一檔案（PDF / DOC / DOCX / ODT）
- 每個項目為一本「詩歌冊」，記錄：名稱、檔案列表、PDF 頁數、TOC 書籤（若 TOC 項目 > 5）
- 左側列表（固定 240px）分區：**釘選** → **全部**
- 右鍵書冊可 **釘選 / 取消釘選**
- 類型標記：W = Word、P = PDF、🔖 = PDF + 書籤索引
- **自動重新掃描**：`QFileSystemWatcher` 監聽資料夾變更（可關閉，防抖 800ms）

### 3.2 搜尋模式（四種）

| 模式 | `search_mode` | 說明 | 輸入 | 右側列表 |
|------|---------------|------|------|----------|
| 書本 | `book` | 過濾書冊 + 檔案/書籤 | 書冊名（Ctrl+Enter 聚焦）、詩歌號/名、Filter | 當前書冊檔案與 PDF 書籤 |
| 全局 | `global` | 跨書冊 | 歌名/關鍵字 | 各書冊匹配的書籤與檔名 |
| 關鍵字 | `keyword` | 全文搜索 | 任意文字（可即時） | PDF/Word 內容命中片段 |
| **排程** | `schedule` | 聚會詩歌順序 | 書冊 + 詩歌號 | 排程清單（可增刪排序） |

**模式快捷鍵：** F1 書本 · F2 全局 · F3 關鍵字 · F4 排程

**排程模式操作：**

| 按鈕 / 快捷鍵 | 行為 |
|---------------|------|
| 加入 | 將書冊 + 詩歌號加入 `setlist` |
| 移除 / 上移 / 下移 / 清空 | 編輯清單 |
| 開啟 | 開啟目前選中或 `setlist_index` 項目 |
| 雙擊列表項 | 開啟該排程項 |
| Alt+← / Alt+→ | 上一首 / 下一首（非排程模式時會自動切換至排程模式） |
| Ctrl+← / Ctrl+→ | Session 歷史上一首 / 下一首（書本模式） |

### 3.3 開啟檔案

- **PDF**：優先 Sumatra PDF → Foxit → Adobe → 系統預設；書籤可跳至指定頁
- **Word**：`ShellExecute` 開啟；自動 Enter → Alt+W → Alt+O（投影視圖）
- 單一搜索結果：按鈕顯示「Open」，Enter 直接開啟
- **一點即開**：設定或頂欄可切換（單擊列表即開檔）
- 開啟後寫入 **Session 歷史**、**最近詩歌**、footer「上次：…」

### 3.4 書名置頂提示（Overlay）

- 開啟 Word/PDF 時於所有螢幕顯示黃色浮層（書名 + 檔案/書籤名）
- 提示位置：右下角 / 置中→右下角（含縮小動畫）
- 提示顯示：1–10 秒 / **一直顯示**（`OVERLAY_DURATION_ALWAYS = -1`）
- 關閉 Word/PDF 視窗時**自動隱藏**（Windows PowerShell 輪詢）
- 可選：開啟後切換延伸桌面為**同步畫面**（Windows `DisplaySwitch.exe`）

### 3.5 全文索引

- 快取路徑：`{hymn_folder}/content_index.json.gz`
- 版本號 `INDEX_VERSION = 1`；依檔案 `mtime` 判斷過期
- 背景 `IndexBuildWorker` 建立，支援取消
- Footer 顯示索引狀態：`索引—` / `索引✓` / `索引!`

### 3.6 Session 與排程資料

| 概念 | 設定鍵 | 說明 |
|------|--------|------|
| 排程清單 | `setlist` | `[{book, num}, ...]` |
| 目前排程位置 | `setlist_index` | 0-based；-1 表示無選中 |
| 開啟歷史 | `session_history` | 最近 50 次開檔 `{book, num, label, ts}` |
| 歷史游標 | `session_cursor` | 0 = 最新；Ctrl+←/→ 導航 |
| 釘選書冊 | `pinned_books` | 書名列表 |
| 最近詩歌 | `recent_hymns` | 最多 5 首（`書冊 / 號`） |
| 上次開啟 | `last_opened` | 顯示於 footer |

### 3.7 其他桌面功能

| 功能 | 說明 |
|------|------|
| 檔案預覽 | 列表選取時下方 2 行預覽（可關閉） |
| 重開上次 | 頂欄按鈕；依 Session 最新一筆重開，**不顯示書名浮層** |
| 設定面板 | 四 Tab：一般 / 開檔 / Mobile / 關於 |
| 設定匯出/匯入 | JSON（`EXPORT_KEYS` 子集） |
| 系統匣 | 可啟動最小化、關閉最小化到 Tray |
| Mobile QR | 設定 → Mobile Tab 顯示 LAN URL QR Code |
| 遙控操作記錄 | 設定 → Mobile Tab 列表（記憶體，重啟清空） |
| 快捷鍵說明 | `?` 鍵開啟對話框 |

### 3.8 桌面 UI 版面

```
┌─ Header（48px）────────────────────────────────────────────┐
│ 🎵 詩歌冊搜索  vN · 日期    [↻ 重開上次] [接受請求] [開檔策略 ▼] [一點即開] [⚙ 設定] │
├─ Sidebar 240px ─┬─ Content ───────────────────────────────┤
│ 📁 詩歌冊        │ [書本(F1)][全局(F2)][關鍵字(F3)][排程(F4)]  [⚙ 設定]  │
│ 釘選/全部        │ （模式專用輸入列）                         │
│                  │ 目前選中書冊摘要                           │
│ 圖例 W/P/🔖      │ （可摺疊設定 Tab 面板）                    │
│                  ├──────────────────────────────────────────┤
│                  │ 檔案 / 書籤 / 排程清單 / 搜索結果         │
│                  │ （列表 + 2 行預覽）                        │
├──────────────────┴──────────────────────────────────────────┤
│ Footer：狀態 | 上次：… | 索引✓ | 待辦 N | 遙控：接受/暫停/關閉 │
├─ Toast ──────────────────────────────────────────────────────┤
└──────────────────────────────────────────────────────────────┘
```

**視窗：** 最小 900×580，預設約 1120×730；**無**全螢幕縮放控制。

---

## 4. Mobile 遙控 HTTP API

### 4.1 基本資訊

| 項目 | 值 |
|------|-----|
| 預設 Port | `8765`（`DEFAULT_PORT`，設定可改） |
| Base URL | `http://{LAN_IP}:{port}/` |
| Mobile 頁面 | `GET /` → `mobile.html` |
| API 文件 | `GET /api/docs`（Swagger UI） |
| 認證 Header | `X-Api-Token: {token}`（若桌面端設定了 Token；空 Token = 不驗證） |
| 啟用條件 | 桌面「啟用 Mobile API」+「接受請求」均為 ON |

### 4.2 端點一覽

| 方法 | 路徑 | 需 Token | 需 accepting | 說明 |
|------|------|----------|--------------|------|
| GET | `/` | — | — | Mobile 分頁 UI |
| GET | `/api/health` | 否 | — | 連線與狀態快照 |
| GET | `/api/books` | 是* | — | 書冊列表（自動完成） |
| GET | `/api/entries` | 是* | — | 指定書冊的檔案/書籤 |
| POST | `/api/open` | 是* | 是 | 開啟詩歌 |
| POST | `/api/next` | 是* | 是 | 排程下一首 |
| GET | `/api/setlist` | 是* | — | 取得排程清單 |
| POST | `/api/setlist/open` | 是* | 是 | 依索引開啟排程項 |
| POST | `/api/setlist/add` | 是* | — | 加入排程清單 |
| GET | `/api/log` | 是* | — | 遙控操作記錄 |
| GET | `/api/pending` | 是* | — | 待確認請求列表 |
| POST | `/api/pending/{id}/approve` | 是* | — | 確認待辦 |
| POST | `/api/pending/{id}/reject` | 是* | — | 拒絕待辦 |

\* 若桌面設定了 `remote_token`，則必須帶正確 Token，否則 `401`。

### 4.3 端點詳細規格

#### `GET /api/health`

無需 Token（用於 Mobile 初次連線檢測）。

**回應範例：**

```json
{
  "ok": true,
  "app_version": 1,
  "build_date": "2026-06-18",
  "api_enabled": true,
  "accepting": true,
  "policy": "auto_single",
  "port": 8765,
  "has_token": false,
  "pending_count": 0,
  "setlist": {
    "entries": [{"book": "讚美詩", "num": "1"}],
    "index": 0,
    "total": 1,
    "current": {"book": "讚美詩", "num": "1"},
    "label": "排程 1/1 · 讚美詩 / 1"
  }
}
```

#### `GET /api/books`

**回應：** 書冊陣列

```json
[
  {"index": 1, "name": "讚美詩", "hymn_count": 400, "bookmark": true}
]
```

#### `GET /api/entries?book={ref}&q={optional}`

| 參數 | 說明 |
|------|------|
| `book` | 書冊序號（1-based）或書名關鍵字 |
| `q` | 可選過濾（檔名/書籤模糊匹配） |

**回應：**

```json
[
  {"kind": "file", "value": "001.docx", "label": "[DOCX] 001.docx"},
  {"kind": "bookmark", "value": "奇異恩典", "label": "奇異恩典 · P.12"}
]
```

#### `POST /api/open`

**Body：**

```json
{"book": "1", "num": "奇異恩典"}
```

| `book` | 書冊序號或書名 |
| `num` | 檔名、書籤標題、詩歌號（可空） |

**回應 `status`：**

| status | HTTP | 說明 |
|--------|------|------|
| `opened` | 200 | 唯一匹配且已開檔（依策略） |
| `ui_populated` | 200 | 多結果或 `ui_only` 策略 → 桌面列表待揀 |
| `queued` | 200 | `confirm` 策略 → 進入待確認佇列 |
| `not_found` | 200 | 無匹配 |
| `not_accepting` | 503 | 桌面暫停接受 |

#### `POST /api/next`

**Body：** `{}`（空 JSON）

排程索引 +1 並開啟；桌面端自動切換至排程模式。

**回應：**

```json
{"status": "opened", "setlist": { "...": "..." }}
```

| status | 說明 |
|--------|------|
| `opened` | 唯一匹配且已開檔（依 `remote_policy`） |
| `ui_populated` | 多結果或 `ui_only` 策略 |
| `queued` | `confirm` 策略 → 待確認佇列 |
| `not_found` | 排程項無法解析 |
| `failed` | 清單空或索引無效 |
| `unsupported` | 桌面未註冊 handler |
| `not_accepting` | 503 |

排程索引會先更新，再依策略處理開檔（同 `/api/open`）。

#### `GET /api/setlist`

**回應：** 與 `health.setlist` 相同結構（`entries`, `index`, `total`, `current`, `label`）。

#### `POST /api/setlist/open`

**Body：**

```json
{"index": 0}
```

| 欄位 | 說明 |
|------|------|
| `index` | 0-based 排程索引 |

**回應：** 同 `/api/next`（含更新後 `setlist`；`status` 依 `remote_policy`）。

#### `POST /api/setlist/add`

**Body：** 同 `/api/open`（`book` + `num`）

**回應 `status`：**

| status | 說明 |
|--------|------|
| `added` | 已加入桌面排程清單 |
| `not_found` | 找不到書冊 |
| `ambiguous` | 書冊匹配多於一本 |
| `invalid` | 書冊與詩歌號皆空 |
| `disabled` | 桌面 API 未啟用 |

不需「接受請求」為 ON；成功時回傳更新後 `setlist`。

#### `GET /api/log?limit=50`

**回應：** 操作記錄陣列（最新在前，上限 200）

```json
[
  {
    "time": "2026-06-18 14:30:00",
    "ts": 1718692200.0,
    "action": "open",
    "detail": "1 / 奇異恩典",
    "client": "192.168.1.10"
  }
]
```

**action 值：** `open`, `next`, `setlist_open`, `setlist_add`, `approve`, `reject`

#### `GET /api/pending`

**回應：**

```json
[
  {
    "id": "a1b2c3d4",
    "book_ref": "1",
    "num_ref": "恩典",
    "summary": "讚美詩 · 奇異恩典 · P.12",
    "match_count": 1
  }
]
```

#### `POST /api/pending/{req_id}/approve` / `reject`

確認或拒絕 Mobile 待辦（`confirm` 策略）。桌面端亦可按 Enter 或點「開啟/忽略」。

### 4.4 Mobile 開檔策略（`remote_policy`）

適用於 **`/api/open`**、**`/api/next`**、**`/api/setlist/open`**（排程會先更新索引，再依策略開檔）。

| 值 | 常數 | 行為 |
|----|------|------|
| `auto_single` | `OPEN_POLICY_AUTO` | 唯一匹配直接開檔 |
| `confirm` | `OPEN_POLICY_CONFIRM` | 一律進待確認佇列 |
| `ui_only` | `OPEN_POLICY_UI` | 一律填入桌面列表，由操作員揀選 |

### 4.5 Mobile UI（`mobile.html`）

| 分頁 | 顯示條件 | 功能 |
|------|----------|------|
| **開啟** | 永遠 | Token、書冊、檔案/書籤輸入、開啟／加入排程按鈕 |
| **排程** | `setlist.total > 0` | 完整清單、點選開啟、上一首/下一首 |

Token 可存 `localStorage`（key: `hymn_remote_api_token`），或 URL `?token=...`。

**開檔結果提示：** `opened` → 已開啟；`queued` / `ui_populated` →「等待操作員操作」；其餘（`not_found`、`failed` 等）→「無法開啟」。

### 4.6 桌面 ↔ API 回調（內部介面）

`RemoteServer.set_handlers(...)` 註冊：

| Handler | 觸發時機 | 實作 |
|---------|----------|------|
| `get_books` | 解析書冊 | `lambda: self.books` |
| `handle_open` | 直接開檔 | `remote_open_payload.emit(payload)` |
| `handle_populate` | 多結果/UI 策略 | `remote_populate_ui.emit(book, num, matches)` |
| `handle_pending` | 待確認 | `remote_pending_request.emit(req)` |
| `handle_setlist_prepare_next` | `/api/next` | `_invoke_on_main` 更新索引 → policy 開檔 |
| `get_setlist_info` | health / setlist | `_get_setlist_info()` |
| `handle_setlist_prepare_open` | `/api/setlist/open` | `_invoke_on_main` 更新索引 → policy 開檔 |
| `handle_setlist_add` | `/api/setlist/add` | `_invoke_on_main(_schedule_add_entry)` |

**PyQt Signal（主執行緒槽）：**

- `remote_open_payload(dict)`
- `remote_populate_ui(str, str, list)`
- `remote_pending_request(object)`
- `_main_invoke(object)` — 通用主執行緒調度

---

## 5. 資料結構

### 5.1 書冊物件 `book`

```python
{
    'name': str,
    'pdf': str | None,
    'docx': str | None,
    'toc': [(lvl, title, page), ...],
    'bookmark': bool,
    'pages': int,
    'hymn_count': int,
    'files': [{'name', 'path', 'ext'}, ...],
}
```

### 5.2 列表項目 payload（`QListWidgetItem` UserRole）

| kind | 用途 |
|------|------|
| `file` | 一般檔案 |
| `bookmark` | PDF 書籤 |
| `content_pdf` | 關鍵字命中 PDF |
| `content_word` | 關鍵字命中 Word |
| `schedule` | 排程項（含 `index`, `book`, `num`） |

### 5.3 開檔 payload 共通欄位

| kind | 主要欄位 |
|------|----------|
| `file` | `path`, `book`, `name` |
| `bookmark` | `pdf_path`, `page`, `title`, `book`, `toc_level` |
| `content_pdf` / `content_word` | `path`, `page`, `snippet` |

解析邏輯見 `hymn_remote/resolver.py`：`resolve_books` → `resolve_targets` → `resolve_open_request`。

---

## 6. 設定（`settings.json`）

### 6.1 完整鍵值

| 鍵 | 類型 | 預設 | 說明 |
|----|------|------|------|
| `version` | int | 2 | 設定 schema 版本 |
| `hymn_folder` | str | `""` → 預設資料夾 | 詩歌根目錄 |
| `theme` | str | `dark` | `dark` / `light` |
| `font_size` | int | 13 | 10–18 |
| `search_mode` | str | `book` | `book` / `global` / `keyword` / `schedule` |
| `open_overlay` | bool | true | 書名浮層 |
| `overlay_duration` | int | 10 | 1–10 或 -1（一直顯示） |
| `overlay_mode` | str | `corner` | `corner` / `center_then_corner` |
| `display_duplicate` | bool | false | 開檔後同步畫面 |
| `keyword_instant` | bool | true | 關鍵字即時搜索 |
| `book_auto_focus_hymn` | bool | true | 唯一書冊匹配後聚焦詩歌輸入 |
| `click_to_open` | bool | false | 單擊列表開檔 |
| `remote_api_enabled` | bool | false | 啟用 HTTP API |
| `remote_accept` | bool | false | 接受 Mobile 請求 |
| `remote_port` | int | 8765 | API 埠 |
| `remote_policy` | str | `auto_single` | 開檔策略 |
| `remote_token` | str | `""` | API Token（空=不驗證） |
| `operator_mode` | bool | false | **已廢棄**，匯入相容保留 |
| `pinned_books` | list | [] | 釘選書名 |
| `recent_hymns` | list | [] | 最近詩歌 |
| `setlist` | list | [] | 排程 `[{book,num}]` |
| `setlist_index` | int | -1 | 目前排程索引 |
| `session_history` | list | [] | 開檔歷史 |
| `session_cursor` | int | 0 | 歷史游標 |
| `show_preview` | bool | true | 檔案預覽 |
| `auto_rescan` | bool | true | 資料夾自動重掃 |
| `startup_tray` | bool | false | 啟動最小化 Tray |
| `minimize_to_tray` | bool | false | 關閉→Tray |
| `last_opened` | str | `""` | 上次開啟摘要 |

### 6.2 匯出鍵（`EXPORT_KEYS`）

匯出/匯入 JSON **不包含** `hymn_folder`、`remote_api_enabled`、`remote_accept`（避免覆蓋本機 API 狀態）。

### 6.3 排程文字格式（`parse_setlist`）

支援每行：`書冊 / 詩歌號`、`書冊|詩歌號`、Tab 分隔、或純書冊名；`#` 開頭為註解。

---

## 7. 常數與路徑

| 常數 | 值 | 說明 |
|------|-----|------|
| `DEFAULT_PORT` | 8765 | Mobile API |
| `DEFAULT_OVERLAY_SECS` | 10 | Overlay 預設秒數 |
| `OVERLAY_DURATION_ALWAYS` | -1 | 一直顯示 |
| `CENTER_OVERLAY_SECS` | 5 | 置中階段秒數 |
| `INDEX_VERSION` | 1 | 全文索引版本 |
| `MAX_SESSION_HISTORY` | 50 | Session 歷史上限 |
| `settings.json` | `{app_dir}/` | UI 設定 |
| `content_index.json.gz` | `{hymn_folder}/` | 全文索引快取 |
| `remote_api.log` | `{app_dir}/` | uvicorn 啟動錯誤 log |

---

## 8. 外部依賴

### 8.1 Python 套件（`requirements.txt`）

```
PyQt6
PyMuPDF
python-docx
pyinstaller
fastapi
uvicorn[standard]
qrcode[pil]
pypinyin
```

### 8.2 系統程式

| 功能 | Windows | 其他平台 |
|------|---------|----------|
| PDF 閱讀器 | Sumatra / Foxit / Adobe / 預設 | 系統預設 |
| Word 開啟 | Microsoft Word | LibreOffice / open |
| 顯示同步 | DisplaySwitch.exe | 不支援 |
| 文件關閉監聽 | PowerShell + 行程標題 | 不隱藏 overlay |
| 防火牆 | `scripts/add_firewall_rule.ps1` | — |

---

## 9. 專案檔案清單

```
hymn_search/
├── hymn_search.py              # 主程式、MainWindow、掃描、索引、Overlay
├── version.py                  # APP_VERSION、BUILD_DATE
├── settings.json               # 執行時設定（exe 旁）
├── assets/                     # app.ico / app.png
├── hymn_features/
│   ├── mixin.py                # Session、排程、Tray、Remote QR
│   ├── session_store.py        # 設定、排程、歷史、匯出
│   ├── fuzzy.py                # 模糊匹配
│   └── preview.py              # 列表預覽
├── hymn_remote/
│   ├── server.py               # uvicorn 生命週期
│   ├── api.py                  # HTTP 路由
│   ├── state.py                # 遙控狀態、策略、待辦
│   ├── resolver.py             # 書冊/詩歌解析
│   ├── operation_log.py        # 操作記錄
│   └── static/mobile.html      # Mobile UI
├── scripts/
│   ├── make_icon.py
│   ├── add_firewall_rule.ps1
│   └── add_firewall_rule.cmd
├── build_exe.ps1               # 打包 + 版本遞增 + BUILD_DATE
├── HymnSearch.spec
├── requirements.txt
├── 神家詩歌集/                  # 預設資料夾（可改名/移動）
└── docs/
    ├── SYSTEM_SPEC.md            # 本文件
    ├── DESIGN.md
    └── USER_GUIDE.md
```

---

## 10. 鍵盤快捷鍵

| 快捷鍵 | 行為 |
|--------|------|
| F1 / F2 / F3 / F4 | 切換書本 / 全局 / 關鍵字 / 排程模式 |
| Ctrl+Enter | 聚焦書冊輸入（非書本模式時會先切換至書本） |
| Enter | 有待確認 Mobile 請求時 → 確認開啟 |
| Ctrl+← / Ctrl+→ | Session 上一首 / 下一首 |
| Alt+← / Alt+→ | 排程上一首 / 下一首 |
| ? | 快捷鍵說明 |
| Filter / Open 按鈕 | 書本模式過濾或開啟唯一結果 |

---

## 11. 限制與已知問題

| 項目 | 說明 |
|------|------|
| `.doc` 全文搜索 | 需 Windows + Word COM（可選 pywin32） |
| PyInstaller exe | 首次啟動較慢；防毒可能誤報 |
| 同步畫面 | 切換後視窗位置可能略為移動 |
| Mobile API | 僅 LAN；無 HTTPS；Token 為明文 Header |
| 遙控記錄 | 僅記憶體，重啟清空 |
| `operator_mode` | 已移除 UI，舊設定鍵保留相容 |
| API 執行緒 | 必須經 Signal / `_invoke_on_main` 操作 Qt，否則 QTimer 警告 |

---

## 12. 修訂紀錄

| 日期 | 變更 |
|------|------|
| 2026-06 | Phase 1 初版 |
| 2026-06-18 | 四種搜尋模式（含排程）、Mobile API 全端點、Session/釘選/Tray、設定 v2、移除操作員模式 UI |
