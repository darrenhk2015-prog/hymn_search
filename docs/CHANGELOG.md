# 更新紀錄（Release Notes）

版本號來自 `version.py` 的 `APP_VERSION`（`build_exe.ps1` 每次打包自動 +1）。  
下列各版內容旁標示**信心來源**，方便日後核對。

| 標記 | 意思 |
|------|------|
| **git** | 有對應 git commit，可 `git show <hash>` 核對 |
| **doc** | 來自 `docs/SYSTEM_SPEC.md` 修訂紀錄或規格書 |
| **ver** | 來自 `version.py` / `BUILD_DATE` |
| **diff** | 由 commit 之間或 commit→工作區 diff 推斷 |
| **推斷** | 無獨立 commit，依 build 序號與改動量估計，可能與實際 build 有出入 |

### Git 對照（僅 4 個 commit）

| Commit | 日期 | 訊息 | 當時 `APP_VERSION` |
|--------|------|------|-------------------|
| `23b3c3a` | 2026-06-16 | Init | （尚無 `version.py`） |
| `02111dd` | 2026-06-19 | Commit Good Version | （尚無 `version.py`） |
| `1fa22a3` | 2026-06-20 | v2 | 2 **ver** |
| `ae17304` | 2026-06-21 | v3 | 4 **ver**（commit 名稱與版號不一致） |

> v3–v6 之間無 git commit；v4–v6 分界為**推斷**。

---

## [未發布]

> 下次執行 `build_exe.ps1` 預期為 **v8**。

---

## v7 · 2026-06-30 · ver + diff

**信心：高** — `version.py` 為 v7、`BUILD_DATE = 2026-06-30` **ver**；含先前未發布之 tunnel 修正與桌面 QR 浮層。

### 新增

- **桌面 QR Code 浮層** **diff**
  - Header 新增 **Show QR Code** 開關；啟用後在桌面置頂顯示會眾頁 QR（與書名浮層同層、不搶焦點）
  - **延伸桌面**：每個螢幕各顯示一份（X/Y 相對各螢幕可用區域左上角）
  - **同步複製畫面**：僅主螢幕顯示，避免重複
  - 設定：**QR 背景透明度**（0–100%，控制卡片背景色 alpha；配合 `#000000` 可做出半透明黑底）
  - 移除獨立「透明黑底」層；邊框/內距/陰影收緊，QR 更大
  - 延伸模組 `DesktopQrOverlay`；設定匯入/匯出含上述鍵

### 修正

- **Cloudflare sub-path 會眾頁 404** **diff**：`/hymn_search`（無尾隨 `/`）→ 308 redirect 至 `/hymn_search/`。

### 變更

- 公開連結（QR、外網會眾、`named_tunnel_url`）統一尾隨 `/` **diff**
- 新增 `tunnel_viewer_url()` **diff**
- 載入設定時自動為 path prefix URL 補 `/` **diff**

---

## v6 · 2026-06-21 · ver + diff（推斷）

**信心：中** — `version.py` 為 v6；功能來自 `ae17304` 之後未 commit 改動，與 v5 分界為推斷。

### 新增 / 變更

- **Sub-path 路由**：API gateway 支援 `/hymn_search` prefix；本機 LAN 仍可用無 prefix 路徑 **diff**
- **`remote_path_prefix`** 設定鍵；從 `named_tunnel_url` 自動解析 **diff**
- 預設公開網址 `https://live.churchofgodtm.com/hymn_search/` **diff**
- **共用 cloudflared 提示**：一部 PC 一個 Windows 服務；與 `/ppt` 等 sub-path 並存 **diff**
- `viewer.html` / `mobile.html`：`apiBasePath()` 適配 sub-path **diff**
- `session_store`、Mobile 設定 UI 小調整 **diff**

---

## v5 · ~2026-06-21 · diff（推斷）

**信心：中低** — 與 v6 同屬 `ae17304` 之後的 build，無獨立 commit。

### 新增

- **`hymn_remote/tunnel.py`**：Cloudflare Tunnel（quick / named / service） **diff**
- 設定「外網（Cloudflare）」、tunnel token、安裝 cloudflared 服務 **diff**
- `tunnel.log` / `cloudflared-tunnel.log` 診斷 **diff**
- QR Code：啟用外網時改指向公開 HTTPS **diff**
- `hymn_remote/local_tls.py`（本機 HTTPS / origin CA，後續簡化為 HTTP origin） **diff**

---

## v4 · 2026-06-21 · git `ae17304` + ver

**信心：高** — commit `ae17304`（訊息為 `v3`，但 `APP_VERSION = 4` **ver**）。

### 新增

- **會眾觀看頁** `viewer.html`（`GET /`、`/view`） **git**
- **混合 iframe mapping** **git** **doc**
  - PDF → `churchofgod.org.hk`（`cog_hymn_urls.json`、`web_hymn_map.py`）
  - Word → Google Drive preview（`gdrive_hymn_index.json`、`gdrive_hymn_map.py`）
- `viewer_hymn_map.py`：桌面開檔 → 會眾 `now_playing` **git**
- Mapping 建置腳本：`build_hymn_url_map.py`、`build_gdrive_hymn_map.py` **git**
- 設定「會眾觀看顯示除錯列」（`viewer_show_debug`） **git** **doc**
- UI：外網會眾 URL 標籤 **git**

---

## v3 · ~2026-06-20 · ver（推斷）

**信心：低** — `APP_VERSION` 由 2 增至 4 之間有一次 build；**無對應 commit**，內容不明。

可能為 v2 後續修正或小改動（推測）。若你記得該次 build 的改動，可補充後移除此條或合併入 v2/v4。

---

## v2 · 2026-06-20 · git `1fa22a3` + ver + doc

**信心：高** — commit `1fa22a3`，`APP_VERSION = 2` **ver**。

### 新增

- **`hymn_remote/`** Mobile HTTP API（FastAPI、uvicorn、`mobile.html`） **git** **doc**
  - 開檔、排程、待確認佇列、操作記錄
  - 開檔策略：`auto` / `confirm` / `ui_only`
- **`hymn_features/`** 增強模組 **git**
  - `mixin.py`：Session、Tray、設定匯入/匯出、QR
  - `session_store.py`：排程、Session 歷史、設定正規化
  - `fuzzy.py`、`preview.py`
- **第四種搜尋模式：排程（F4）** **doc**
- **Session**：開啟歷史、Ctrl+←/→ 導航 **doc**
- **釘選書冊**、系統匣、設定四 Tab **doc**
- **`settings.json`** schema v2、**`version.py`** 版本追蹤開始 **git** **ver**
- Windows 防火牆腳本 `scripts/add_firewall_rule.*` **git**
- `docs/SYSTEM_SPEC.md` 大幅擴充 **git**

### 變更

- 主程式重構：Remote 信號、`EnhancementMixin` 整合 **git**

---

## v1 · 2026-06-16 · git `23b3c3a`

**信心：高** — 首次 commit `Init`，首個 `HymnSearch.exe` **git**。

### 新增（Phase 1 桌面）

- PyQt6 桌面：書本 / 全局 / 關鍵字三種搜尋 **git**
- 掃描 `神家詩歌集`：PDF / Word / ODT；PDF TOC 書籤跳頁 **git**
- 全文搜索 + `content_index.json.gz` 持久化索引 **git**
- 開檔：Sumatra/Foxit/Adobe PDF；Word 投影視圖快捷 **git**
- **書名置頂浮層**（多螢幕、置中→角落動畫） **git**
- 延伸桌面同步畫面（Windows `DisplaySwitch.exe`） **git**
- 深/淺色主題、`settings.json` 基礎設定 **git**
- PyInstaller 打包（`build_exe.ps1`、`docs/BUILD.md`） **git**

### 備註

- `02111dd`（2026-06-19，Commit Good Version）為 v1→v2 過渡 **git**；已出現 Remote 相關程式碼雛形，但 `hymn_remote/` 套件至 v2 才提交，**不視為獨立對外版本**。

---

## 維護建議

1. 每次執行 `build_exe.ps1` **之前**，在本檔 `[未發布]` 或新版本標題下寫好 release note，打包後改為 `vN · BUILD_DATE` 並 **git commit**。
2. 理想格式：`git tag vN` + 更新本檔，避免 v3–v6 再度難以追溯。
3. 詳細規格見 [SYSTEM_SPEC.md](SYSTEM_SPEC.md) §12 修訂紀錄。
