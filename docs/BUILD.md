# 詩歌冊搜索系統 — 封裝（打包）指南

本指南說明如何將 Python 原始碼封裝為 Windows 可執行檔（`.exe`），以及如何更新會眾觀看所需的 mapping JSON。

---

## 1. 前置需求

| 項目 | 版本 |
|------|------|
| Windows | 10 / 11（64-bit） |
| Python | 3.10 或以上 |
| 磁碟空間 | 約 500 MB（含虛擬環境與建置暫存） |

---

## 2. 快速打包（推薦）

在 PowerShell 中執行：

```powershell
cd C:\python\hymn_search

# 建立虛擬環境（首次）
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 安裝依賴（含 PyInstaller）
pip install -r requirements.txt

# 一鍵打包
.\build_exe.ps1
```

**輸出：**

```
C:\python\hymn_search\dist\HymnSearch.exe
```

雙擊即可執行；可複製到任意位置，無需附帶 Python。

`build_exe.ps1` 會自動：

- 遞增 `version.py` 的 `APP_VERSION` 並寫入 `BUILD_DATE`
- 若缺少 `gdrive_hymn_index.json`，首次執行時以 `gdown` 從 Google Drive 建立
- 打包 `assets`、`hymn_remote/static`、`hymn_remote/data`（含 mapping JSON）

---

## 3. 更新 Mapping JSON（打包前）

會眾觀看頁依兩個 JSON 對應 iframe URL；**更新後必須重新打包 exe** 才會生效。

| 檔案 | 用途 | 建立指令 |
|------|------|----------|
| `hymn_remote/data/cog_hymn_urls.json` | PDF → 教會網站 `churchofgod.org.hk` | `python scripts\build_hymn_url_map.py` |
| `hymn_remote/data/gdrive_hymn_index.json` | Word → Google Drive preview | `python scripts\build_gdrive_hymn_map.py` |

### 3.1 教會網站 mapping（PDF）

從教會網站 sitemap 抓取詩歌 code → URL：

```powershell
python scripts\build_hymn_url_map.py
python scripts\review_hymn_map.py    # 可選：檢查覆蓋率
```

**需要：** 可連線 `churchofgod.org.hk`

### 3.2 Google Drive mapping（Word）

從公開分享的 Drive 資料夾列出所有檔案（書冊/檔名 → file ID）：

```powershell
pip install gdown
python scripts\build_gdrive_hymn_map.py
python scripts\review_gdrive_map.py   # 可選：測試對應
```

**預設資料夾 ID：** `18uMCSYNDgipvGzHwlhyMcxrBaAcjChSZ`

**換資料夾：**

```powershell
python scripts\build_gdrive_hymn_map.py --folder-id YOUR_FOLDER_ID
```

**前提：**

- Drive 資料夾結構應與本地 `神家詩歌集` 一致（例如 `01 神家詩歌/P001  001 父的名.doc`）
- 分享權限設為 **「知道連結的使用者均可檢視」**（會眾手機才能載入 preview）

### 3.3 完整更新流程

```powershell
cd C:\python\hymn_search
.\.venv\Scripts\Activate.ps1

# 1. 更新 mapping（按需）
python scripts\build_hymn_url_map.py
pip install gdown
python scripts\build_gdrive_hymn_map.py

# 2. 重新打包
.\build_exe.ps1
```

---

## 4. 打包腳本說明

### 4.1 `build_exe.ps1`

主要 PyInstaller 參數：

```powershell
pyinstaller --noconfirm --onefile --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --add-data "hymn_remote\static;hymn_remote\static" `
    --add-data "hymn_remote\data;hymn_remote\data" `
    --hidden-import fitz `
    --hidden-import hymn_remote.viewer_hymn_map `
    --hidden-import hymn_remote.gdrive_hymn_map `
    --hidden-import hymn_remote.web_hymn_map `
    ...（FastAPI / uvicorn / qrcode 等）
    hymn_search.py
```

| 參數 | 說明 |
|------|------|
| `--onefile` | 單一 exe（啟動時解壓至暫存目錄） |
| `--windowed` | 不顯示黑色命令列視窗 |
| `--icon` | exe 檔案圖示 |
| `--add-data "assets;assets"` | 打包圖示等資源（Windows 用 `;` 分隔） |
| `--add-data "hymn_remote\static;…"` | 會眾 `viewer.html`、操作員 `mobile.html` |
| `--add-data "hymn_remote\data;…"` | `cog_hymn_urls.json`、`gdrive_hymn_index.json` |
| `--hidden-import fitz` | 確保 PyMuPDF 被包含 |
| `--hidden-import hymn_remote.*_map` | 會眾 iframe mapping 模組 |

### 4.2 `HymnSearch.spec`

PyInstaller 產生的規格檔，可重複使用：

```powershell
pyinstaller HymnSearch.spec
```

修改 spec 後以此命令重建，無需重新輸入參數。`datas` 已包含 `hymn_remote/data` 與 `hymn_remote/static`。

---

## 5. 應用程式圖示

### 5.1 檔案位置

```
assets/
├── app.png    # 來源圖（256×256 建議）
└── app.ico    # Windows 圖示（exe + 視窗）
```

### 5.2 更換圖示

1. 替換 `assets/app.png`
2. 執行：

```powershell
python scripts\make_icon.py
```

3. 重新打包

程式透過 `app_resource_path('assets', 'app.ico')` 載入圖示，開發與打包環境皆適用。

---

## 6. 目錄結構（打包相關）

```
hymn_search/
├── hymn_search.py              # 主程式入口
├── assets/                     # 圖示資源（會打入 exe）
├── hymn_remote/
│   ├── data/                   # ★ mapping JSON（會打入 exe）
│   │   ├── cog_hymn_urls.json
│   │   └── gdrive_hymn_index.json
│   └── static/                 # ★ viewer.html、mobile.html
├── scripts/
│   ├── make_icon.py
│   ├── build_hymn_url_map.py
│   ├── build_gdrive_hymn_map.py
│   ├── review_hymn_map.py
│   └── review_gdrive_map.py
├── build_exe.ps1               # 打包腳本
├── HymnSearch.spec             # PyInstaller 規格
├── requirements.txt            # 運行 + 打包依賴
├── build/                      # 建置暫存（可刪除）
└── dist/
    └── HymnSearch.exe          # ★ 最終產物
```

---

## 7. 分發建議

### 7.1 最小分發包

只需提供：

```
HymnSearch.exe
```

詩歌資料讀取 exe 旁的 `神家詩歌集`（或 `settings.json` 中設定的 `hymn_folder`）。

### 7.2 可選一併提供

- 捷徑（指向 exe）
- 簡短使用說明（`docs/USER_GUIDE.md` 摘錄）
- `神家詩歌集` 資料夾（若使用者無本地副本）

### 7.3 注意事項

| 問題 | 說明 |
|------|------|
| 首次啟動慢 | onefile 需解壓，約 5–15 秒 |
| 體積大 | 含 PyQt6 + PyMuPDF + mapping JSON，約 50–90 MB |
| 防毒誤報 | 可簽署代碼或加入白名單 |
| 路徑含中文 | 一般可正常運行；建議避免過深路徑 |
| 會眾 Drive | mapping 內嵌 exe；Drive 資料夾須保持公開分享 |

---

## 8. 進階選項

### 8.1 資料夾模式（啟動較快）

若覺得 onefile 太慢，可改為 onedir（省略 `--onefile`）。產物在 `dist\HymnSearch\` 資料夾，需整包分發。

### 8.2 修改 exe 名稱為中文

```powershell
pyinstaller ... --name "詩歌冊搜索" ...
```

部分舊系統對中文檔名支援較差，建議維持 `HymnSearch.exe`，捷徑使用中文名稱。

### 8.3 清理建置暫存

```powershell
Remove-Item -Recurse -Force build, dist
```

---

## 9. 疑難排解

### `ModuleNotFoundError: fitz`

```powershell
pip install PyMuPDF
pyinstaller ... --hidden-import fitz
```

### 執行 exe 無圖示

確認 `--add-data "assets;assets"` 已包含，且已執行 `make_icon.py`。

### 會眾 `/view` 無內容

1. 桌面端是否已「啟用 Mobile API」
2. Word 詩歌：檢查 `gdrive_hymn_index.json` 是否過期；Drive 是否仍可公開存取
3. PDF 詩歌：執行 `build_hymn_url_map.py` 更新網站 mapping
4. 設定 → Mobile → 可暫時開啟「會眾觀看顯示除錯列」查看 mapping 結果

### `build_gdrive_hymn_map.py` 失敗

```powershell
pip install gdown
python scripts\build_gdrive_hymn_map.py
```

確認 Drive 資料夾連結有效且為公開分享。

### exe 閃退

暫時去掉 `--windowed` 重建，從命令列執行 exe 查看錯誤訊息：

```powershell
pyinstaller --onefile --icon assets\app.ico hymn_search.py
.\dist\HymnSearch.exe
```

### Mobile API 啟動失敗

查看 exe 同目錄的 `remote_api.log`（常見原因：Port 被佔用、防火牆）。

---

## 10. 相關文檔

- [SYSTEM_SPEC.md](SYSTEM_SPEC.md) — 系統規格（含會眾觀看、mapping 邏輯）
- [DESIGN.md](DESIGN.md) — 架構設計
- [USER_GUIDE.md](USER_GUIDE.md) — 使用手冊
