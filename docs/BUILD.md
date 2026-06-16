# 詩歌冊搜索系統 — 封裝（打包）指南

本指南說明如何將 Python 原始碼封裝為 Windows 可執行檔（`.exe`），無需使用者安裝 Python。

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

---

## 3. 打包腳本說明

### 3.1 `build_exe.ps1`

```powershell
pyinstaller --noconfirm --onefile --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --hidden-import fitz `
    hymn_search.py
```

| 參數 | 說明 |
|------|------|
| `--onefile` | 單一 exe（啟動時解壓至暫存目錄） |
| `--windowed` | 不顯示黑色命令列視窗 |
| `--icon` | exe 檔案圖示 |
| `--add-data "assets;assets"` | 打包圖示等資源（Windows 用 `;` 分隔） |
| `--hidden-import fitz` | 確保 PyMuPDF 被包含 |

### 3.2 `HymnSearch.spec`

PyInstaller 產生的規格檔，可重複使用：

```powershell
pyinstaller HymnSearch.spec
```

修改 spec 後以此命令重建，無需重新輸入參數。

---

## 4. 應用程式圖示

### 4.1 檔案位置

```
assets/
├── app.png    # 來源圖（256×256 建議）
└── app.ico    # Windows 圖示（exe + 視窗）
```

### 4.2 更換圖示

1. 替換 `assets/app.png`
2. 執行：

```powershell
python scripts\make_icon.py
```

3. 重新打包

程式透過 `app_resource_path('assets', 'app.ico')` 載入圖示，開發與打包環境皆適用。

---

## 5. 目錄結構（打包相關）

```
hymn_search/
├── hymn_search.py       # 主程式入口
├── assets/              # 圖示資源（會打入 exe）
├── scripts/
│   └── make_icon.py     # PNG → ICO
├── build_exe.ps1        # 打包腳本
├── HymnSearch.spec      # PyInstaller 規格
├── requirements.txt     # 運行 + 打包依賴
├── build/               # 建置暫存（可刪除）
└── dist/
    └── HymnSearch.exe   # ★ 最終產物
```

---

## 6. 分發建議

### 6.1 最小分發包

只需提供：

```
HymnSearch.exe
```

詩歌資料仍讀取使用者桌面的 `神家詩歌集new`（或修改程式內 `HYMN_FOLDER` 後重新打包）。

### 6.2 可選一併提供

- 捷徑（指向 exe）
- 簡短使用說明（`docs/USER_GUIDE.md` 摘錄）

### 6.3 注意事項

| 問題 | 說明 |
|------|------|
| 首次啟動慢 | onefile 需解壓，約 5–15 秒 |
| 體積大 | 含 PyQt6 + PyMuPDF，約 50–80 MB |
| 防毒誤報 | 可簽署代碼或加入白名單 |
| 路徑含中文 | 一般可正常運行；建議避免過深路徑 |

---

## 7. 進階選項

### 7.1 資料夾模式（啟動較快）

若覺得 onefile 太慢，可改為 onedir：

```powershell
pyinstaller --noconfirm --windowed `
    --name "HymnSearch" `
    --icon "assets\app.ico" `
    --add-data "assets;assets" `
    --hidden-import fitz `
    hymn_search.py
```

（省略 `--onefile`）產物在 `dist\HymnSearch\` 資料夾，需整包分發。

### 7.2 修改 exe 名稱為中文

```powershell
pyinstaller ... --name "詩歌冊搜索" ...
```

部分舊系統對中文檔名支援較差，建議維持 `HymnSearch.exe`，捷徑使用中文名稱。

### 7.3 清理建置暫存

```powershell
Remove-Item -Recurse -Force build, dist
Remove-Item HymnSearch.spec   # 若需完全重來
```

---

## 8. 疑難排解

### `ModuleNotFoundError: fitz`

```powershell
pip install PyMuPDF
pyinstaller ... --hidden-import fitz
```

### 執行 exe 無圖示

確認 `--add-data "assets;assets"` 已包含，且已執行 `make_icon.py`。

### exe 閃退

暫時去掉 `--windowed` 重建，從命令列執行 exe 查看錯誤訊息：

```powershell
pyinstaller --onefile --icon assets\app.ico hymn_search.py
.\dist\hymn_search.exe
```

---

## 9. 相關文檔

- [SYSTEM_SPEC.md](SYSTEM_SPEC.md) — 系統規格
- [DESIGN.md](DESIGN.md) — 架構設計
- [USER_GUIDE.md](USER_GUIDE.md) — 使用手冊
