# 監造計畫書材料系統

所有 Python 程式都集中在 `streamlit_app.py`。

## 保留檔案

- `streamlit_app.py`：完整網站、登入、Word 解析、選材及四表輸出。
- 四份正式 DOCX：網站上傳與輸出的原始格式。
- `requirements.txt`：Python 套件。
- `packages.txt`：Streamlit Cloud 的 LibreOffice、Poppler 與中文字型。
- `.streamlit/config.toml`：Streamlit 設定。
- `.streamlit/secrets.toml.example`：密碼設定範例。

## 本機啟動

```powershell
.\.venv\Scripts\streamlit.exe run streamlit_app.py
```

若尚未建立虛擬環境：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\streamlit.exe run streamlit_app.py
```

## 設定密碼

```powershell
.\.venv\Scripts\python.exe streamlit_app.py --hash-password
```

將產生的完整 `$scrypt$...` 字串填入 `.streamlit/secrets.toml`：

```toml
[app]
password_hash = "$scrypt$..."
allow_unsecured_local = false
```

正式密碼與真正的 `.streamlit/secrets.toml` 不可提交到 Git。

## 使用流程

1. 網站一次上傳四份正式 DOCX。
2. 系統讀取材料設備送審管制總表的 69 項材料，排序後以 `1.材料名稱` 至 `69.材料名稱` 顯示選單。
3. 材料超過 8 項時，按「＋ 新增材料」繼續增加材料列。
4. 每項材料填寫契約詳細表項次及契約數量；單位直接填在契約數量內，例如 `66CM2`、`555kg`。
5. 整份文件只填一次預定送審年月，格式為 `YYYY.M`（例如 `2026.8`）；空白代表不填。
6. 按「完成」產生一個 Word。
7. 輸出順序固定為送審管制總表、品質標準抽驗表、品質抽驗紀錄表、檢（試）驗管制總表。

上傳資料、選材內容及產出 Word 只保存在本次 Streamlit 記憶體，不會永久寫入伺服器。

使用者填寫的預定送審年月會套用到全部已選材料，並同時寫入送審管制總表的「預定送審日期」，以及檢（試）驗管制總表的「預定進場日期」。契約數量不會寫入進場數量。
