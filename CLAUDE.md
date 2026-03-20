# ePrice Query Tool

## 專案概述

內部工具，用 Selenium 自動登入 ePricebook 查詢產品 Base Price，再到 anydoor SAP Report 查詢 Cost(Highest)，計算 E2E%。

## 架構

- **`web_app.py`** — Flask 後端（網頁版），提供 SSE 串流 API (`/api/query_stream`)
- **`eprice_gui.py`** — Tkinter GUI 版（獨立桌面應用）
- **`templates/index.html`** — 前端單頁，含查詢表單、進度條、結果表格
- 兩個版本的核心邏輯相同但獨立維護，修改時**兩邊都要同步更新**

## 重要業務邏輯

- **產品名稱匹配**：使用 `_cell_matches_product()` 處理多行欄位（含 spec 描述的產品如 ND-6150）
- **外購品判斷**：產品編號以 `92` 開頭為外購品，Cost(Highest) 顯示「外購品直接參考Cost (Lowest)」，E2E% 改用 Cost(Lowest) 計算
- **Cost(Lowest)**：值為 NTD，不可再做匯率轉換
- **E2E% 計算**：`(Selling Price NTD - Cost NTD) / Selling Price NTD * 100`，所有幣別先透過 `toNTD()` / `_to_ntd()` 換算
- **anydoor 資料收集**：ACT Cost、ACT-Mat、ACT-MOH 三欄都必須 > 0 才收集該列

## 技術注意事項

- Chrome headless 模式需加 `--remote-debugging-port=0` 避免 Windows 上 crash
- `driver.quit()` 用 15 秒逾時的 Thread 包裝，避免 Chrome 關閉卡住
- ePrice 表格 cell 的 `textContent` 可能不含 `\n`（HTML 用 `<br>` 換行），`split("\n")[0]` 取 PN 時要注意
- anydoor 診斷訊息（逾時、表頭缺欄位、無有效資料）會透過 warning 顯示在前端

## 開發指引

- 語言：Python 3.8+，前端原生 JS（無框架）
- 註解與 UI 文字使用繁體中文
- 不要引入額外前端框架或打包工具
- 修改查詢邏輯時，`web_app.py` 和 `eprice_gui.py` 都要改
- 敏感資訊（帳號密碼）不可寫入程式碼或 commit

## 常用指令

```bash
# 啟動網頁版
python web_app.py

# 啟動 GUI 版
python eprice_gui.py

# 安裝依賴
pip install -r requirements.txt
```
