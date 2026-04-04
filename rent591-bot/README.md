# 591 租屋通知 LINE 機器人

LINE 聊天機器人，定期查詢 591.com.tw 新北市新刊登的租屋物件，主動推播給訂閱使用者。

## 篩選條件

- **地區**：新北市 — 板橋、土城、中和、永和、新店、新莊、三重
- **類型**：獨立套房、整層住家
- **建築**：電梯大樓、公寓
- **租金**：36,000 以下
- **排除**：頂樓加蓋

## LINE 指令

| 指令 | 說明 |
|---|---|
| `/list` | 立即查詢最新 5 筆 |
| `/subscribe` | 訂閱通知 |
| `/unsubscribe` | 取消訂閱 |

加入好友時會自動訂閱。

## 環境變數

| 變數 | 說明 | 必填 |
|---|---|---|
| `LINE_CHANNEL_ACCESS_TOKEN` | LINE Developers Console 取得 | ✅ |
| `LINE_CHANNEL_SECRET` | Webhook 簽名驗證 | ✅ |
| `CHECK_INTERVAL_MINUTES` | 檢查間隔，預設 5 分鐘 | 選填 |
| `BOT_PORT` | Flask port，預設 8080 | 選填 |
| `BOT_DB_PATH` | SQLite 路徑，預設 `/data/bot_state.db` | 選填 |

## 本機開發

```bash
# 安裝依賴
pip install -r requirements.txt

# 設定環境變數
export LINE_CHANNEL_ACCESS_TOKEN=xxx
export LINE_CHANNEL_SECRET=yyy
export BOT_DB_PATH=./bot_state.db

# 啟動
python line_bot.py

# 開發用 HTTPS（另一個終端）
ngrok http 8080
# 將 https://xxx.ngrok.io/webhook 填入 LINE Developers Console
```

## 測試爬蟲（不需 LINE 帳號）

```bash
python -c "
import scraper_591
s = scraper_591._init_session()
listings = scraper_591.fetch_listings(s)
print(f'取得 {len(listings)} 筆')
for l in listings[:3]:
    print(f'  {l[\"section_name\"]} {l[\"title\"]} NT\${l[\"price\"]}')
    print(f'  {l[\"detail_url\"]}')
"
```

## 部署到 Fly.io

```bash
# 1. 安裝 Fly CLI
curl -L https://fly.io/install.sh | sh

# 2. 登入
fly auth login

# 3. 建立 app
fly apps create rent591-line-bot

# 4. 建立持久磁碟（SQLite 用）
fly volumes create bot_data --region nrt --size 1

# 5. 設定密鑰
fly secrets set LINE_CHANNEL_ACCESS_TOKEN=xxx LINE_CHANNEL_SECRET=yyy

# 6. 部署
fly deploy

# 7. 到 LINE Developers Console 設定 Webhook URL
#    https://rent591-line-bot.fly.dev/webhook
```

## 架構

```
line_bot.py       — Flask Webhook + APScheduler + SQLite
scraper_591.py    — 591.com.tw HTTP API 客戶端
Dockerfile        — 容器化部署
fly.toml          — Fly.io 設定
```

## LINE Developer 帳號申請步驟

1. 前往 https://developers.line.biz/ 以 LINE 帳號登入
2. 建立 Provider（名稱自訂）
3. 建立 Messaging API Channel
4. 在 Channel 設定頁取得 **Channel secret**
5. 在 Messaging API 頁籤發行 **Channel access token (long-lived)**
6. 在 Messaging API 頁籤設定 **Webhook URL**：`https://your-app.fly.dev/webhook`
7. 開啟 **Use webhook**
8. 關閉 **Auto-reply messages**（在 LINE Official Account Manager）

## Fly.io 帳號申請步驟

1. 前往 https://fly.io/ 點擊 Sign Up
2. 可用 GitHub 帳號登入
3. 免費方案包含：3 個 shared-cpu VM + 3GB 持久磁碟
4. 不需信用卡即可開始（部分功能可能需要）
