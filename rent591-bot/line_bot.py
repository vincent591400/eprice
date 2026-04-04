"""
LINE 591 租屋通知機器人
主程式：Flask Webhook + APScheduler 定期推播 + SQLite 狀態管理
"""

import os
import sqlite3
import logging
import threading
from datetime import datetime, timedelta

from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi,
    ReplyMessageRequest, PushMessageRequest, TextMessage,
)
from linebot.v3.webhooks import (
    MessageEvent, TextMessageContent,
    FollowEvent, UnfollowEvent,
)
from linebot.v3.exceptions import InvalidSignatureError
from apscheduler.schedulers.background import BackgroundScheduler

import scraper_591

# ── 環境變數（密鑰不可有預設值） ──────────────────────────
LINE_CHANNEL_ACCESS_TOKEN = os.environ["LINE_CHANNEL_ACCESS_TOKEN"]
LINE_CHANNEL_SECRET = os.environ["LINE_CHANNEL_SECRET"]
CHECK_INTERVAL_MINUTES = int(os.getenv("CHECK_INTERVAL_MINUTES", "5"))
BOT_PORT = int(os.getenv("BOT_PORT", "8080"))
DB_PATH = os.getenv("BOT_DB_PATH", "/data/bot_state.db")

# ── 初始化 ────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
)
logger = logging.getLogger(__name__)

app = Flask(__name__)
config = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)
db_lock = threading.Lock()

# ── SQLite 資料庫 ─────────────────────────────────────

def init_db():
    """建立資料表（若不存在）。"""
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.execute("""
                CREATE TABLE IF NOT EXISTS seen_listings (
                    post_id TEXT PRIMARY KEY,
                    seen_at TEXT NOT NULL
                )
            """)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS subscribers (
                    user_id    TEXT PRIMARY KEY,
                    created_at TEXT NOT NULL
                )
            """)
            conn.commit()
        finally:
            conn.close()


def get_seen_ids() -> set:
    """回傳所有已見過的 post_id 集合。"""
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            rows = conn.execute("SELECT post_id FROM seen_listings").fetchall()
            return {r[0] for r in rows}
        finally:
            conn.close()


def mark_seen(post_ids: list):
    """將新的 post_id 批次插入 seen_listings。"""
    now = datetime.utcnow().isoformat()
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.executemany(
                "INSERT OR IGNORE INTO seen_listings (post_id, seen_at) VALUES (?, ?)",
                [(pid, now) for pid in post_ids],
            )
            conn.commit()
        finally:
            conn.close()


def cleanup_old_seen(days: int = 30):
    """清除超過指定天數的 seen_listings 紀錄。"""
    cutoff = (datetime.utcnow() - timedelta(days=days)).isoformat()
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            conn.execute("DELETE FROM seen_listings WHERE seen_at < ?", (cutoff,))
            conn.commit()
        finally:
            conn.close()


def add_subscriber(user_id: str) -> bool:
    """新增訂閱者，若已存在回傳 False。"""
    now = datetime.utcnow().isoformat()
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            cursor = conn.execute(
                "INSERT OR IGNORE INTO subscribers (user_id, created_at) VALUES (?, ?)",
                (user_id, now),
            )
            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()


def remove_subscriber(user_id: str) -> bool:
    """移除訂閱者，若不存在回傳 False。"""
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            cursor = conn.execute("DELETE FROM subscribers WHERE user_id = ?", (user_id,))
            conn.commit()
            return cursor.rowcount > 0
        finally:
            conn.close()


def get_all_subscribers() -> list:
    """回傳所有訂閱者的 user_id list。"""
    with db_lock:
        conn = sqlite3.connect(DB_PATH)
        try:
            rows = conn.execute("SELECT user_id FROM subscribers").fetchall()
            return [r[0] for r in rows]
        finally:
            conn.close()


# ── LINE 訊息工具 ─────────────────────────────────────

def send_reply(reply_token: str, text: str):
    """回覆訊息。"""
    try:
        with ApiClient(config) as api_client:
            api = MessagingApi(api_client)
            api.reply_message(ReplyMessageRequest(
                reply_token=reply_token,
                messages=[TextMessage(text=text)],
            ))
    except Exception as e:
        logger.error("回覆訊息失敗: %s", e)


def send_push(user_id: str, text: str):
    """主動推播訊息給指定使用者。"""
    try:
        with ApiClient(config) as api_client:
            api = MessagingApi(api_client)
            api.push_message(PushMessageRequest(
                to=user_id,
                messages=[TextMessage(text=text)],
            ))
    except Exception as e:
        logger.error("推播訊息給 %s 失敗: %s", user_id, e)


# ── 格式化 ────────────────────────────────────────────

def format_listing(listing: dict, idx: int, total: int) -> str:
    """將單一物件格式化為 LINE 訊息文字。"""
    section = listing.get("section_name", "")
    kind = listing.get("kind_name", "")
    shape = listing.get("shape_name", "")
    title = listing.get("title", "")
    price = listing.get("price", "")
    layout = listing.get("layout", "")
    area = listing.get("area", "")
    floor_info = listing.get("floor_info", "")
    address = listing.get("address", "")
    url = listing.get("detail_url", "")

    # 格局資訊（layout ｜ area坪 ｜ floor）
    detail_parts = []
    if layout:
        detail_parts.append(layout)
    if area:
        detail_parts.append(f"{area}坪")
    if floor_info:
        detail_parts.append(floor_info)
    detail_line = " ｜ ".join(detail_parts)

    lines = [
        f"🏠 [{idx}/{total}] {section} ─ {kind}／{shape}",
        f"標題：{title}",
        f"租金：NT${price}/月",
    ]
    if detail_line:
        lines.append(f"格局：{detail_line}")
    if address:
        lines.append(f"地址：{address}")
    lines.append(f"👉 {url}")

    return "\n".join(lines)


def format_listings_message(listings: list) -> str:
    """將多筆物件格式化為完整推播訊息。"""
    total = len(listings)
    header = f"【新上架租屋】共 {total} 筆\n"
    body = "\n───────────────\n".join(
        format_listing(l, i + 1, total) for i, l in enumerate(listings)
    )
    return header + "\n" + body


# ── 排程核心 ──────────────────────────────────────────

def check_and_notify():
    """
    排程器定期執行：查詢 591 新物件 → 推播給所有訂閱者。
    """
    logger.info("開始檢查 591 新物件...")

    try:
        session = scraper_591._init_session()
    except Exception as e:
        logger.error("初始化 591 session 失敗: %s", e)
        return

    seen_ids = get_seen_ids()
    new_listings = scraper_591.get_new_listings(session, seen_ids)

    if not new_listings:
        logger.info("無新物件")
        return

    subscribers = get_all_subscribers()
    if not subscribers:
        logger.info("有 %d 筆新物件但無訂閱者，僅更新 seen_listings", len(new_listings))
        mark_seen([l["post_id"] for l in new_listings])
        return

    # 每批最多 5 筆，避免訊息過長
    batch_size = 5
    for i in range(0, len(new_listings), batch_size):
        batch = new_listings[i:i + batch_size]
        message = format_listings_message(batch)

        for user_id in subscribers:
            send_push(user_id, message)

    # 更新已見過的 post_id
    mark_seen([l["post_id"] for l in new_listings])

    # 清理舊紀錄
    cleanup_old_seen(days=30)

    logger.info("已推播 %d 筆新物件給 %d 位訂閱者", len(new_listings), len(subscribers))


# ── Flask 路由 ────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    """Health check endpoint for Fly.io."""
    return "OK", 200


@app.route("/webhook", methods=["POST"])
def webhook():
    """LINE Webhook 接收端點。"""
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"


# ── LINE 事件 Handlers ────────────────────────────────

WELCOME_MSG = (
    "歡迎使用 591 租屋通知機器人！🏠\n\n"
    "已自動為您訂閱新北市租屋通知。\n"
    "篩選條件：\n"
    "📍 板橋、土城、中和、永和、新店、新莊、三重\n"
    "🏢 獨立套房 / 整層住家\n"
    "🏗️ 電梯大樓 / 公寓\n"
    "💰 租金 36,000 以下\n"
    "🚫 排除頂樓加蓋\n\n"
    "指令：\n"
    "/list — 立即查詢最新 5 筆\n"
    "/subscribe — 訂閱通知\n"
    "/unsubscribe — 取消訂閱"
)

HELP_MSG = (
    "591 租屋通知機器人 指令：\n\n"
    "/list — 立即查詢最新 5 筆\n"
    "/subscribe — 訂閱通知\n"
    "/unsubscribe — 取消訂閱\n\n"
    "機器人會自動每 {interval} 分鐘檢查新物件並推播。"
).format(interval=CHECK_INTERVAL_MINUTES)


@handler.add(FollowEvent)
def on_follow(event):
    """使用者加入好友時自動訂閱。"""
    user_id = event.source.user_id
    add_subscriber(user_id)
    logger.info("新訂閱者: %s", user_id)
    send_reply(event.reply_token, WELCOME_MSG)


@handler.add(UnfollowEvent)
def on_unfollow(event):
    """使用者封鎖時自動取消訂閱。"""
    user_id = event.source.user_id
    remove_subscriber(user_id)
    logger.info("取消訂閱: %s", user_id)


@handler.add(MessageEvent, message=TextMessageContent)
def on_message(event):
    """處理使用者文字訊息，路由到對應指令。"""
    text = event.message.text.strip().lower()
    user_id = event.source.user_id

    if text == "/subscribe":
        added = add_subscriber(user_id)
        if added:
            send_reply(event.reply_token, "✅ 已訂閱租屋通知！有新物件時會主動通知您。")
        else:
            send_reply(event.reply_token, "您已經訂閱囉！")

    elif text == "/unsubscribe":
        removed = remove_subscriber(user_id)
        if removed:
            send_reply(event.reply_token, "❌ 已取消訂閱。隨時可用 /subscribe 重新訂閱。")
        else:
            send_reply(event.reply_token, "您目前沒有訂閱。")

    elif text == "/list":
        send_reply(event.reply_token, "🔍 查詢中，請稍候...")
        try:
            session = scraper_591._init_session()
            listings = scraper_591.fetch_listings(session)[:5]
            if listings:
                message = format_listings_message(listings)
                send_push(user_id, message)
            else:
                send_push(user_id, "目前沒有符合條件的物件。")
        except Exception as e:
            logger.error("/list 查詢失敗: %s", e)
            send_push(user_id, "查詢失敗，請稍後再試。")

    else:
        send_reply(event.reply_token, HELP_MSG)


# ── 排程器 ────────────────────────────────────────────

def start_scheduler() -> BackgroundScheduler:
    """建立並啟動背景排程器。"""
    scheduler = BackgroundScheduler()
    scheduler.add_job(
        check_and_notify,
        "interval",
        minutes=CHECK_INTERVAL_MINUTES,
        max_instances=1,
        next_run_time=datetime.now(),  # 啟動後立即執行一次
    )
    scheduler.start()
    logger.info("排程器已啟動，每 %d 分鐘檢查一次", CHECK_INTERVAL_MINUTES)
    return scheduler


# ── 主程式入口 ────────────────────────────────────────

if __name__ == "__main__":
    # 確保 DB 目錄存在
    db_dir = os.path.dirname(DB_PATH)
    if db_dir:
        os.makedirs(db_dir, exist_ok=True)

    init_db()
    scheduler = start_scheduler()

    try:
        app.run(host="0.0.0.0", port=BOT_PORT, debug=False)
    finally:
        scheduler.shutdown()
