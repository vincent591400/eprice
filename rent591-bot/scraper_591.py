"""
591 租屋網 API 客戶端
查詢新北市指定區域的租屋物件（獨立套房/整層住家、電梯/公寓、≤36000、排除頂加）
"""

import re
import time
import logging
import requests

logger = logging.getLogger(__name__)

BASE_URL = "https://rent.591.com.tw"
API_URL = f"{BASE_URL}/home/search/rsList"

# 搜尋條件
SEARCH_PARAMS = {
    "is_new_list": 1,
    "type": 1,
    "region": 3,                        # 新北市
    "section": "26,39,38,37,34,44,43",  # 板橋,土城,中和,永和,新店,新莊,三重
    "kind": "1,2",                      # 1=整層住家, 2=獨立套房
    "shape": "1,2",                     # 1=公寓, 2=電梯大樓
    "rentprice": "0,36000",             # 租金上限 36000
    "not_cover": 1,                     # 排除頂樓加蓋
    "order": "posttime",                # 依刊登時間排序
    "orderType": "desc",                # 最新優先
    "firstRow": 0,
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
}


def _init_session() -> requests.Session:
    """
    建立帶有有效 Cookie 與 CSRF token 的 requests Session。
    先 GET 首頁取得 session cookie，再從 HTML 抽出 CSRF token。
    """
    session = requests.Session()
    session.headers.update(HEADERS)

    # GET 首頁取 cookie
    resp = session.get(BASE_URL, timeout=15)
    resp.raise_for_status()

    # 從 <meta name="csrf-token" content="..."> 抓 token
    match = re.search(r'<meta\s+name="csrf-token"\s+content="([^"]+)"', resp.text)
    if not match:
        # 備用：從 cookie 中找
        token = session.cookies.get("XSRF-TOKEN", "")
        if not token:
            logger.warning("無法取得 CSRF token，API 請求可能失敗")
    else:
        token = match.group(1)

    session.headers.update({
        "X-CSRF-TOKEN": token,
        "X-Requested-With": "XMLHttpRequest",
        "Referer": BASE_URL,
    })

    # 延遲避免被擋
    time.sleep(1)
    return session


def fetch_listings(session: requests.Session) -> list:
    """
    呼叫 591 搜尋 API，回傳租屋物件清單。
    每筆包含: post_id, title, price, section_name, address,
             layout, area, floor_info, kind_name, shape_name, detail_url
    失敗回傳空 list。
    """
    try:
        resp = session.get(API_URL, params=SEARCH_PARAMS, timeout=15)
        resp.raise_for_status()
        data = resp.json()
    except Exception as e:
        logger.error("591 API 請求失敗: %s", e)
        return []

    raw_list = data.get("data", {}).get("data", [])
    if not raw_list:
        logger.info("591 API 回傳 0 筆資料")
        return []

    listings = []
    for item in raw_list:
        post_id = str(item.get("post_id", ""))
        if not post_id:
            continue

        floor = item.get("floor", "")
        allfloor = item.get("allfloor", "")
        floor_info = f"{floor}/{allfloor}樓" if floor and allfloor else ""

        listings.append({
            "post_id": post_id,
            "title": item.get("title", ""),
            "price": item.get("price", ""),
            "section_name": item.get("section_name", ""),
            "address": item.get("address", ""),
            "layout": item.get("layout", ""),
            "area": item.get("area", ""),
            "floor_info": floor_info,
            "kind_name": item.get("kind_name", ""),
            "shape_name": item.get("shape_name", ""),
            "detail_url": f"{BASE_URL}/rent-detail-{post_id}.html",
        })

    logger.info("591 API 取得 %d 筆物件", len(listings))
    return listings


def get_new_listings(session: requests.Session, seen_ids: set) -> list:
    """
    呼叫 fetch_listings，過濾掉 seen_ids 中已存在的 post_id。
    回傳僅包含新物件的 list。
    """
    all_listings = fetch_listings(session)
    new_listings = [l for l in all_listings if l["post_id"] not in seen_ids]
    if new_listings:
        logger.info("發現 %d 筆新物件（共 %d 筆）", len(new_listings), len(all_listings))
    return new_listings
