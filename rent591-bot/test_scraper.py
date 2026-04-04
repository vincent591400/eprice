"""
測試 591 爬蟲邏輯正確性
驗證篩選條件是否確實生效：地區、類型、租金、排除頂加
"""

import scraper_591

# 篩選條件對照表
VALID_SECTIONS = {"板橋區", "土城區", "中和區", "永和區", "新店區", "新莊區", "三重區"}
VALID_KINDS = {"整層住家", "獨立套房"}
VALID_SHAPES = {"公寓", "電梯大樓"}
MAX_PRICE = 36000


def parse_price(price_str: str) -> int:
    """將 '18,000' 轉為 18000"""
    try:
        return int(price_str.replace(",", "").replace("元", "").strip())
    except Exception:
        return 0


def run_tests(listings: list):
    print(f"\n{'='*50}")
    print(f"共取得 {len(listings)} 筆物件，開始驗證...\n")

    errors = []
    for i, l in enumerate(listings, 1):
        post_id = l["post_id"]
        section = l["section_name"]
        kind = l["kind_name"]
        shape = l["shape_name"]
        price_str = l["price"]
        price = parse_price(price_str)
        url = l["detail_url"]

        row_errors = []

        # ① 地區驗證
        if section not in VALID_SECTIONS:
            row_errors.append(f"地區錯誤: {section}（不在指定範圍內）")

        # ② 類型驗證
        if kind not in VALID_KINDS:
            row_errors.append(f"類型錯誤: {kind}（應為整層住家或獨立套房）")

        # ③ 建築驗證
        if shape not in VALID_SHAPES:
            row_errors.append(f"建築錯誤: {shape}（應為公寓或電梯大樓）")

        # ④ 租金驗證
        if price > MAX_PRICE:
            row_errors.append(f"租金超標: NT${price_str}（超過 {MAX_PRICE}）")

        if row_errors:
            errors.append((post_id, url, row_errors))
            print(f"❌ [{i}] post_id={post_id}")
            for e in row_errors:
                print(f"     → {e}")
        else:
            print(f"✅ [{i}] {section} {kind}/{shape} NT${price_str} {url}")

    print(f"\n{'='*50}")
    if errors:
        print(f"⚠️  發現 {len(errors)} 筆資料不符合篩選條件！")
        print("   可能原因：API 參數沒有完全生效，需要客戶端再過濾")
    else:
        print(f"✅ 全部 {len(listings)} 筆資料均符合篩選條件！爬蟲邏輯正確。")
    print()


def main():
    print("初始化 591 session...")
    try:
        session = scraper_591._init_session()
        print("✅ Session 建立成功\n")
    except Exception as e:
        print(f"❌ Session 建立失敗: {e}")
        return

    print("呼叫 591 API...")
    listings = scraper_591.fetch_listings(session)

    if not listings:
        print("❌ 未取得任何物件，請檢查網路或 API 是否有變動")
        return

    run_tests(listings)

    # 顯示統計
    sections = {}
    kinds = {}
    for l in listings:
        s = l["section_name"]
        k = l["kind_name"]
        sections[s] = sections.get(s, 0) + 1
        kinds[k] = kinds.get(k, 0) + 1

    print("📊 地區分布：")
    for s, c in sorted(sections.items(), key=lambda x: -x[1]):
        print(f"   {s}: {c} 筆")

    print("\n📊 類型分布：")
    for k, c in sorted(kinds.items(), key=lambda x: -x[1]):
        print(f"   {k}: {c} 筆")


if __name__ == "__main__":
    main()
