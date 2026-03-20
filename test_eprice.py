"""
ePrice 自動化測試腳本
批次查詢產品並產出 HTML + CSV 測試報告。
"""
import argparse
import csv
import getpass
import json
import os
import re
import sys
import time
from datetime import datetime

# ── 從 web_app 匯入核心函式 ──
from web_app import fetch_base_price, _parse_numeric


# ═══════════════════════════════════════════
# B. 設定檔載入
# ═══════════════════════════════════════════

def load_config(path: str) -> dict:
    """載入 test_products.json 並做基本檢查。"""
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    if "products" not in cfg or not isinstance(cfg["products"], list):
        raise ValueError("設定檔缺少 products 陣列")
    for i, p in enumerate(cfg["products"]):
        if "name" not in p:
            raise ValueError(f"products[{i}] 缺少 name 欄位")
    return cfg


# ═══════════════════════════════════════════
# D. 測試執行（分批呼叫 fetch_base_price）
# ═══════════════════════════════════════════

def run_batches(username: str, password: str, cfg: dict) -> list:
    """
    將產品清單分批，每批呼叫一次 fetch_base_price。
    回傳 list of dict，每個 dict 含原始產品設定 + 查詢結果。
    """
    products = cfg["products"]
    batch_size = cfg.get("batch_size", 5)
    period_start = cfg.get("period_start", "202501")
    total = len(products)

    results = []

    for batch_start in range(0, total, batch_size):
        batch = products[batch_start: batch_start + batch_size]
        batch_end = min(batch_start + batch_size, total)
        batch_label = f"[批次 {batch_start + 1}-{batch_end}/{total}]"

        names = [p["name"] for p in batch]
        print(f"\n{batch_label} 查詢: {', '.join(names)}")

        t0 = time.time()
        try:
            ok, warning, product_list = fetch_base_price(
                username, password, names,
                period_start=period_start,
                on_progress=lambda step, label: print(f"  [{step}] {label}"),
            )
            elapsed = time.time() - t0
            print(f"{batch_label} 完成 ({elapsed:.1f}s), ok={ok}, 結果數={len(product_list)}")
            if warning:
                print(f"  ⚠ {warning}")
        except Exception as ex:
            elapsed = time.time() - t0
            ok, warning, product_list = False, str(ex), []
            print(f"{batch_label} 例外 ({elapsed:.1f}s): {ex}")

        # 將每個產品的結果配對回設定
        for p_cfg in batch:
            matched = [r for r in product_list if r.get("query_name") == p_cfg["name"]]
            results.append({
                "config": p_cfg,
                "ok": ok,
                "warning": warning,
                "matched_results": matched,
                "elapsed": elapsed / len(batch),  # 平均每產品耗時
            })

    return results


# ═══════════════════════════════════════════
# E. 驗證引擎
# ═══════════════════════════════════════════

VALID_CURRENCIES = {"USD", "NTD", "RMB", "EUR", "JPY"}

RATE_MAP = {}  # 執行時從 cfg 填入


def _to_ntd(value: float, currency: str) -> float:
    """簡易匯率轉換（與 web_app 邏輯一致）。"""
    cur = currency.upper().strip()
    if cur == "NTD":
        return value
    rate = RATE_MAP.get(cur)
    if rate:
        return value * rate
    return value  # 無匯率資訊時回傳原值


def validate_product(result: dict) -> list:
    """
    對單一產品結果執行所有驗證。
    回傳 list of (check_id, passed: bool, detail: str)。
    """
    cfg = result["config"]
    matched = result["matched_results"]
    expect_found = cfg.get("expect_found", True)
    is_waigoupin = cfg.get("is_waigoupin", False)
    selling_price = cfg.get("selling_price")
    sell_currency = cfg.get("sell_currency", "USD")

    checks = []

    # FOUND
    if expect_found:
        found = len(matched) > 0
        checks.append(("FOUND", found, f"找到 {len(matched)} 筆" if found else "未找到"))
    else:
        found = len(matched) == 0
        checks.append(("FOUND", found, "正確未找到" if found else f"預期無結果但找到 {len(matched)} 筆"))

    # 後續檢查只在「預期找到且確實找到」時執行
    if not expect_found or not matched:
        return checks

    # 取第一筆結果做驗證
    r = matched[0]

    # FIELDS
    has_fields = bool(r.get("product_number_name_spec")) and bool(r.get("currency")) and bool(r.get("base_price"))
    checks.append(("FIELDS", has_fields, "必要欄位完整" if has_fields else "缺少欄位"))

    # BP_NUMERIC
    bp_val = _parse_numeric(r.get("base_price", ""))
    bp_ok = bp_val is not None and bp_val > 0
    checks.append(("BP_NUMERIC", bp_ok, f"Base Price={bp_val}" if bp_ok else f"無法解析: {r.get('base_price')}"))

    # CURRENCY
    cur = (r.get("currency") or "").strip().upper()
    cur_ok = cur in VALID_CURRENCIES
    checks.append(("CURRENCY", cur_ok, f"幣別={cur}" if cur_ok else f"不明幣別: {cur}"))

    # COST_H（外購品跳過）
    cost_entries = r.get("cost_highest_entries", [])
    if not is_waigoupin:
        has_cost = len(cost_entries) > 0 and any(e.get("cost", 0) > 0 for e in cost_entries)
        checks.append(("COST_H", has_cost, f"{len(cost_entries)} 筆 cost" if has_cost else "無 Cost(Highest)"))
    else:
        checks.append(("COST_H", True, "外購品跳過"))

    # WAIGOU
    if is_waigoupin:
        cl = r.get("cost_lowest", "")
        cl_val = _parse_numeric(cl)
        waigou_ok = cl_val is not None and cl_val > 0
        checks.append(("WAIGOU", waigou_ok, f"Cost(Lowest)={cl_val}" if waigou_ok else f"外購品缺 Cost(Lowest): {cl}"))
    else:
        checks.append(("WAIGOU", True, "非外購品"))

    # E2E_SANE
    if selling_price and bp_val:
        # 決定用哪個 cost 計算 E2E%
        if is_waigoupin:
            cost_ntd = _parse_numeric(r.get("cost_lowest", ""))
        else:
            if cost_entries:
                cost_ntd = _to_ntd(cost_entries[0]["cost"], cost_entries[0].get("currency", "NTD"))
            else:
                cost_ntd = None

        selling_ntd = _to_ntd(selling_price, sell_currency)

        if cost_ntd and selling_ntd:
            e2e = (selling_ntd - cost_ntd) / selling_ntd * 100
            sane = -50 <= e2e <= 95
            checks.append(("E2E_SANE", sane, f"E2E%={e2e:.1f}%"))
        else:
            checks.append(("E2E_SANE", False, "無法計算 E2E%（缺 cost 或 selling price）"))
    else:
        checks.append(("E2E_SANE", True, "無 selling_price，跳過"))

    return checks


# ═══════════════════════════════════════════
# F. 報告產生
# ═══════════════════════════════════════════

def _ensure_report_dir():
    """確保 test_reports/ 目錄存在。"""
    os.makedirs("test_reports", exist_ok=True)


def generate_html_report(all_results: list, username: str, total_elapsed: float, timestamp: str) -> str:
    """產生 HTML 報告，回傳檔案路徑。"""
    _ensure_report_dir()
    filename = f"test_reports/test_report_{timestamp}.html"

    total = len(all_results)
    passed = sum(1 for r in all_results if all(c[1] for c in r["checks"]))
    failed = total - passed

    rows_html = []
    for r in all_results:
        cfg = r["config"]
        checks = r["checks"]
        all_pass = all(c[1] for c in checks)
        status_color = "#27ae60" if all_pass else "#e74c3c"
        status_text = "PASS" if all_pass else "FAIL"

        # 檢查結果 HTML
        checks_html_parts = []
        for cid, ok, detail in checks:
            color = "#27ae60" if ok else "#e74c3c"
            icon = "&#10004;" if ok else "&#10008;"
            checks_html_parts.append(
                f'<span style="color:{color}" title="{detail}">{icon} {cid}</span>'
            )
        checks_html = " &nbsp; ".join(checks_html_parts)

        # 提取結果欄位
        matched = r["matched_results"]
        bp = matched[0].get("base_price", "") if matched else ""
        cur = matched[0].get("currency", "") if matched else ""
        cost_h = ""
        cost_l = ""
        if matched:
            entries = matched[0].get("cost_highest_entries", [])
            if entries:
                cost_h = f"{entries[0]['cost']:.2f} {entries[0].get('currency', '')}"
            cost_l = matched[0].get("cost_lowest", "")

        # 計算 E2E%
        e2e_str = ""
        for cid, ok, detail in checks:
            if cid == "E2E_SANE" and "E2E%=" in detail:
                e2e_str = detail.replace("E2E%=", "")

        rows_html.append(f"""
        <tr>
            <td>{cfg['name']}</td>
            <td>{cfg.get('description', '')}</td>
            <td style="color:{status_color};font-weight:bold">{status_text}</td>
            <td style="font-size:0.85em">{checks_html}</td>
            <td>{bp}</td>
            <td>{cur}</td>
            <td>{cost_h}</td>
            <td>{cost_l}</td>
            <td>{e2e_str}</td>
            <td>{r['elapsed']:.1f}s</td>
        </tr>""")

    html = f"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<title>ePrice 測試報告 - {timestamp}</title>
<style>
  body {{ font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #f5f5f5; }}
  h1 {{ color: #2c3e50; }}
  .summary {{ display: flex; gap: 20px; margin: 15px 0; }}
  .summary .card {{
    padding: 12px 20px; border-radius: 8px; color: #fff; font-size: 1.1em;
  }}
  .card.total {{ background: #3498db; }}
  .card.pass {{ background: #27ae60; }}
  .card.fail {{ background: #e74c3c; }}
  .card.time {{ background: #8e44ad; }}
  table {{
    border-collapse: collapse; width: 100%; background: #fff;
    box-shadow: 0 1px 3px rgba(0,0,0,0.1);
  }}
  th {{ background: #34495e; color: #fff; padding: 10px 8px; text-align: left; font-size: 0.9em; }}
  td {{ padding: 8px; border-bottom: 1px solid #eee; font-size: 0.9em; }}
  tr:hover {{ background: #f0f6ff; }}
  .meta {{ color: #7f8c8d; margin-bottom: 10px; }}
</style>
</head>
<body>
<h1>ePrice 自動化測試報告</h1>
<div class="meta">
  測試時間: {timestamp} &nbsp;|&nbsp; 帳號: {username}
</div>
<div class="summary">
  <div class="card total">總數: {total}</div>
  <div class="card pass">Pass: {passed}</div>
  <div class="card fail">Fail: {failed}</div>
  <div class="card time">耗時: {total_elapsed:.1f}s</div>
</div>
<table>
<thead>
<tr>
  <th>產品名稱</th><th>描述</th><th>狀態</th><th>檢查項目</th>
  <th>Base Price</th><th>Currency</th><th>Cost(Highest)</th><th>Cost(Lowest)</th>
  <th>E2E%</th><th>耗時</th>
</tr>
</thead>
<tbody>
{''.join(rows_html)}
</tbody>
</table>
</body>
</html>"""

    with open(filename, "w", encoding="utf-8") as f:
        f.write(html)
    return filename


def generate_csv_report(all_results: list, timestamp: str) -> str:
    """產生 CSV 報告，回傳檔案路徑。"""
    _ensure_report_dir()
    filename = f"test_reports/test_report_{timestamp}.csv"

    with open(filename, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "產品名稱", "描述", "狀態", "Base Price", "Currency",
            "Cost(Highest)", "Cost(Lowest)", "E2E%", "耗時(s)",
            "FOUND", "FIELDS", "BP_NUMERIC", "CURRENCY", "COST_H", "WAIGOU", "E2E_SANE",
        ])

        for r in all_results:
            cfg = r["config"]
            checks = r["checks"]
            all_pass = all(c[1] for c in checks)
            matched = r["matched_results"]

            bp = matched[0].get("base_price", "") if matched else ""
            cur = matched[0].get("currency", "") if matched else ""
            cost_h = ""
            cost_l = ""
            if matched:
                entries = matched[0].get("cost_highest_entries", [])
                if entries:
                    cost_h = f"{entries[0]['cost']:.2f} {entries[0].get('currency', '')}"
                cost_l = matched[0].get("cost_lowest", "")

            e2e_str = ""
            for cid, ok, detail in checks:
                if cid == "E2E_SANE" and "E2E%=" in detail:
                    e2e_str = detail.replace("E2E%=", "")

            # 每項檢查的 pass/fail
            check_map = {cid: ("PASS" if ok else f"FAIL: {detail}") for cid, ok, detail in checks}

            writer.writerow([
                cfg["name"], cfg.get("description", ""),
                "PASS" if all_pass else "FAIL",
                bp, cur, cost_h, cost_l, e2e_str,
                f"{r['elapsed']:.1f}",
                check_map.get("FOUND", ""),
                check_map.get("FIELDS", ""),
                check_map.get("BP_NUMERIC", ""),
                check_map.get("CURRENCY", ""),
                check_map.get("COST_H", ""),
                check_map.get("WAIGOU", ""),
                check_map.get("E2E_SANE", ""),
            ])

    return filename


# ═══════════════════════════════════════════
# G. 主程式
# ═══════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="ePrice 自動化測試腳本")
    parser.add_argument("--config", default="test_products.json", help="產品測試清單 JSON 路徑")
    parser.add_argument("--username", help="ePrice 帳號（未指定則互動輸入）")
    parser.add_argument("--password", help="ePrice 密碼（未指定則互動輸入）")
    parser.add_argument("--dry-run", action="store_true", help="僅驗證設定檔，不執行查詢")
    args = parser.parse_args()

    # 載入設定檔
    print(f"載入設定檔: {args.config}")
    try:
        cfg = load_config(args.config)
    except Exception as ex:
        print(f"設定檔錯誤: {ex}")
        sys.exit(1)

    products = cfg["products"]
    print(f"產品數量: {len(products)}")
    print(f"批次大小: {cfg.get('batch_size', 5)}")
    print(f"Period Start: {cfg.get('period_start', '202501')}")
    print(f"匯率 - USD: {cfg.get('usd_rate', '未設定')}, RMB: {cfg.get('rmb_rate', '未設定')}")
    print()

    for i, p in enumerate(products):
        tag = "找到" if p.get("expect_found", True) else "不存在"
        waigou = " [外購品]" if p.get("is_waigoupin", False) else ""
        print(f"  {i+1}. {p['name']} — {p.get('description', '')}{waigou} (預期: {tag})")

    if args.dry_run:
        print("\n[OK] 設定檔格式正確（--dry-run 模式，不執行查詢）")
        sys.exit(0)

    # 帳密輸入
    username = args.username or input("\nePrice 帳號: ").strip()
    password = args.password or getpass.getpass("ePrice 密碼: ")
    if not username or not password:
        print("帳號或密碼不可為空")
        sys.exit(1)

    # 設定匯率
    global RATE_MAP
    RATE_MAP = {
        "USD": cfg.get("usd_rate", 32),
        "RMB": cfg.get("rmb_rate", 4.4),
    }

    # 執行測試
    print(f"\n{'='*60}")
    print("開始執行測試")
    print(f"{'='*60}")

    t_start = time.time()
    results = run_batches(username, password, cfg)
    total_elapsed = time.time() - t_start

    # 執行驗證
    print(f"\n{'='*60}")
    print("驗證結果")
    print(f"{'='*60}")

    for r in results:
        checks = validate_product(r)
        r["checks"] = checks

        name = r["config"]["name"]
        all_pass = all(c[1] for c in checks)
        status = "PASS" if all_pass else "FAIL"
        print(f"\n  {name}: {status}")
        for cid, ok, detail in checks:
            icon = "  [v]" if ok else "  [x]"
            print(f"    {icon} {cid}: {detail}")

    # 產生報告
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = generate_html_report(results, username, total_elapsed, timestamp)
    csv_path = generate_csv_report(results, timestamp)

    passed = sum(1 for r in results if all(c[1] for c in r["checks"]))
    failed = len(results) - passed

    print(f"\n{'='*60}")
    print(f"測試完成: {passed} Pass / {failed} Fail / 總計 {len(results)} | 耗時 {total_elapsed:.1f}s")
    print(f"HTML 報告: {html_path}")
    print(f"CSV  報告: {csv_path}")
    print(f"{'='*60}")

    sys.exit(1 if failed > 0 else 0)


if __name__ == "__main__":
    main()
