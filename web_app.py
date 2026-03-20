"""
ePrice 查詢工具 - 網頁版後端
Flask + Selenium：登入 epricebook 並查詢 Product 的 Base Price，再到 anydoor 查 Cost(highest)
"""
import json
import re
import time
from datetime import datetime
from typing import Tuple
from urllib.parse import quote

from flask import Flask, request, render_template, jsonify, Response, stream_with_context

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException

app = Flask(__name__)

LOGIN_URL = "http://epricebook.adlinktech.com/Login.aspx"
PRICE_URL = "http://epricebook.adlinktech.com/Price/PriceList.aspx"
EIP_LOGIN_URL = (
    "https://eip.adlinktech.com/GAIA/Account/Logon"
    "?returnUrl=https%3a%2f%2feip.adlinktech.com%2fGaia%2fportal%2fTPHQ"
)
ANYDOOR_URL = "https://anydoor.adlinktech.com/AnydoorWebNew"


def _parse_numeric(s):
    """從含數字的字串中擷取第一個浮點數，失敗回傳 None。"""
    m = re.search(r"[\d,]+\.?\d*", (s or "").replace(",", ""))
    return float(m.group().replace(",", "")) if m else None


def _fetch_cost_highest_from_anydoor(driver, product_numbers, username, password, period_start="202501", on_progress=None):
    """
    1. 先登入 EIP  2. 進 anydoor SAP Report  3. 逐 PN 各查一次
    回傳 (dict, diag_msg)
    """
    result = {}
    if not product_numbers:
        return result, "product_numbers 為空"

    step = "初始化"
    try:
        wait = WebDriverWait(driver, 10)

        # ════ EIP 登入 ════
        step = "EIP-Login: 開啟"
        if on_progress:
            on_progress(3, "登入 EIP 中…")
        driver.get(EIP_LOGIN_URL)

        if "Account/Logon" in driver.current_url:
            step = "EIP-Login: 填入帳密"
            acct_input = wait.until(
                EC.presence_of_element_located((By.ID, "account"))
            )
            pwd_input = driver.find_element(By.ID, "password")
            acct_input.clear()
            acct_input.send_keys(username)
            pwd_input.clear()
            pwd_input.send_keys(password)

            step = "EIP-Login: 點擊登入"
            login_btn = driver.find_element(
                By.CSS_SELECTOR, "button.btn-primary.btn-block"
            )
            login_btn.click()

            step = "EIP-Login: 等待完成"
            try:
                WebDriverWait(driver, 10).until(
                    lambda d: "Account/Logon" not in d.current_url
                )
            except TimeoutException:
                return result, f"EIP 登入失敗: 仍在登入頁 URL={driver.current_url}"

        # ════ 進入 anydoor（URL 嵌入帳密通過 HTTP 認證）════
        step = "Step1: 開啟 anydoor (HTTP auth)"
        if on_progress:
            on_progress(3, "開啟 Anydoor…")
        safe_user = quote(username, safe="")
        safe_pass = quote(password, safe="")
        driver.get(
            f"https://{safe_user}:{safe_pass}"
            f"@anydoor.adlinktech.com/AnydoorWebNew"
        )
        time.sleep(2)

        src500 = driver.page_source[:500]
        if "401" in src500 or "Unauthorized" in src500:
            return result, f"Step1: anydoor 認證失敗"

        step = "Step2: 選擇 SAP Report"
        sys_select = wait.until(
            EC.presence_of_element_located((By.ID, "SystemList"))
        )
        time.sleep(0.3)
        options = sys_select.find_elements(By.TAG_NAME, "option")
        found_sap = False
        for opt in options:
            if "SAP Report" in opt.text:
                opt.click()
                found_sap = True
                break
        if not found_sap:
            return result, f"Step2: 無 SAP Report 選項"
        time.sleep(0.5)

        step = "Step3: 點擊報表連結"
        if on_progress:
            on_progress(3, "開啟 SAP Report…")
        report_link = wait.until(
            EC.presence_of_element_located(
                (By.XPATH,
                 "//a[contains(text(),'Standard/Actual Cost Report')]")
            )
        )
        try:
            if not report_link.is_displayed():
                parent_a = report_link.find_element(
                    By.XPATH,
                    "./ancestor::ul[contains(@class,'nav-second-level')]"
                    "/preceding-sibling::a",
                )
                parent_a.click()
                time.sleep(0.2)
        except Exception:
            pass
        report_link.click()
        time.sleep(0.5)

        step = "Step4: 切換 iframe"
        if on_progress:
            on_progress(3, "準備查詢介面…")
        iframe = wait.until(
            EC.presence_of_element_located((By.ID, "pageContent"))
        )
        driver.switch_to.frame(iframe)
        time.sleep(0.3)

        today_str = datetime.now().strftime("%Y%m")
        diag_parts = []

        for pn_idx, pn in enumerate(product_numbers):
            pn_label = f"PN[{pn_idx+1}/{len(product_numbers)}] {pn}"
            if on_progress:
                on_progress(4, f"查詢 Cost(Highest) ({pn_idx+1}/{len(product_numbers)})")
            step = f"{pn_label}: 填表單"

            ps_input = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='PeriodStart']")
                )
            )
            ps_input.clear()
            ps_input.send_keys(period_start)

            pe_input = driver.find_element(
                By.CSS_SELECTOR, "input[ng-model='Search.PeriodEnd']"
            )
            pe_input.clear()
            pe_input.send_keys(today_str)

            pn_area = driver.find_element(
                By.CSS_SELECTOR, "textarea[ng-model='Search.PN']"
            )
            pn_area.clear()
            pn_area.send_keys(pn)

            step = f"{pn_label}: Search"
            search_btn = driver.find_element(
                By.CSS_SELECTOR, "input[type='button'][value='Search']"
            )
            # 記錄搜尋前的完整指紋（行數 + 第一列全文 + 最後一列前80字元）
            old_fp = driver.execute_script("""
                var rows = document.querySelectorAll('table tbody tr');
                if (rows.length === 0) return '';
                var first = rows[0].textContent.trim();
                var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
                return rows.length + ':' + first + '|' + last;
            """) or ""
            search_btn.click()
            time.sleep(0.5)  # 等待頁面開始載入

            step = f"{pn_label}: 等待"

            def data_ready(drv):
                return drv.execute_script("""
                    var rows = document.querySelectorAll('table tbody tr');
                    if (rows.length === 0) return false;
                    var first = rows[0].textContent.trim();
                    var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
                    var fp = rows.length + ':' + first + '|' + last;
                    return fp !== arguments[0];
                """, old_fp)

            if old_fp:
                try:
                    WebDriverWait(driver, 30, poll_frequency=0.5).until(data_ready)
                except TimeoutException:
                    diag_parts.append(f"{pn}: 逾時")
                    continue
            else:
                def has_rows(drv):
                    return drv.execute_script(
                        "return document.querySelectorAll('table tbody tr').length > 0"
                    )
                try:
                    WebDriverWait(driver, 30, poll_frequency=0.5).until(has_rows)
                except TimeoutException:
                    diag_parts.append(f"{pn}: 逾時")
                    continue

            step = f"{pn_label}: JS 批次提取"
            tbl = driver.execute_script("""
                var ths = document.querySelectorAll('table thead tr th');
                var hdr = [];
                for (var i = 0; i < ths.length; i++)
                    hdr.push(ths[i].textContent.trim().replace(/\\u00a0/g,' '));
                var rows = document.querySelectorAll('table tbody tr');
                var data = [];
                for (var i = 0; i < rows.length; i++) {
                    var cells = rows[i].querySelectorAll('td');
                    var r = [];
                    for (var j = 0; j < cells.length; j++)
                        r.push(cells[j].textContent.trim());
                    data.push(r);
                }
                return {h: hdr, d: data};
            """)

            col = {}
            for i, txt in enumerate(tbl["h"]):
                if "ACT Cost" in txt or txt.startswith("ACT Cost"):
                    col["act_cost"] = i
                elif "ACT-Mat" in txt:
                    col["act_mat"] = i
                elif "ACT-MOH" in txt:
                    col["act_moh"] = i
                if (txt == "PN" or txt.startswith("PN")) \
                        and "PNText" not in txt and "PN Type" not in txt:
                    col["pn"] = i
                if txt == "Currency":
                    col["currency"] = i

            needed = ["act_cost", "act_mat", "act_moh", "pn"]
            if not all(k in col for k in needed):
                diag_parts.append(f"{pn}: 表頭缺欄位")
                continue

            max_idx = max(col.values())
            costs_for_pn = []
            cur_idx = col.get("currency")

            def _collect_rows(rows_data):
                count = 0
                for rd in rows_data:
                    if len(rd) <= max_idx:
                        continue
                    cv = _parse_numeric(rd[col["act_cost"]])
                    mv = _parse_numeric(rd[col["act_mat"]])
                    hv = _parse_numeric(rd[col["act_moh"]])
                    if cv is not None and cv > 0 and mv is not None and mv > 0 and hv is not None and hv > 0:
                        rc = rd[cur_idx] if cur_idx is not None and len(rd) > cur_idx else ""
                        costs_for_pn.append((cv, rc))
                    count += 1
                return count

            total_data_rows = _collect_rows(tbl["d"])

            page_num = 1
            while True:
                try:
                    next_info = driver.execute_script("""
                        var items = document.querySelectorAll('ul.pagination li');
                        for (var i = 0; i < items.length; i++) {
                            var a = items[i].querySelector('a');
                            if (!a) continue;
                            var t = a.textContent.trim();
                            if (t==='>'||t==='›'||t==='»'||t==='Next') {
                                if (items[i].className.indexOf('disabled') === -1) {
                                    a.click();
                                    return true;
                                }
                            }
                        }
                        return false;
                    """)
                    if not next_info:
                        break
                    page_num += 1
                    # 每翻一頁就送出進度，避免長時間無訊息導致 SSE 逾時
                    if on_progress:
                        on_progress(4, f"查詢 Cost(Highest) ({pn_idx+1}/{len(product_numbers)}) - 第{page_num}頁")
                    time.sleep(0.5)
                    page_data = driver.execute_script("""
                        var rows = document.querySelectorAll('table tbody tr');
                        var data = [];
                        for (var i = 0; i < rows.length; i++) {
                            var cells = rows[i].querySelectorAll('td');
                            var r = [];
                            for (var j = 0; j < cells.length; j++)
                                r.push(cells[j].textContent.trim());
                            data.push(r);
                        }
                        return data;
                    """)
                    total_data_rows += _collect_rows(page_data)
                except Exception:
                    break

            if costs_for_pn:
                result[pn] = costs_for_pn
            else:
                diag_parts.append(f"{pn}: {total_data_rows} 列無有效資料")

        driver.switch_to.default_content()

        if diag_parts and not result:
            return result, "anydoor 全部 PN 失敗:\n" + "\n".join(diag_parts)
        if diag_parts:
            return result, "部分 PN 失敗:\n" + "\n".join(diag_parts)
        return result, None

    except Exception as ex:
        try:
            driver.switch_to.default_content()
        except Exception:
            pass
        return result, f"{step} 例外: {type(ex).__name__}: {ex}"


def _fetch_bom_with_cost(username, password, pn, plant, on_progress=None):
    """
    到 Anydoor > SAP Report > BOM with cost 查詢指定 PN 的 BOM 物料資料，
    篩選含 ARM / DDR / eMMC 的料，回傳 (list_of_dicts, error_msg)。
    """
    driver = None
    step = "初始化"
    try:
        if on_progress:
            on_progress(0, "啟動瀏覽器")

        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--auth-server-allowlist=*.adlinktech.com")
        options.add_argument("--remote-debugging-port=0")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options,
        )
        wait = WebDriverWait(driver, 15)

        # ════ EIP 登入 ════
        step = "EIP-Login: 開啟"
        if on_progress:
            on_progress(1, "登入 EIP 中…")
        driver.get(EIP_LOGIN_URL)

        if "Account/Logon" in driver.current_url:
            step = "EIP-Login: 填入帳密"
            acct_input = wait.until(
                EC.presence_of_element_located((By.ID, "account"))
            )
            pwd_input = driver.find_element(By.ID, "password")
            acct_input.clear()
            acct_input.send_keys(username)
            pwd_input.clear()
            pwd_input.send_keys(password)

            step = "EIP-Login: 點擊登入"
            login_btn = driver.find_element(
                By.CSS_SELECTOR, "button.btn-primary.btn-block"
            )
            login_btn.click()

            step = "EIP-Login: 等待完成"
            try:
                WebDriverWait(driver, 10).until(
                    lambda d: "Account/Logon" not in d.current_url
                )
            except TimeoutException:
                return [], f"EIP 登入失敗: 仍在登入頁 URL={driver.current_url}"

        # ════ 進入 anydoor（URL 嵌入帳密通過 HTTP 認證）════
        step = "Step1: 開啟 anydoor (HTTP auth)"
        if on_progress:
            on_progress(1, "開啟 Anydoor…")
        safe_user = quote(username, safe="")
        safe_pass = quote(password, safe="")
        driver.get(
            f"https://{safe_user}:{safe_pass}"
            f"@anydoor.adlinktech.com/AnydoorWebNew"
        )
        time.sleep(2)

        src500 = driver.page_source[:500]
        if "401" in src500 or "Unauthorized" in src500:
            return [], "anydoor 認證失敗"

        # ════ 選擇 SAP Report ════
        step = "Step2: 選擇 SAP Report"
        sys_select = wait.until(
            EC.presence_of_element_located((By.ID, "SystemList"))
        )
        time.sleep(0.3)
        found_sap = False
        for opt in sys_select.find_elements(By.TAG_NAME, "option"):
            if "SAP Report" in opt.text:
                opt.click()
                found_sap = True
                break
        if not found_sap:
            return [], "無 SAP Report 選項"
        time.sleep(0.5)

        # ════ 點擊 BOM with cost 報表（先展開父選單再點子項）════
        step = "Step3: 展開 BOM With Cost 選單"
        if on_progress:
            on_progress(2, "開啟 BOM with cost 報表…")
        # 找到所有含 "bom with cost" 的連結
        bom_links = driver.find_elements(
            By.XPATH,
            "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bom with cost')]"
        )
        if not bom_links:
            bom_links = [wait.until(EC.presence_of_element_located(
                (By.XPATH, "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bom')]")
            ))]
        if len(bom_links) >= 2:
            # 第一個是父選單，點擊展開；第二個是子項目
            bom_links[0].click()
            time.sleep(0.3)
            step = "Step3: 點擊 BOM With Cost 子項"
            bom_links[1].click()
        else:
            # 只找到一個，嘗試先展開父選單
            link = bom_links[0]
            try:
                if not link.is_displayed():
                    parent_a = link.find_element(
                        By.XPATH,
                        "./ancestor::ul[contains(@class,'nav-second-level')]"
                        "/preceding-sibling::a",
                    )
                    parent_a.click()
                    time.sleep(0.2)
            except Exception:
                pass
            link.click()
        time.sleep(0.5)

        # ════ 切換 iframe ════
        step = "Step4: 切換 iframe"
        iframe = wait.until(
            EC.presence_of_element_located((By.ID, "pageContent"))
        )
        driver.switch_to.frame(iframe)
        time.sleep(1)

        # BOM with cost 頁面可能還有內層 iframe，嘗試偵測並切入
        try:
            inner_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if inner_iframes:
                driver.switch_to.frame(inner_iframes[0])
                time.sleep(0.5)
        except Exception:
            pass

        # ════ 填入 PN 與 Plant ════
        step = "Step5: 填入 PN 與 Plant"
        if on_progress:
            on_progress(2, f"填入 PN: {pn}，Plant: {plant}")

        # 等待頁面載入完成後再找元素，加長等待時間
        pn_input = None
        for sel in ["#MATERIAL_NO", "textarea[ng-model='Search.PN']", "input[ng-model='Search.PN']",
                     "textarea[name='PN']", "input[name='PN']"]:
            try:
                pn_input = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                )
                break
            except TimeoutException:
                continue
        if pn_input is None:
            # 診斷：dump 頁面上所有 input/select/textarea 的 id、name、type
            diag = driver.execute_script("""
                var result = [];
                var els = document.querySelectorAll('input, select, textarea');
                for (var i = 0; i < els.length; i++) {
                    result.push({
                        tag: els[i].tagName,
                        id: els[i].id || '',
                        name: els[i].name || '',
                        type: els[i].type || '',
                        cls: els[i].className || ''
                    });
                }
                // 也檢查有無 iframe
                var iframes = document.querySelectorAll('iframe');
                for (var i = 0; i < iframes.length; i++) {
                    result.push({tag: 'IFRAME', id: iframes[i].id || '', name: iframes[i].name || '', src: (iframes[i].src || '').substring(0, 100)});
                }
                return JSON.stringify(result);
            """)
            return [], f"找不到 PN 輸入欄位。頁面元素: {diag}"
        pn_input.clear()
        pn_input.send_keys(pn)

        # Plant 下拉選單（BOM with cost 頁面欄位為 select#PLANT）
        try:
            plant_el = driver.find_element(By.CSS_SELECTOR, "#PLANT")
            plant_select = Select(plant_el)
            plant_select.select_by_value(plant)
        except Exception:
            # fallback: 用 JS 直接設定
            driver.execute_script("""
                var sel = document.getElementById('PLANT');
                if (sel) {
                    sel.value = arguments[0];
                    sel.dispatchEvent(new Event('change', {bubbles: true}));
                }
            """, plant)

        # ════ 點擊 Search ════
        step = "Step6: Search"
        if on_progress:
            on_progress(3, "查詢 BOM 資料…")

        search_btn = driver.find_element(
            By.CSS_SELECTOR, "input[type='button'][value='Search']"
        )
        # Kendo-aware fingerprint（Search 前記錄，Search 後比對變化）
        old_fp = driver.execute_script("""
            try {
                var ssEl = document.querySelector('.k-spreadsheet');
                if (ssEl && window.jQuery) {
                    var ss = jQuery(ssEl).data('kendoSpreadsheet');
                    if (ss) {
                        var json = ss.activeSheet().toJSON();
                        var sRows = json.rows || [];
                        if (sRows.length > 1) {
                            var fc = sRows[1].cells || [];
                            var fv = fc.length > 0 ? String(fc[0].value || '') : '';
                            return sRows.length + ':' + fv;
                        }
                        return sRows.length + ':';
                    }
                }
            } catch(e) {}
            var rows = document.querySelectorAll('table tbody tr');
            if (rows.length === 0) return '';
            var first = rows[0].textContent.trim();
            var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
            return rows.length + ':' + first + '|' + last;
        """) or ""
        search_btn.click()
        time.sleep(0.5)

        # 等待資料載入（Kendo-aware）
        step = "Step7: 等待資料"
        def data_ready(drv):
            return drv.execute_script("""
                try {
                    var ssEl = document.querySelector('.k-spreadsheet');
                    if (ssEl && window.jQuery) {
                        var ss = jQuery(ssEl).data('kendoSpreadsheet');
                        if (ss) {
                            var json = ss.activeSheet().toJSON();
                            var sRows = json.rows || [];
                            if (sRows.length > 1) {
                                var fc = sRows[1].cells || [];
                                var fv = fc.length > 0 ? String(fc[0].value || '') : '';
                                var fp2 = sRows.length + ':' + fv;
                                return fp2 !== arguments[0];
                            }
                            return sRows.length > 0;
                        }
                    }
                } catch(e) {}
                var rows = document.querySelectorAll('table tbody tr');
                if (rows.length === 0) return false;
                var first = rows[0].textContent.trim();
                var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
                var fp = rows.length + ':' + first + '|' + last;
                return fp !== arguments[0];
            """, old_fp)

        try:
            WebDriverWait(driver, 30, poll_frequency=0.5).until(data_ready)
        except TimeoutException:
            return [], f"BOM 查詢逾時: {pn}"

        # ════ 點擊 material Data 1 頁籤 ════
        step = "Step8: 點擊 material Data 1"
        if on_progress:
            on_progress(3, "切換到 material Data 1…")
        try:
            tab = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "span.k-link[title='Material Data 1']")
                )
            )
            tab.click()
            time.sleep(1)
        except TimeoutException:
            pass  # 頁籤可能預設就是 Material Data 1

        # ════ 提取表格資料 ════
        step = "Step9: 提取資料"
        if on_progress:
            on_progress(4, "提取 BOM 資料…")

        BOM_PN_PREFIXES = ("71-", "72-")

        def _build_col_map(headers):
            """建立欄位索引對照表"""
            col_map = {}
            want_cols = {
                "part_no": ["PART NO", "Part No", "PartNo", "Material"],
                "part_spec": ["Part Spec", "PartSpec", "Description", "Material Description"],
                "lead_time": ["Lead Time", "LeadTime", "Lead time"],
                "accumu_qty": ["Accumu Qty", "AccumuQty", "Accum Qty", "Quantity"],
                "last_purchasing_price": ["Last Purchasing Price", "LastPurchasingPrice", "Last Price"],
                "local_currency": ["Local Currency", "LocalCurrency", "Currency"],
            }
            for key, aliases in want_cols.items():
                for i, h in enumerate(headers):
                    h_clean = h.strip()
                    for alias in aliases:
                        if alias.lower() in h_clean.lower():
                            col_map[key] = i
                            break
                    if key in col_map:
                        break
            return col_map, want_cols

        def _extract_and_filter(rows_data, headers):
            """從表格資料中篩選 PART NO 以 71 或 72 開頭的列"""
            col_map, want_cols = _build_col_map(headers)
            items = []
            pn_idx = col_map.get("part_no")
            for row in rows_data:
                pn_val = row[pn_idx].strip() if pn_idx is not None and pn_idx < len(row) else ""
                if not pn_val or not pn_val.startswith(BOM_PN_PREFIXES):
                    continue
                item = {}
                for key in want_cols:
                    idx = col_map.get(key)
                    item[key] = row[idx] if idx is not None and idx < len(row) else ""
                items.append(item)
            return items

        # Kendo Spreadsheet API 提取資料（BOM with cost 結果是 Kendo Spreadsheet）
        JS_GET_TABLE = """
            // 嘗試 Kendo Spreadsheet
            try {
                var ssEl = document.querySelector('.k-spreadsheet');
                if (ssEl && window.jQuery) {
                    var ss = jQuery(ssEl).data('kendoSpreadsheet');
                    if (ss) {
                        var sheet = ss.activeSheet();
                        var json = sheet.toJSON();
                        var sRows = json.rows || [];
                        if (sRows.length > 0) {
                            var maxCol = 0;
                            for (var r = 0; r < sRows.length; r++) {
                                var cs = sRows[r].cells || [];
                                for (var c = 0; c < cs.length; c++) {
                                    var ci = cs[c].index != null ? cs[c].index : c;
                                    if (ci > maxCol) maxCol = ci;
                                }
                            }
                            var hdr = [];
                            var hdrCells = sRows[0].cells || [];
                            for (var col = 0; col <= maxCol; col++) {
                                var val = '';
                                for (var c = 0; c < hdrCells.length; c++) {
                                    var ci2 = hdrCells[c].index != null ? hdrCells[c].index : c;
                                    if (ci2 === col) { val = String(hdrCells[c].value || '').trim(); break; }
                                }
                                hdr.push(val);
                            }
                            var data = [];
                            for (var r = 1; r < sRows.length; r++) {
                                var rowCells = sRows[r].cells || [];
                                var rowData = [];
                                var hasVal = false;
                                for (var col = 0; col <= maxCol; col++) {
                                    var v = '';
                                    for (var c = 0; c < rowCells.length; c++) {
                                        var ci3 = rowCells[c].index != null ? rowCells[c].index : c;
                                        if (ci3 === col) { v = String(rowCells[c].value || '').trim(); break; }
                                    }
                                    rowData.push(v);
                                    if (v) hasVal = true;
                                }
                                if (hasVal) data.push(rowData);
                            }
                            return {h: hdr, d: data, src: 'kendo'};
                        }
                    }
                }
            } catch(e) {}
            // Fallback: 普通 HTML table
            var ths = document.querySelectorAll('table thead tr th');
            var hdr = [];
            for (var i = 0; i < ths.length; i++)
                hdr.push(ths[i].textContent.trim().replace(/\\u00a0/g,' '));
            var rows = document.querySelectorAll('table tbody tr');
            var data = [];
            for (var i = 0; i < rows.length; i++) {
                var cells = rows[i].querySelectorAll('td');
                var r = [];
                for (var j = 0; j < cells.length; j++)
                    r.push(cells[j].textContent.trim());
                data.push(r);
            }
            return {h: hdr, d: data, src: 'table'};
        """

        # Kendo Spreadsheet 無翻頁（資料一次全在 sheet 裡），所以不需要 JS_NEXT_PAGE/JS_GET_ROWS
        # 但保留 fallback 給普通 table 的情況
        JS_NEXT_PAGE = """
            var items = document.querySelectorAll('ul.pagination li');
            for (var i = 0; i < items.length; i++) {
                var a = items[i].querySelector('a');
                if (!a) continue;
                var t = a.textContent.trim();
                if (t==='>'||t==='›'||t==='»'||t==='Next') {
                    if (items[i].className.indexOf('disabled') === -1) {
                        a.click();
                        return true;
                    }
                }
            }
            return false;
        """
        JS_GET_ROWS = """
            var rows = document.querySelectorAll('table tbody tr');
            var data = [];
            for (var i = 0; i < rows.length; i++) {
                var cells = rows[i].querySelectorAll('td');
                var r = [];
                for (var j = 0; j < cells.length; j++)
                    r.push(cells[j].textContent.trim());
                data.push(r);
            }
            return data;
        """

        # 取得表頭與全部資料（Kendo Spreadsheet 一次全拿到，不需翻頁）
        tbl = driver.execute_script(JS_GET_TABLE)
        is_kendo = tbl.get("src") == "kendo"
        all_items = _extract_and_filter(tbl["d"], tbl["h"])

        # 翻頁處理（只有普通 table 才需要，Kendo Spreadsheet 跳過）
        if not is_kendo:
            page_num = 1
            while True:
                try:
                    next_info = driver.execute_script(JS_NEXT_PAGE)
                    if not next_info:
                        break
                    page_num += 1
                    if on_progress:
                        on_progress(4, f"提取 BOM 資料 - 第{page_num}頁…")
                    time.sleep(0.5)
                    page_data = driver.execute_script(JS_GET_ROWS)
                    all_items.extend(_extract_and_filter(page_data, tbl["h"]))
                except Exception:
                    break

        driver.switch_to.default_content()

        if not all_items:
            return [], f"BOM 中未找到 PART NO 以 71-/72- 開頭的物料 ({pn})"

        return all_items, None

    except Exception as ex:
        try:
            if driver:
                driver.switch_to.default_content()
        except Exception:
            pass
        return [], f"{step} 例外: {type(ex).__name__}: {ex}"
    finally:
        if driver:
            try:
                import threading as _thr
                quit_thread = _thr.Thread(target=driver.quit, daemon=True)
                quit_thread.start()
                quit_thread.join(timeout=15)
            except Exception:
                pass


def _fetch_batch_bom_with_cost(username, password, pn_list, plant, on_progress=None):
    """
    批次 BOM 查詢：開一個 Chrome，登入 EIP + anydoor 一次，
    逐 PN 查 BOM with cost，篩選 71/72 開頭物料，回傳彙整結果。
    回傳 (dict{pn: [items]}, list[warnings])
    """
    BOM_PN_PREFIXES = ("71-", "72-")
    driver = None
    step = "初始化"
    results = {}       # {pn: [item_dicts]}
    warnings = []
    total = len(pn_list)

    def _extract_and_filter(rows_data, headers):
        """從表格資料中篩選 PART NO 以 71 或 72 開頭的列"""
        col_map = {}
        want_cols = {
            "part_no": ["PART NO", "Part No", "PartNo", "Material"],
            "part_spec": ["Part Spec", "PartSpec", "Description", "Material Description"],
            "lead_time": ["Lead Time", "LeadTime", "Lead time"],
            "accumu_qty": ["Accumu Qty", "AccumuQty", "Accum Qty", "Quantity"],
            "last_purchasing_price": ["Last Purchasing Price", "LastPurchasingPrice", "Last Price"],
            "local_currency": ["Local Currency", "LocalCurrency", "Currency"],
        }
        for key, aliases in want_cols.items():
            for i, h in enumerate(headers):
                h_clean = h.strip()
                for alias in aliases:
                    if alias.lower() in h_clean.lower():
                        col_map[key] = i
                        break
                if key in col_map:
                    break

        items = []
        pn_idx = col_map.get("part_no")
        for row in rows_data:
            pn_val = row[pn_idx].strip() if pn_idx is not None and pn_idx < len(row) else ""
            if not pn_val or not pn_val.startswith(BOM_PN_PREFIXES):
                continue
            item = {}
            for key in want_cols:
                idx = col_map.get(key)
                item[key] = row[idx] if idx is not None and idx < len(row) else ""
            items.append(item)
        return items

    try:
        if on_progress:
            on_progress(0, "啟動瀏覽器")

        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--auth-server-allowlist=*.adlinktech.com")
        options.add_argument("--remote-debugging-port=0")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options,
        )
        wait = WebDriverWait(driver, 15)

        # ════ EIP 登入 ════
        step = "EIP-Login"
        if on_progress:
            on_progress(1, "登入 EIP 中…")
        driver.get(EIP_LOGIN_URL)

        if "Account/Logon" in driver.current_url:
            acct_input = wait.until(EC.presence_of_element_located((By.ID, "account")))
            pwd_input = driver.find_element(By.ID, "password")
            acct_input.clear()
            acct_input.send_keys(username)
            pwd_input.clear()
            pwd_input.send_keys(password)
            login_btn = driver.find_element(By.CSS_SELECTOR, "button.btn-primary.btn-block")
            login_btn.click()
            try:
                WebDriverWait(driver, 10).until(lambda d: "Account/Logon" not in d.current_url)
            except TimeoutException:
                return {}, [f"EIP 登入失敗: 仍在登入頁 URL={driver.current_url}"]

        # ════ 進入 anydoor ════
        step = "開啟 anydoor"
        if on_progress:
            on_progress(1, "開啟 Anydoor…")
        safe_user = quote(username, safe="")
        safe_pass = quote(password, safe="")
        driver.get(f"https://{safe_user}:{safe_pass}@anydoor.adlinktech.com/AnydoorWebNew")
        time.sleep(2)

        src500 = driver.page_source[:500]
        if "401" in src500 or "Unauthorized" in src500:
            return {}, ["anydoor 認證失敗"]

        # ════ 選擇 SAP Report ════
        step = "選擇 SAP Report"
        sys_select = wait.until(EC.presence_of_element_located((By.ID, "SystemList")))
        time.sleep(0.3)
        found_sap = False
        for opt in sys_select.find_elements(By.TAG_NAME, "option"):
            if "SAP Report" in opt.text:
                opt.click()
                found_sap = True
                break
        if not found_sap:
            return {}, ["無 SAP Report 選項"]
        time.sleep(0.5)

        # ════ 點擊 BOM with cost ════
        step = "展開 BOM With Cost"
        if on_progress:
            on_progress(2, "開啟 BOM with cost 報表…")
        bom_links = driver.find_elements(
            By.XPATH,
            "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bom with cost')]"
        )
        if not bom_links:
            bom_links = [wait.until(EC.presence_of_element_located(
                (By.XPATH, "//a[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bom')]")
            ))]
        if len(bom_links) >= 2:
            bom_links[0].click()
            time.sleep(0.3)
            bom_links[1].click()
        else:
            link = bom_links[0]
            try:
                if not link.is_displayed():
                    parent_a = link.find_element(
                        By.XPATH, "./ancestor::ul[contains(@class,'nav-second-level')]/preceding-sibling::a")
                    parent_a.click()
                    time.sleep(0.2)
            except Exception:
                pass
            link.click()
        time.sleep(0.5)

        # ════ 切換 iframe ════
        step = "切換 iframe"
        iframe = wait.until(EC.presence_of_element_located((By.ID, "pageContent")))
        driver.switch_to.frame(iframe)
        time.sleep(1)
        try:
            inner_iframes = driver.find_elements(By.TAG_NAME, "iframe")
            if inner_iframes:
                driver.switch_to.frame(inner_iframes[0])
                time.sleep(0.5)
        except Exception:
            pass

        # ════ 找 PN 輸入欄位（只找一次）════
        step = "找 PN 輸入欄位"
        pn_input = None
        for sel in ["#MATERIAL_NO", "textarea[ng-model='Search.PN']", "input[ng-model='Search.PN']",
                     "textarea[name='PN']", "input[name='PN']"]:
            try:
                pn_input = WebDriverWait(driver, 8).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, sel))
                )
                break
            except TimeoutException:
                continue
        if pn_input is None:
            return {}, ["找不到 PN 輸入欄位"]

        # 取得表頭（第一次就抓，後續頁面共用）
        cached_headers = None

        # ════ 逐 PN 查詢 ════
        for pn_idx, pn in enumerate(pn_list):
            pn = pn.strip()
            if not pn:
                continue
            step = f"查詢 {pn}"
            if on_progress:
                on_progress(3, f"查詢 BOM ({pn_idx+1}/{total}) — {pn}")

            # 清空並填入 PN
            pn_input.clear()
            pn_input.send_keys(pn)

            # 設定 Plant
            try:
                plant_el = driver.find_element(By.CSS_SELECTOR, "#PLANT")
                plant_select = Select(plant_el)
                plant_select.select_by_value(plant)
            except Exception:
                driver.execute_script("""
                    var sel = document.getElementById('PLANT');
                    if (sel) { sel.value = arguments[0]; sel.dispatchEvent(new Event('change', {bubbles: true})); }
                """, plant)

            # 記錄舊指紋以偵測資料更新
            old_fp = driver.execute_script("""
                var rows = document.querySelectorAll('table tbody tr');
                if (rows.length === 0) return '';
                var first = rows[0].textContent.trim();
                var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
                return rows.length + ':' + first + '|' + last;
            """) or ""

            # 點擊 Search
            try:
                search_btn = driver.find_element(By.CSS_SELECTOR, "input[type='button'][value='Search']")
                search_btn.click()
            except Exception as ex:
                warnings.append(f"{pn}: 找不到 Search 按鈕 — {ex}")
                continue
            time.sleep(0.5)

            # 等待資料載入
            def data_ready(drv):
                return drv.execute_script("""
                    var rows = document.querySelectorAll('table tbody tr');
                    if (rows.length === 0) return false;
                    var first = rows[0].textContent.trim();
                    var last = rows[rows.length - 1].textContent.trim().substring(0, 80);
                    var fp = rows.length + ':' + first + '|' + last;
                    return fp !== arguments[0];
                """, old_fp)

            try:
                if old_fp:
                    WebDriverWait(driver, 30, poll_frequency=0.5).until(data_ready)
                else:
                    WebDriverWait(driver, 30, poll_frequency=0.5).until(
                        lambda drv: drv.execute_script("return document.querySelectorAll('table tbody tr').length > 0")
                    )
            except TimeoutException:
                warnings.append(f"{pn}: BOM 查詢逾時")
                continue

            # 切換到 material Data 1 頁籤
            try:
                tab = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, "span.k-link[title='Material Data 1']"))
                )
                tab.click()
                time.sleep(1)
            except TimeoutException:
                pass

            # 提取表頭（第一次查時取得）
            tbl = driver.execute_script("""
                var ths = document.querySelectorAll('table thead tr th');
                var hdr = [];
                for (var i = 0; i < ths.length; i++)
                    hdr.push(ths[i].textContent.trim().replace(/\\u00a0/g,' '));
                var rows = document.querySelectorAll('table tbody tr');
                var data = [];
                for (var i = 0; i < rows.length; i++) {
                    var cells = rows[i].querySelectorAll('td');
                    var r = [];
                    for (var j = 0; j < cells.length; j++)
                        r.push(cells[j].textContent.trim());
                    data.push(r);
                }
                return {h: hdr, d: data};
            """)
            if cached_headers is None:
                cached_headers = tbl["h"]

            all_items = _extract_and_filter(tbl["d"], cached_headers)

            # 翻頁
            page_num = 1
            while True:
                try:
                    next_info = driver.execute_script("""
                        var items = document.querySelectorAll('ul.pagination li');
                        for (var i = 0; i < items.length; i++) {
                            var a = items[i].querySelector('a');
                            if (!a) continue;
                            var t = a.textContent.trim();
                            if (t==='>'||t==='›'||t==='»'||t==='Next') {
                                if (items[i].className.indexOf('disabled') === -1) {
                                    a.click();
                                    return true;
                                }
                            }
                        }
                        return false;
                    """)
                    if not next_info:
                        break
                    page_num += 1
                    if on_progress:
                        on_progress(3, f"查詢 BOM ({pn_idx+1}/{total}) — {pn} 第{page_num}頁")
                    time.sleep(0.5)
                    page_data = driver.execute_script("""
                        var rows = document.querySelectorAll('table tbody tr');
                        var data = [];
                        for (var i = 0; i < rows.length; i++) {
                            var cells = rows[i].querySelectorAll('td');
                            var r = [];
                            for (var j = 0; j < cells.length; j++)
                                r.push(cells[j].textContent.trim());
                            data.push(r);
                        }
                        return data;
                    """)
                    all_items.extend(_extract_and_filter(page_data, cached_headers))
                except Exception:
                    break

            if all_items:
                results[pn] = all_items
            else:
                warnings.append(f"{pn}: 未找到 71/72 開頭物料")

        return results, warnings

    except Exception as ex:
        warnings.append(f"{step} 例外: {type(ex).__name__}: {ex}")
        return results, warnings
    finally:
        if driver:
            try:
                import threading as _thr
                quit_thread = _thr.Thread(target=driver.quit, daemon=True)
                quit_thread.start()
                quit_thread.join(timeout=15)
            except Exception:
                pass


def _extract_product_name_from_cell(cell_text):
    """
    從 PartNumber/Name/Spec 欄位擷取用於比對的 Product Name；
    取 (G) 或 (EA) 中較早出現之前的內容。回傳 (擷取出的名稱, 該段完整字串)，若無則 ("", "").
    """
    if not cell_text or not cell_text.strip():
        return "", ""
    s = cell_text.strip()
    s = re.sub(r"\b(?!9\d-)\d+-\d+[-A-Za-z0-9]*\b", " ", s)
    s = " ".join(s.split())
    for marker in ["(G)", "(EA)"]:
        if marker in s:
            s = s.split(marker)[0]
            break
    before_g = s.strip()
    tokens = before_g.split()
    extracted = tokens[-1] if tokens else ""
    return extracted, before_g


def _cell_matches_product(cell_text, want):
    """
    判斷欄位文字是否匹配搜尋的產品名稱。
    先用 _extract_product_name_from_cell 取最後 token 比對；
    若不符，再檢查 want 是否出現在清理後文字的 token 中（處理多行含 spec 的情況，
    例如 "92-97108-0020\\nNuDAM ND-6150\\n8DI, 8DO, Modbus RTU"）。
    回傳 (是否匹配, before_g 字串)。
    """
    extracted, before_g = _extract_product_name_from_cell(cell_text)
    if extracted == want:
        return True, before_g
    # 多行欄位：want 可能不在最後一個 token，改檢查是否存在於 token 列表中
    if want in before_g.split():
        return True, before_g
    # 多字產品名（含空格）：改用子字串比對
    if " " in want and want in before_g:
        return True, before_g
    return False, before_g


def fetch_base_price(username: str, password: str, product_names: list, period_start: str = "202501", on_progress=None) -> Tuple[bool, str, list]:
    """
    使用 Selenium 登入並依序查詢多個產品的 Base Price，再統一到 anydoor 查 Cost(highest)。
    product_names: 產品名稱列表
    回傳 (成功與否, 警告/錯誤訊息, 產品列表 [{product_number_name_spec, base_price, query_index, query_name, ...}])
    """
    def progress(step_idx, label):
        if on_progress:
            on_progress(step_idx, label)

    driver = None
    try:
        progress(0, "啟動瀏覽器")
        options = webdriver.ChromeOptions()
        options.add_argument("--headless=new")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--auth-server-allowlist=*.adlinktech.com")
        options.add_argument("--remote-debugging-port=0")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])

        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options,
        )

        wait = WebDriverWait(driver, 20)

        # 1. 登入
        progress(1, "登入 ePrice")
        driver.get(LOGIN_URL)
        user_input = wait.until(
            EC.presence_of_element_located((By.ID, "TB_Username"))
        )
        pass_input = driver.find_element(By.ID, "TB_Password")
        login_btn = driver.find_element(By.ID, "Btn_OK")
        user_input.clear()
        user_input.send_keys(username)
        pass_input.clear()
        pass_input.send_keys(password)
        login_btn.click()
        wait.until(lambda d: "Login.aspx" not in d.current_url)

        # 2. 依序搜尋每個產品的 Base Price
        all_products = []
        query_warnings = []
        total_queries = len(product_names)

        for q_idx, product_name in enumerate(product_names):
            want = product_name.strip()
            if not want:
                continue

            progress(2, f"搜尋 Base Price ({q_idx+1}/{total_queries}) - {want}")

            driver.get(PRICE_URL)
            wait.until(
                EC.presence_of_element_located(
                    (By.ID, "MainContent_txt_productName")
                )
            )

            pn_input = wait.until(
                EC.element_to_be_clickable((By.ID, "MainContent_txt_productName"))
            )
            driver.execute_script("arguments[0].scrollIntoView(true);", pn_input)
            time.sleep(0.1)

            ac = ActionChains(driver)
            ac.move_to_element(pn_input).click()
            ac.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)
            ac.send_keys(want)
            ac.perform()
            time.sleep(0.1)

            driver.execute_script(
                "var v = arguments[0];"
                "var inp = document.getElementById('MainContent_txt_productName');"
                "if(inp){"
                "  inp.value = v; inp.setAttribute('value', v);"
                "  inp.dispatchEvent(new Event('input',{bubbles:true}));"
                "  inp.dispatchEvent(new Event('change',{bubbles:true}));"
                "}",
                want,
            )
            time.sleep(0.2)

            search_btn = wait.until(
                EC.element_to_be_clickable((By.ID, "MainContent_btn_search"))
            )
            search_btn.click()

            wait_long = WebDriverWait(driver, 15)

            def has_data_row(drv):
                return drv.execute_script("""
                    var g = document.getElementById('MainContent_gv');
                    if (!g) return false;
                    var rows = g.querySelectorAll('tbody > tr');
                    for (var i = 0; i < rows.length; i++) {
                        if (rows[i].querySelectorAll('td').length > 5) return true;
                    }
                    return false;
                """)

            try:
                wait_long.until(has_data_row)
            except TimeoutException:
                query_warnings.append(f"「{want}」查無價格資料或查詢逾時")
                continue

            raw = driver.execute_script("""
                var g = document.getElementById('MainContent_gv');
                var ths = g.querySelectorAll('thead tr th');
                var headers = [];
                for (var i = 0; i < ths.length; i++) headers.push(ths[i].textContent.trim().toLowerCase());
                var rows = g.querySelectorAll('tbody > tr');
                var data = [];
                for (var i = 0; i < rows.length; i++) {
                    var cells = rows[i].querySelectorAll('td');
                    var r = [];
                    for (var j = 0; j < cells.length; j++) {
                        var obj = {t: cells[j].textContent.trim()};
                        var sp = cells[j].querySelector('span[title]');
                        if (sp) obj.title = sp.getAttribute('title') || '';
                        r.push(obj);
                    }
                    data.push(r);
                }
                return {headers: headers, rows: data};
            """)

            header_texts = raw["headers"]
            base_col_index = None
            quantity_col_index = None
            currency_col_index = None
            for i, text in enumerate(header_texts):
                if "base price" in text or "baseprice" in text:
                    base_col_index = i
                if "quantity" in text:
                    quantity_col_index = i
                if "currency" in text:
                    currency_col_index = i

            if base_col_index is None:
                query_warnings.append(f"「{want}」表頭找不到 Base Price 欄位")
                continue
            if quantity_col_index is None:
                query_warnings.append(f"「{want}」表頭找不到 Quantity 欄位")
                continue

            PRODUCT_COL = 1
            product_list = []
            need_cols = max(
                base_col_index, PRODUCT_COL, quantity_col_index,
                currency_col_index if currency_col_index is not None else 0,
            )
            for row_data in raw["rows"]:
                if len(row_data) <= need_cols:
                    continue
                if row_data[quantity_col_index]["t"] != "1":
                    continue
                pn_ns = row_data[PRODUCT_COL]["t"]
                matched, before_g = _cell_matches_product(pn_ns, want)
                if not matched:
                    continue
                # 取 want 在 before_g 中的位置來判斷 prefix
                idx = before_g.find(want)
                prefix = before_g[:idx].rstrip() if idx > 0 else ""
                if prefix and (" for " in prefix or "kit" in prefix.lower()):
                    continue
                currency = row_data[currency_col_index]["t"] if currency_col_index is not None else ""
                bp = row_data[base_col_index]["t"]

                cost_lowest = ""
                title = row_data[base_col_index].get("title", "")
                if title:
                    for line in title.splitlines():
                        if "Cost(Lowest" in line:
                            _, sep, rest = line.partition("：")
                            cost_lowest = (rest if sep else line).strip()
                            break

                product_list.append(
                    {
                        "product_number_name_spec": pn_ns,
                        "currency": currency,
                        "base_price": bp,
                        "cost_lowest": cost_lowest,
                        "query_index": q_idx,
                        "query_name": want,
                    }
                )

            if not product_list:
                query_warnings.append(f"「{want}」查無符合的價格資料")
                continue

            all_products.extend(product_list)

        if not all_products:
            msg = "所有產品皆查無價格資料"
            if query_warnings:
                msg += "\n" + "\n".join(query_warnings)
            return False, msg, []

        # ── 查詢 anydoor SAP Report 取得 Cost(highest) ──
        progress(3, "登入 EIP / Anydoor")
        unique_pns = list({
            p["product_number_name_spec"].split("\n")[0].strip()
            for p in all_products
            if p["product_number_name_spec"].strip()
        })
        cost_highest_map, anydoor_diag = _fetch_cost_highest_from_anydoor(
            driver, unique_pns, username, password, period_start=period_start,
            on_progress=on_progress,
        )
        for p in all_products:
            pn_key = p["product_number_name_spec"].split("\n")[0].strip()
            ch_entries = cost_highest_map.get(pn_key, [])
            p["cost_highest_entries"] = [
                {"cost": cv, "currency": cc} for cv, cc in ch_entries
            ]

        progress(5, "整理結果")

        warn_parts = []
        if query_warnings:
            warn_parts.extend(query_warnings)
        if anydoor_diag:
            warn_parts.append(f"anydoor 診斷: {anydoor_diag}")
        warn = "\n".join(warn_parts) if warn_parts else ""

        return True, warn, all_products

    except Exception as e:
        return False, str(e), []
    finally:
        if driver:
            try:
                # 用執行緒限制 driver.quit() 最多 15 秒，避免 Chrome 關閉卡住
                import threading as _thr
                quit_thread = _thr.Thread(target=driver.quit, daemon=True)
                quit_thread.start()
                quit_thread.join(timeout=15)
            except Exception:
                pass


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/test_config")
def api_test_config():
    """讀取測試設定檔，回傳產品清單供前端自動帶入。"""
    filename = request.args.get("file", "test_products.json")
    # 安全檢查：只允許當前目錄下的 json 檔
    if "/" in filename or "\\" in filename or ".." in filename:
        return jsonify({"error": "不合法的檔案路徑"}), 400
    import os
    filepath = os.path.join(os.path.dirname(__file__) or ".", filename)
    if not os.path.isfile(filepath):
        return jsonify({"error": f"找不到設定檔: {filename}"}), 404
    with open(filepath, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    return jsonify(cfg)


@app.route("/api/test_report", methods=["POST"])
def api_test_report():
    """接收查詢結果與測試設定，執行驗證並產出 HTML + CSV 報告。"""
    import os
    from test_eprice import validate_product, generate_html_report, generate_csv_report, RATE_MAP

    data = request.get_json(force=True, silent=True) or {}
    config_file = data.get("config_file", "test_products.json")
    query_products = data.get("products")
    username = data.get("username", "unknown")
    elapsed = data.get("elapsed", 0)
    warning = data.get("warning", "")

    # 載入測試設定
    filepath = os.path.join(os.path.dirname(__file__) or ".", config_file)
    if not os.path.isfile(filepath):
        return jsonify({"error": f"找不到設定檔: {config_file}"}), 404
    with open(filepath, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    # 設定匯率
    RATE_MAP["USD"] = cfg.get("usd_rate", 32)
    RATE_MAP["RMB"] = cfg.get("rmb_rate", 4.4)

    # 將查詢結果配對回設定
    all_results = []
    for p_cfg in cfg["products"]:
        matched = [r for r in query_products if r.get("query_name") == p_cfg["name"]]
        result = {
            "config": p_cfg,
            "ok": True,
            "warning": warning,
            "matched_results": matched,
            "elapsed": elapsed / max(len(cfg["products"]), 1),
        }
        result["checks"] = validate_product(result)
        all_results.append(result)

    # 產生報告
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    html_path = generate_html_report(all_results, username, elapsed, timestamp)
    csv_path = generate_csv_report(all_results, timestamp)

    passed = sum(1 for r in all_results if all(c[1] for c in r["checks"]))
    failed = len(all_results) - passed

    return jsonify({
        "total": len(all_results),
        "passed": passed,
        "failed": failed,
        "html_report": html_path,
        "csv_report": csv_path,
    })


@app.route("/api/query_stream", methods=["POST"])
def api_query_stream():
    """SSE 串流版查詢，即時回報進度步驟。"""
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    raw_product = (data.get("product_name") or "").strip()
    period_start = (data.get("period_start") or "").strip() or "202501"

    # 支援換行或分號分隔多個產品
    product_names = [n.strip() for n in re.split(r'[;\n]+', raw_product) if n.strip()]

    if not username or not password or not product_names:
        def err_stream():
            payload = json.dumps({"type": "error", "error": "請填寫 User Name、Password 與 Product Name。"})
            yield f"data: {payload}\n\n"
        return Response(stream_with_context(err_stream()), content_type="text/event-stream")

    import queue, threading
    progress_queue = queue.Queue()

    def on_progress(step_idx, label):
        progress_queue.put(("progress", step_idx, label))

    def worker():
        try:
            ok, error_msg, product_list = fetch_base_price(
                username, password, product_names, period_start=period_start, on_progress=on_progress
            )
            if ok:
                progress_queue.put(("done", {
                    "success": True,
                    "products": product_list,
                    "count": len(product_list),
                    "error": None,
                    "warning": error_msg if error_msg else None,
                }))
            else:
                progress_queue.put(("done", {
                    "success": False,
                    "error": error_msg,
                    "products": [],
                    "count": 0,
                }))
        except Exception as ex:
            progress_queue.put(("done", {
                "success": False,
                "error": str(ex),
                "products": [],
                "count": 0,
            }))

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    def generate():
        total_wait = 0
        max_wait = 300          # 總計最多等 5 分鐘
        poll_interval = 10      # 每 10 秒檢查一次
        while True:
            try:
                msg = progress_queue.get(timeout=poll_interval)
                total_wait = 0  # 收到訊息就重置計時
            except queue.Empty:
                total_wait += poll_interval
                if total_wait >= max_wait:
                    yield f"data: {json.dumps({'type': 'error', 'error': '查詢逾時（超過 5 分鐘無回應）'})}\n\n"
                    return
                # 送出 SSE 心跳，防止連線斷開
                yield ": keepalive\n\n"
                continue
            if msg[0] == "progress":
                payload = json.dumps({"type": "progress", "step": msg[1], "label": msg[2]})
                yield f"data: {payload}\n\n"
            elif msg[0] == "done":
                result = msg[1]
                result["type"] = "result"
                yield f"data: {json.dumps(result)}\n\n"
                return

    return Response(stream_with_context(generate()), content_type="text/event-stream")


@app.route("/api/bom_query_stream", methods=["POST"])
def api_bom_query_stream():
    """SSE 串流版 BOM 查詢，點擊料號後查詢 BOM with cost 中的 ARM/DDR/eMMC 物料。"""
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    pn = (data.get("pn") or "").strip()
    plant = (data.get("plant") or "").strip()

    if not username or not password or not pn or not plant:
        def err_stream():
            payload = json.dumps({"type": "error", "error": "請確認帳號、密碼、料號與工廠皆已填寫。"})
            yield f"data: {payload}\n\n"
        return Response(stream_with_context(err_stream()), content_type="text/event-stream")

    import queue, threading
    progress_queue = queue.Queue()

    def on_progress(step_idx, label):
        progress_queue.put(("progress", step_idx, label))

    def worker():
        try:
            items, error_msg = _fetch_bom_with_cost(
                username, password, pn, plant, on_progress=on_progress
            )
            if error_msg and not items:
                progress_queue.put(("done", {
                    "success": False,
                    "error": error_msg,
                    "pn": pn,
                    "plant": plant,
                    "bom_items": [],
                }))
            else:
                progress_queue.put(("done", {
                    "success": True,
                    "pn": pn,
                    "plant": plant,
                    "bom_items": items,
                    "count": len(items),
                    "warning": error_msg,
                }))
        except Exception as ex:
            progress_queue.put(("done", {
                "success": False,
                "error": str(ex),
                "pn": pn,
                "plant": plant,
                "bom_items": [],
            }))

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    def generate():
        total_wait = 0
        max_wait = 300
        poll_interval = 10
        while True:
            try:
                msg = progress_queue.get(timeout=poll_interval)
                total_wait = 0
            except queue.Empty:
                total_wait += poll_interval
                if total_wait >= max_wait:
                    yield f"data: {json.dumps({'type': 'error', 'error': 'BOM 查詢逾時（超過 5 分鐘無回應）'})}\n\n"
                    return
                yield ": keepalive\n\n"
                continue
            if msg[0] == "progress":
                payload = json.dumps({"type": "progress", "step": msg[1], "label": msg[2]})
                yield f"data: {payload}\n\n"
            elif msg[0] == "done":
                result = msg[1]
                result["type"] = "result"
                yield f"data: {json.dumps(result)}\n\n"
                return

    return Response(stream_with_context(generate()), content_type="text/event-stream")


@app.route("/api/batch_bom_query_stream", methods=["POST"])
def api_batch_bom_query_stream():
    """SSE 串流版批次 BOM 查詢，逐 PN 查 BOM with cost 篩選 71/72 物料。"""
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    raw_pns = (data.get("pn_list") or "").strip()
    plant = (data.get("plant") or "").strip()

    pn_list = [p.strip() for p in re.split(r'[\n;]+', raw_pns) if p.strip()]

    if not username or not password or not pn_list or not plant:
        def err_stream():
            payload = json.dumps({"type": "error", "error": "請確認帳號、密碼、料號清單與工廠皆已填寫。"})
            yield f"data: {payload}\n\n"
        return Response(stream_with_context(err_stream()), content_type="text/event-stream")

    import queue, threading
    progress_queue = queue.Queue()

    def on_progress(step_idx, label):
        progress_queue.put(("progress", step_idx, label))

    def worker():
        try:
            result_map, warn_list = _fetch_batch_bom_with_cost(
                username, password, pn_list, plant, on_progress=on_progress
            )
            # 彙整所有結果
            all_items = []
            for pn in pn_list:
                pn = pn.strip()
                items = result_map.get(pn, [])
                for item in items:
                    item["source_pn"] = pn
                all_items.extend(items)

            progress_queue.put(("done", {
                "success": True,
                "bom_items": all_items,
                "count": len(all_items),
                "total_pn": len(pn_list),
                "matched_pn": len(result_map),
                "plant": plant,
                "warning": "\n".join(warn_list) if warn_list else None,
            }))
        except Exception as ex:
            progress_queue.put(("done", {
                "success": False,
                "error": str(ex),
                "bom_items": [],
                "count": 0,
            }))

    t = threading.Thread(target=worker, daemon=True)
    t.start()

    def generate():
        total_wait = 0
        max_wait = 600          # 批次查詢可能較久，最多等 10 分鐘
        poll_interval = 10
        while True:
            try:
                msg = progress_queue.get(timeout=poll_interval)
                total_wait = 0
            except queue.Empty:
                total_wait += poll_interval
                if total_wait >= max_wait:
                    yield f"data: {json.dumps({'type': 'error', 'error': '批次 BOM 查詢逾時（超過 10 分鐘無回應）'})}\n\n"
                    return
                yield ": keepalive\n\n"
                continue
            if msg[0] == "progress":
                payload = json.dumps({"type": "progress", "step": msg[1], "label": msg[2]})
                yield f"data: {payload}\n\n"
            elif msg[0] == "done":
                result = msg[1]
                result["type"] = "result"
                yield f"data: {json.dumps(result)}\n\n"
                return

    return Response(stream_with_context(generate()), content_type="text/event-stream")


@app.route("/api/query", methods=["POST"])
def api_query():
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    raw_product = (data.get("product_name") or "").strip()
    period_start = (data.get("period_start") or "").strip() or "202501"

    # 支援換行或分號分隔多個產品
    product_names = [n.strip() for n in re.split(r'[;\n]+', raw_product) if n.strip()]

    if not username or not password or not product_names:
        return jsonify({
            "success": False,
            "error": "請填寫 User Name、Password 與 Product Name。",
            "base_price": None,
        }), 400

    ok, error_msg, product_list = fetch_base_price(username, password, product_names, period_start=period_start)
    if ok:
        return jsonify({
            "success": True,
            "products": product_list,
            "count": len(product_list),
            "error": None,
            "warning": error_msg if error_msg else None,
        })
    return jsonify({
        "success": False,
        "error": error_msg,
        "products": [],
        "count": 0,
    }), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
