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
            old_fp = driver.execute_script("""
                var rows = document.querySelectorAll('table tbody tr');
                if (rows.length === 0) return '';
                var first = rows[0].textContent.trim();
                return rows.length + ':' + first.substring(0, 80);
            """) or ""
            search_btn.click()

            step = f"{pn_label}: 等待"

            def data_ready(drv):
                return drv.execute_script("""
                    var rows = document.querySelectorAll('table tbody tr');
                    if (rows.length === 0) return false;
                    var first = rows[0].textContent.trim();
                    var fp = rows.length + ':' + first.substring(0, 80);
                    return fp !== arguments[0];
                """, old_fp)

            if old_fp:
                try:
                    WebDriverWait(driver, 15, poll_frequency=0.3).until(data_ready)
                except TimeoutException:
                    diag_parts.append(f"{pn}: 逾時")
                    continue
            else:
                def has_rows(drv):
                    return drv.execute_script(
                        "return document.querySelectorAll('table tbody tr').length > 0"
                    )
                try:
                    WebDriverWait(driver, 15, poll_frequency=0.3).until(has_rows)
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
    return False, before_g


def fetch_base_price(username: str, password: str, product_name: str, period_start: str = "202501", on_progress=None) -> Tuple[bool, str, list]:
    """
    使用 Selenium 登入並查詢，列出所有結果的 ProductNumber/Name/Spec 與 Base Price。
    回傳 (成功與否, 錯誤訊息若失敗, 產品列表 [{product_number_name_spec, base_price}] )
    on_progress: callback(step_index, step_label) 用於回報進度
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

        # 2. 價目頁（先等 Product Name 輸入框出現，不強制等 Grid）
        progress(2, "搜尋產品 Base Price")
        driver.get(PRICE_URL)
        wait.until(
            EC.presence_of_element_located(
                (By.ID, "MainContent_txt_productName")
            )
        )

        # 3. 等待 Product Name 輸入框可互動
        pn_input = wait.until(
            EC.element_to_be_clickable((By.ID, "MainContent_txt_productName"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", pn_input)
        time.sleep(0.1)

        ac = ActionChains(driver)
        ac.move_to_element(pn_input).click()
        ac.key_down(Keys.CONTROL).send_keys("a").key_up(Keys.CONTROL)
        ac.send_keys(product_name)
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
            product_name,
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
            debug_msg = driver.execute_script("""
                var g = document.getElementById('MainContent_gv');
                if (!g) return '無法讀取 Grid';
                var rows = g.querySelectorAll('tbody > tr');
                var info = [];
                for (var i = 0; i < rows.length; i++)
                    info.push('第'+(i+1)+'列:'+rows[i].querySelectorAll('td').length+'欄');
                return 'tbody 共 '+rows.length+' 列。 '+info.join('；');
            """) or "Grid 不存在"
            return False, f"查無產品「{product_name}」的價格資料或查詢逾時。Debug: {debug_msg}", []

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
            return False, "在表頭找不到 'Base Price' 欄位。", []
        if quantity_col_index is None:
            return False, "在表頭找不到 'Quantity' 欄位。", []

        PRODUCT_COL = 1
        product_list = []
        need_cols = max(
            base_col_index, PRODUCT_COL, quantity_col_index,
            currency_col_index if currency_col_index is not None else 0,
        )
        want = product_name.strip()
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
                }
            )

        if not product_list:
            return False, "查無產品價格資料，或結果表無 Base Price 欄位。", []

        # ── 查詢 anydoor SAP Report 取得 Cost(highest) ──
        progress(3, "登入 EIP / Anydoor")
        unique_pns = list({
            p["product_number_name_spec"].split("\n")[0].strip()
            for p in product_list
            if p["product_number_name_spec"].strip()
        })
        cost_highest_map, anydoor_diag = _fetch_cost_highest_from_anydoor(
            driver, unique_pns, username, password, period_start=period_start,
            on_progress=on_progress,
        )
        for p in product_list:
            pn_key = p["product_number_name_spec"].split("\n")[0].strip()
            ch_entries = cost_highest_map.get(pn_key, [])
            p["cost_highest_entries"] = [
                {"cost": cv, "currency": cc} for cv, cc in ch_entries
            ]

        progress(5, "整理結果")

        warn = ""
        if anydoor_diag:
            warn = f" (anydoor 診斷: {anydoor_diag})"

        return True, warn, product_list

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


@app.route("/api/query_stream", methods=["POST"])
def api_query_stream():
    """SSE 串流版查詢，即時回報進度步驟。"""
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    product_name = (data.get("product_name") or "").strip()
    period_start = (data.get("period_start") or "").strip() or "202501"

    if not username or not password or not product_name:
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
                username, password, product_name, period_start=period_start, on_progress=on_progress
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


@app.route("/api/query", methods=["POST"])
def api_query():
    data = request.get_json(force=True, silent=True) or {}
    username = (data.get("username") or "").strip()
    password = (data.get("password") or "").strip()
    product_name = (data.get("product_name") or "").strip()
    period_start = (data.get("period_start") or "").strip() or "202501"

    if not username or not password or not product_name:
        return jsonify({
            "success": False,
            "error": "請填寫 User Name、Password 與 Product Name。",
            "base_price": None,
        }), 400

    ok, error_msg, product_list = fetch_base_price(username, password, product_name, period_start=period_start)
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
