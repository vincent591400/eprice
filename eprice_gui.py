import re
import tkinter as tk
from tkinter import messagebox, ttk
import time
from datetime import datetime
from urllib.parse import quote
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import TimeoutException


LOGIN_URL = "http://epricebook.adlinktech.com/Login.aspx"
PRICE_URL = "http://epricebook.adlinktech.com/Price/PriceList.aspx"
EIP_LOGIN_URL = (
    "https://eip.adlinktech.com/GAIA/Account/Logon"
    "?returnUrl=https%3a%2f%2feip.adlinktech.com%2fGaia%2fportal%2fTPHQ"
)
ANYDOOR_URL = "https://anydoor.adlinktech.com/AnydoorWebNew"

_product_cache = []  # 儲存最近一次查詢結果，供即時重算 E2E 使用

_driver = None
_chromedriver_path = None
_anydoor_iframe_ready = False
_anydoor_tab_handle = None


def _get_or_create_driver():
    """取得現有瀏覽器或建立新的。回傳 (driver, is_new)。"""
    global _driver, _chromedriver_path
    if _driver is not None:
        try:
            _ = _driver.title
            return _driver, False
        except Exception:
            _driver = None
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-extensions')
    if _chromedriver_path is None:
        _chromedriver_path = ChromeDriverManager().install()
    _driver = webdriver.Chrome(
        service=Service(_chromedriver_path),
        options=options,
    )
    return _driver, True


def _close_driver():
    global _driver, _anydoor_iframe_ready, _anydoor_tab_handle
    if _driver:
        try:
            _driver.quit()
        except Exception:
            pass
        _driver = None
    _anydoor_iframe_ready = False
    _anydoor_tab_handle = None


def _parse_numeric(s):
    """從含數字的字串中擷取第一個浮點數，失敗回傳 None。"""
    m = re.search(r"[\d,]+\.?\d*", (s or "").replace(",", ""))
    return float(m.group().replace(",", "")) if m else None


def _extract_currency_from_str(s):
    """從字串中擷取貨幣代碼（如 USD、NTD、RMB）。"""
    m = re.search(r"\b([A-Z]{2,4})\b", (s or ""))
    return m.group(1) if m else ""


def _to_ntd(value, currency, usd_rate, rmb_rate):
    """將任意幣別金額換算為台幣。"""
    cur = (currency or "").strip().upper()
    if cur in ("USD",):
        return value * usd_rate
    if cur in ("RMB", "CNY", "人民幣"):
        return value * rmb_rate
    return value  # NTD / TWD / 未知 → 原值


def _render_results():
    """依 _product_cache 與目前的 Selling Price / 幣別 / 匯率即時渲染結果 Listbox。"""
    selling_price_str = entry_selling.get().strip()
    selling_price = _parse_numeric(selling_price_str)
    sel_currency = var_sell_currency.get()
    usd_rate = _parse_numeric(entry_usd_rate.get().strip()) or 1.0
    rmb_rate = _parse_numeric(entry_rmb_rate.get().strip()) or 1.0

    list_result.delete(0, tk.END)
    if not _product_cache:
        return
    for item in _product_cache:
        pn_ns, currency, bp, cost_lowest = item[0], item[1], item[2], item[3]
        ch_entries = item[4] if len(item) > 4 else []
        cost_display = cost_lowest if cost_lowest else "-"
        cur_display = currency if currency else "-"
        bp_cur_label = f" ({currency})" if currency else ""

        # 將每筆 (cost, currency) 依目前匯率換算 NTD，取最大者
        ch_display = "-"
        best_cost_val = None
        best_cost_cur = ""
        if ch_entries:
            best_ntd = None
            for cv, cc in ch_entries:
                ntd_v = _to_ntd(cv, cc, usd_rate, rmb_rate)
                if best_ntd is None or ntd_v > best_ntd:
                    best_ntd = ntd_v
                    best_cost_val = cv
                    best_cost_cur = cc
            if best_cost_val is not None:
                ch_cur_label = f" ({best_cost_cur})" if best_cost_cur else ""
                ch_display = f"{best_cost_val:.2f}{ch_cur_label}"

        e2e_display = "-"
        if selling_price is not None and selling_price > 0 and best_cost_val is not None:
            sp_ntd = _to_ntd(selling_price, sel_currency, usd_rate, rmb_rate)
            cost_ntd = _to_ntd(best_cost_val, best_cost_cur, usd_rate, rmb_rate)
            if sp_ntd > 0:
                e2e = (sp_ntd - cost_ntd) / sp_ntd * 100
                e2e_display = f"{e2e:.1f}%  [NTD基準]"

        list_result.insert(
            tk.END,
            f"{pn_ns}  |  Currency: {cur_display}  |  Base Price: {bp}{bp_cur_label}"
            f"  |  Cost(Lowest): {cost_display}  |  Cost(Highest): {ch_display}"
            f"  |  E2E: {e2e_display}",
        )
    lbl_result.config(text=f"共 {len(_product_cache)} 筆")


def _fetch_cost_highest_from_anydoor(driver, product_numbers, username, password, period_start="202501"):
    """
    1. 先登入 EIP (eip.adlinktech.com)
    2. 再進 anydoor SAP Report
    3. 每個 PN 依序各查一次（非批次）
    回傳 (dict, diag_msg):
      dict = {pn: [(cost_float, currency_str), ...]}
      diag_msg = 診斷訊息（成功時為 None）
    """
    global _anydoor_iframe_ready, _anydoor_tab_handle
    result = {}
    if not product_numbers:
        return result, "product_numbers 為空"

    step = "初始化"
    try:
        wait = WebDriverWait(driver, 10)
        eprice_handle = driver.current_window_handle


        # ── 切換到 anydoor 分頁（Tab 隔離，保留 iframe context）──
        need_setup = True
        if _anydoor_tab_handle and _anydoor_tab_handle in driver.window_handles:
            driver.switch_to.window(_anydoor_tab_handle)
            if _anydoor_iframe_ready:
                try:
                    iframe = driver.find_element(By.ID, "pageContent")
                    driver.switch_to.frame(iframe)
                    driver.find_element(
                        By.CSS_SELECTOR, "input[name='PeriodStart']"
                    )
                    need_setup = False
                except Exception:
                    _anydoor_iframe_ready = False
                    try:
                        driver.switch_to.default_content()
                    except Exception:
                        pass
        else:
            driver.execute_script("window.open('about:blank','_blank');")
            all_h = driver.window_handles
            _anydoor_tab_handle = [h for h in all_h if h != eprice_handle][0]
            driver.switch_to.window(_anydoor_tab_handle)


        if need_setup:
            # ════════ EIP 登入 ════════
            step = "EIP-Login: 開啟 EIP 登入頁"
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

                step = "EIP-Login: 等待登入完成"
                try:
                    WebDriverWait(driver, 10).until(
                        lambda d: "Account/Logon" not in d.current_url
                    )
                except TimeoutException:
                    err_el = driver.find_elements(By.ID, "error")
                    err_text = err_el[0].text.strip() if err_el else ""
                    return result, (
                        f"EIP 登入失敗: 10 秒後仍在登入頁。"
                        f"URL={driver.current_url}, 錯誤={err_text}"
                    )


            # ════════ 進入 anydoor（URL 嵌入帳密通過 HTTP 認證）════════
            step = "Step1: 開啟 anydoor (HTTP auth)"
            safe_user = quote(username, safe="")
            safe_pass = quote(password, safe="")
            driver.get(
                f"https://{safe_user}:{safe_pass}"
                f"@anydoor.adlinktech.com/AnydoorWebNew"
            )
            time.sleep(2)

            src500 = driver.page_source[:500]
            if "401" in src500 or "Unauthorized" in src500:
                return result, (
                    f"Step1 失敗: anydoor 認證失敗 "
                    f"(URL={driver.current_url})"
                )

            # ── Step 2: 選擇 SAP Report ──
            step = "Step2: 選擇 SAP Report"
            sys_select = wait.until(
                EC.presence_of_element_located((By.ID, "SystemList"))
            )
            time.sleep(0.3)
            options = sys_select.find_elements(By.TAG_NAME, "option")
            opt_texts = [o.text for o in options]
            found_sap = False
            for opt in options:
                if "SAP Report" in opt.text:
                    opt.click()
                    found_sap = True
                    break
            if not found_sap:
                return result, (
                    f"Step2 失敗: SystemList 無 'SAP Report'。"
                    f"可用選項: {opt_texts}"
                )
            time.sleep(0.5)

            # ── Step 3: 點擊 Standard/Actual Cost Report(SAP) ──
            step = "Step3: 點擊報表連結"
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

            # ── Step 4: 切換到 iframe ──
            step = "Step4: 切換 iframe"
            iframe = wait.until(
                EC.presence_of_element_located((By.ID, "pageContent"))
            )
            driver.switch_to.frame(iframe)
            time.sleep(0.3)
            _anydoor_iframe_ready = True

        today_str = datetime.now().strftime("%Y%m")
        diag_parts = []

        # ════════ 逐 PN 查詢 ════════
        for pn_idx, pn in enumerate(product_numbers):
            pn_label = f"PN[{pn_idx+1}/{len(product_numbers)}] {pn}"
            step = f"{pn_label}: 填入表單"


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

            step = f"{pn_label}: 點擊 Search"
            search_btn = driver.find_element(
                By.CSS_SELECTOR,
                "input[type='button'][value='Search']",
            )
            # 記錄搜尋前的資料指紋，用於偵測結果更新
            old_fp = driver.execute_script("""
                var rows = document.querySelectorAll('table tbody tr');
                if (rows.length === 0) return '';
                var first = rows[0].textContent.trim();
                return rows.length + ':' + first.substring(0, 80);
            """) or ""
            search_btn.click()

            step = f"{pn_label}: 等待結果"

            def data_ready(drv):
                return drv.execute_script("""
                    var rows = document.querySelectorAll('table tbody tr');
                    if (rows.length === 0) return false;
                    var first = rows[0].textContent.trim();
                    var fp = rows.length + ':' + first.substring(0, 80);
                    return fp !== arguments[0];
                """, old_fp)

            if old_fp:
                # 已有舊資料 → 等資料指紋改變
                try:
                    WebDriverWait(driver, 15, poll_frequency=0.3).until(data_ready)
                except TimeoutException:
                    diag_parts.append(f"{pn}: 等待逾時無結果")
                    continue
            else:
                # 無舊資料 → 等出現任何列
                def has_rows(drv):
                    return drv.execute_script(
                        "return document.querySelectorAll('table tbody tr').length > 0"
                    )
                try:
                    WebDriverWait(driver, 15, poll_frequency=0.3).until(has_rows)
                except TimeoutException:
                    diag_parts.append(f"{pn}: 等待逾時無結果")
                    continue

            # JS 批次提取：表頭 + 所有資料列
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
            missing = [k for k in needed if k not in col]
            if missing:
                diag_parts.append(
                    f"{pn}: 表頭缺 {missing} (headers={tbl['h']})"
                )
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

            # 遍歷後續分頁
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
                diag_parts.append(
                    f"{pn}: {total_data_rows} 列但無三欄皆有值的資料"
                )


        # 切回 ePrice 分頁
        try:
            driver.switch_to.window(eprice_handle)
        except Exception:
            pass

        if diag_parts and not result:
            return result, "anydoor 全部 PN 查詢失敗:\n" + "\n".join(diag_parts)
        if diag_parts:
            return result, "部分 PN 查詢失敗:\n" + "\n".join(diag_parts)
        return result, None

    except Exception as ex:
        _anydoor_iframe_ready = False
        try:
            driver.switch_to.window(eprice_handle)
        except Exception:
            pass
        return result, f"{step} 例外: {type(ex).__name__}: {ex}"


def _extract_product_name_from_cell(cell_text):
    """
    從 PartNumber/Name/Spec 欄位擷取用於比對的 Product Name：
    - 去除非 9x-xx 開頭的 PartNumber（只保留 9<digit>- 開頭的編號）
    - 取 (G) 或 (EA) 中較早出現之前的內容，再取最後一段作為產品名稱（如 EMU-200）
    回傳 (擷取出的名稱, 該段完整字串)，若無則 ("", "").
    """
    if not cell_text or not cell_text.strip():
        return "", ""
    s = cell_text.strip()
    # 去除非 9x-xx 開頭的「數字開頭」PartNumber，保留產品名如 SDAQ-204
    s = re.sub(r"\b(?!9\d-)\d+-\d+[-A-Za-z0-9]*\b", " ", s)
    s = " ".join(s.split())
    # 取 (G) 或 (EA) 中較早出現之前的內容
    for marker in ["(G)", "(EA)"]:
        if marker in s:
            s = s.split(marker)[0]
            break
    before_g = s.strip()
    tokens = before_g.split()
    extracted = tokens[-1] if tokens else ""
    return extracted, before_g


def run_query():
    global _product_cache, _anydoor_iframe_ready
    username = entry_user.get().strip()
    password = entry_pass.get().strip()
    product_name = entry_product.get().strip()

    if not username or not password or not product_name:
        messagebox.showwarning("提醒", "User Name / Password / Product Name 都要填寫。")
        return

    try:
        btn_run.config(state="disabled")
        lbl_result.config(text="查詢中…")
        root.update()
        t0 = time.time()
        driver, is_new = _get_or_create_driver()
        wait = WebDriverWait(driver, 10)

        def _status(msg):
            lbl_result.config(text=f"{msg} ({time.time()-t0:.1f}s)")
            root.update()

        _status("連線 ePrice…")
        driver.get(PRICE_URL)
        time.sleep(0.3)

        if "Login.aspx" in driver.current_url:
            driver.get(LOGIN_URL)
            user_input = wait.until(
                EC.presence_of_element_located((By.ID, "TB_Username"))
            )
            pass_input = driver.find_element(By.ID, "TB_Password")
            login_button = driver.find_element(By.ID, "Btn_OK")

            user_input.clear()
            user_input.send_keys(username)
            pass_input.clear()
            pass_input.send_keys(password)
            login_button.click()

            wait.until(lambda drv: "Login.aspx" not in drv.current_url)
            driver.get(PRICE_URL)

        wait.until(EC.presence_of_element_located(
            (By.ID, "MainContent_txt_productName")
        ))

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
            "if(inp){ inp.value = v; inp.setAttribute('value', v);"
            "inp.dispatchEvent(new Event('input',{bubbles:true})); "
            "inp.dispatchEvent(new Event('change',{bubbles:true})); }",
            product_name,
        )
        time.sleep(0.2)

        search_btn = wait.until(
            EC.element_to_be_clickable((By.ID, "MainContent_btn_search"))
        )
        _status("搜尋 ePrice…")
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
                for (var i = 0; i < rows.length; i++) {
                    info.push('第'+(i+1)+'列:'+rows[i].querySelectorAll('td').length+'欄');
                }
                return 'tbody 共 '+rows.length+' 列。 '+info.join('；');
            """) or "Grid 不存在"
            raise Exception("查無產品價格資料或查詢逾時。Debug: " + debug_msg)

        # 用 JS 一次性提取表頭 + 所有資料列（含 tooltip）
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
            raise Exception("在表頭找不到 'Base Price' 欄位，請確認欄位名稱後調整關鍵字。")
        if quantity_col_index is None:
            raise Exception("在表頭找不到 'Quantity' 欄位，請確認欄位名稱後調整關鍵字。")

        PRODUCT_COL = 1
        product_list = []
        need_cols = max(
            base_col_index, PRODUCT_COL, quantity_col_index,
            currency_col_index if currency_col_index is not None else 0,
        )
        for row_data in raw["rows"]:
            if len(row_data) <= need_cols:
                continue
            qty_text = row_data[quantity_col_index]["t"]
            if qty_text != "1":
                continue
            try:
                pn_ns = row_data[PRODUCT_COL]["t"]
                extracted_name, before_g = _extract_product_name_from_cell(pn_ns)
                if extracted_name != product_name:
                    continue
                prefix = before_g.rstrip()[:-len(product_name)].rstrip() if before_g.rstrip().endswith(product_name) else ""
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

                product_list.append((pn_ns, currency, bp, cost_lowest))
            except Exception:
                continue


        # ── 查詢 anydoor SAP Report 取得 Cost(highest) ──
        unique_pns = list({
            pn_ns.split("\n")[0].strip()
            for pn_ns, *_ in product_list
            if pn_ns.strip()
        })
        cost_highest_map = {}
        anydoor_diag = None
        if unique_pns:
            _status(f"查詢 anydoor ({len(unique_pns)} PNs)…")
            period_start = entry_period_start.get().strip() or "202501"
            cost_highest_map, anydoor_diag = _fetch_cost_highest_from_anydoor(
                driver, unique_pns, username, password, period_start=period_start
            )
            if anydoor_diag:
                messagebox.showwarning(
                    "anydoor 診斷",
                    f"Cost(highest) 查詢未能取得完整資料:\n\n{anydoor_diag}",
                )

        # 將 Cost(highest) entries 加入 tuple
        # → (pn_ns, currency, bp, cost_lowest, ch_entries)
        # ch_entries = [(cost_float, currency_str), ...] — render 時依匯率換算取最大
        product_list_ext = []
        for pn_ns, currency, bp, cost_lowest in product_list:
            pn_key = pn_ns.split("\n")[0].strip()
            ch_entries = cost_highest_map.get(pn_key, [])
            product_list_ext.append(
                (pn_ns, currency, bp, cost_lowest, ch_entries)
            )

        elapsed = time.time() - t0
        if not product_list_ext:
            _product_cache = []
            list_result.delete(0, tk.END)
            list_result.insert(tk.END, "查無產品價格資料。")
            lbl_result.config(text=f"共 0 筆 ({elapsed:.1f}s)")
        else:
            _product_cache = product_list_ext
            _render_results()
            lbl_result.config(
                text=f"共 {len(_product_cache)} 筆 ({elapsed:.1f}s)"
            )

    except Exception as e:
        _anydoor_iframe_ready = False
        messagebox.showerror("錯誤", f"執行時發生錯誤：\n{e}")
    finally:
        btn_run.config(state="normal")


# ==== Tkinter 視窗 ====
root = tk.Tk()
root.title("ePrice 查詢工具")

# User Name
tk.Label(root, text="User Name:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
entry_user = tk.Entry(root, width=30)
entry_user.grid(row=0, column=1, padx=5, pady=5)

# Password
tk.Label(root, text="Password:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
entry_pass = tk.Entry(root, width=30, show="*")
entry_pass.grid(row=1, column=1, padx=5, pady=5)

# Product Name
tk.Label(root, text="Product Name:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
entry_product = tk.Entry(root, width=30)
entry_product.grid(row=2, column=1, padx=5, pady=5)

# Selling Price（選填，查詢後可即時修改以重算 E2E%）
tk.Label(root, text="Selling Price:").grid(row=3, column=0, padx=5, pady=5, sticky="e")
sell_frame = tk.Frame(root)
sell_frame.grid(row=3, column=1, padx=5, pady=5, sticky="w")
entry_selling = tk.Entry(sell_frame, width=16)
entry_selling.pack(side=tk.LEFT)
var_sell_currency = tk.StringVar(value="USD")
currency_menu = ttk.Combobox(
    sell_frame, textvariable=var_sell_currency,
    values=["USD", "NTD", "RMB"], width=6, state="readonly",
)
currency_menu.pack(side=tk.LEFT, padx=(4, 0))
tk.Label(root, text="(選填；修改數字/幣別/匯率即時重算 E2E%，無需重新查詢)", fg="gray").grid(
    row=4, column=0, columnspan=2, padx=5, pady=0, sticky="w"
)

# 匯率設定
rate_frame = tk.Frame(root)
rate_frame.grid(row=5, column=0, columnspan=2, padx=5, pady=4, sticky="w")
tk.Label(rate_frame, text="匯率：1 USD =").pack(side=tk.LEFT)
entry_usd_rate = tk.Entry(rate_frame, width=7)
entry_usd_rate.insert(0, "32")
entry_usd_rate.pack(side=tk.LEFT, padx=(2, 8))
tk.Label(rate_frame, text="NTD　　1 RMB =").pack(side=tk.LEFT)
entry_rmb_rate = tk.Entry(rate_frame, width=7)
entry_rmb_rate.insert(0, "4.4")
entry_rmb_rate.pack(side=tk.LEFT, padx=(2, 4))
tk.Label(rate_frame, text="NTD　（E2E% 統一以台幣換算）", fg="gray").pack(side=tk.LEFT)

# Cost(Highest) 起始時間
tk.Label(root, text="Cost(Highest) 起始時間:").grid(row=6, column=0, padx=5, pady=5, sticky="e")
period_frame = tk.Frame(root)
period_frame.grid(row=6, column=1, padx=5, pady=5, sticky="w")
entry_period_start = tk.Entry(period_frame, width=10)
entry_period_start.insert(0, "202501")
entry_period_start.pack(side=tk.LEFT)
tk.Label(period_frame, text="（格式 YYYYMM，例如 202501）", fg="gray").pack(side=tk.LEFT, padx=(4, 0))

# 執行按鈕
btn_run = tk.Button(root, text="查詢 Base Price", command=run_query)
btn_run.grid(row=7, column=0, columnspan=2, padx=5, pady=10)

# 結果：筆數
lbl_result = tk.Label(root, text="共 0 筆", fg="blue")
lbl_result.grid(row=8, column=0, columnspan=2, padx=5, pady=2)

# 結果列表
tk.Label(
    root,
    text="Product Number / Name / Spec | Currency | Base Price | Cost(Lowest) | Cost(Highest) | E2E%：",
).grid(row=9, column=0, columnspan=2, padx=5, pady=2, sticky="w")
list_frame = tk.Frame(root)
list_frame.grid(row=10, column=0, columnspan=2, padx=5, pady=5, sticky="nsew")
scrollbar = tk.Scrollbar(list_frame)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
list_result = tk.Listbox(list_frame, height=12, width=100, yscrollcommand=scrollbar.set, font=("Consolas", 9))
list_result.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
scrollbar.config(command=list_result.yview)

# 水平捲軸
scrollbar_h = tk.Scrollbar(root, orient=tk.HORIZONTAL, command=list_result.xview)
scrollbar_h.grid(row=11, column=0, columnspan=2, sticky="ew", padx=5)
list_result.config(xscrollcommand=scrollbar_h.set)
root.columnconfigure(1, weight=1)
root.rowconfigure(10, weight=1)

# 即時重算：Selling Price / 幣別 / 匯率改變時，立即更新 E2E（不重新查詢）
entry_selling.bind("<KeyRelease>", lambda e: _render_results())
var_sell_currency.trace_add("write", lambda *_: _render_results())
entry_usd_rate.bind("<KeyRelease>", lambda e: _render_results())
entry_rmb_rate.bind("<KeyRelease>", lambda e: _render_results())

def _on_closing():
    _close_driver()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", _on_closing)
root.mainloop()