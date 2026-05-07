"""
XPO Browser Automation - Final Version
Confirmed from DevTools:
  Popup rows : li.invoice-item
  Checkbox   : span.mat-checkbox-inner-container
  Bulk button: button#download-invoice  (top bar)
  Confirm dlg: mat-dialog-container > button.mat-primary  (red Download btn)
  Row icon   : button#download-invoice  (inside each li row)
"""

import time
from datetime import datetime
from pathlib import Path
from typing import List, Dict
import logging

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, StaleElementReferenceException
)


class XPOBrowser:

    ELEMENT_TIMEOUT  = 30
    DOWNLOAD_TIMEOUT = 90   # seconds to wait for file

    def __init__(self, download_dir: Path, logger: logging.Logger):
        self.download_dir         = download_dir
        self.logger               = logger
        self.driver               = None
        self.wait                 = None
        self.session_download_dir = None

    # ─────────────────────────────────────── SETUP
    def setup(self):
        date_str = datetime.now().strftime("%Y-%m-%d")
        self.session_download_dir = self.download_dir / date_str
        self.session_download_dir.mkdir(parents=True, exist_ok=True)
        self.logger.info(f"Download folder: {self.session_download_dir}")

        opts = Options()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-notifications")
        opts.add_argument("--no-sandbox")
        opts.add_argument("--disable-dev-shm-usage")
        opts.add_experimental_option("prefs", {
            "download.default_directory":         str(self.session_download_dir.resolve()),
            "download.prompt_for_download":       False,
            "download.directory_upgrade":         True,
            "plugins.always_open_pdf_externally": True,
            "safebrowsing.enabled":               True,
        })

        try:
            self.driver = webdriver.Chrome(options=opts)
        except Exception:
            from webdriver_manager.chrome import ChromeDriverManager
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()), options=opts)

        self.driver.set_page_load_timeout(30)
        self.wait = WebDriverWait(self.driver, self.ELEMENT_TIMEOUT)
        self.logger.info("Browser ready.")

    # ─────────────────────────────────────── LOGIN
    def login(self, url: str, login_id: str, password: str):
        self.logger.info(f"Opening: {url}")
        self.driver.get(url)
        time.sleep(4)

        # Username
        for sel in [(By.XPATH, "//input[@type='email']"),
                    (By.XPATH, "//input[@type='text']"),
                    (By.ID, "username"), (By.NAME, "username")]:
            try:
                f = self.wait.until(EC.element_to_be_clickable(sel))
                f.clear(); f.send_keys(login_id)
                self.logger.info("Username entered.")
                break
            except TimeoutException:
                continue

        # Password
        pw = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='password']")))
        pw.clear(); pw.send_keys(password)
        self.logger.info("Password entered.")

        # Submit
        for sel in [(By.XPATH, "//button[@type='submit']"),
                    (By.XPATH, "//input[@type='submit']")]:
            try:
                self.driver.find_element(*sel).click()
                self.logger.info("Login submitted.")
                break
            except NoSuchElementException:
                continue

        time.sleep(5)
        self.logger.info(f"Logged in. URL: {self.driver.current_url}")

    # ─────────────────────────────────────── DISMISS ANY POPUP
    def dismiss_popup(self):
        """Close any visible modal/popup. Safe to call any time."""
        selectors = [
            "//button[normalize-space(.)='Not Now']",
            "//button[contains(normalize-space(.),'Not Now')]",
            "//button[normalize-space(.)='Close']",
            "//button[normalize-space(.)='Cancel']",
            "//button[normalize-space(.)='Skip']",
            "//button[normalize-space(.)='Later']",
            "//button[normalize-space(.)='No Thanks']",
            "//button[@aria-label='Close']",
            "//button[@aria-label='close']",
            "//mat-dialog-container//button[contains(@class,'mat-icon-button')]",
        ]
        for xpath in selectors:
            try:
                els = self.driver.find_elements(By.XPATH, xpath)
                for el in els:
                    if el.is_displayed() and el.is_enabled():
                        txt = el.text.strip()
                        self.driver.execute_script("arguments[0].click();", el)
                        self.logger.info(f"  [popup] Dismissed: '{txt}'")
                        time.sleep(1)
                        return
            except Exception:
                continue

    # ─────────────────────────────────────── NAVIGATE TO BILLING
    def navigate_to_billing(self):
        self.logger.info("Navigating to Billing page...")
        self.dismiss_popup()

        # Try nav link first
        for sel in [(By.XPATH, "//a[normalize-space(text())='Billing']"),
                    (By.XPATH, "//nav//a[contains(.,'Billing')]"),
                    (By.XPATH, "//a[contains(@href,'billing')]")]:
            try:
                el = self.wait.until(EC.element_to_be_clickable(sel))
                self.driver.execute_script("arguments[0].click();", el)
                time.sleep(2)
                break
            except (TimeoutException, StaleElementReferenceException):
                continue

        # Always go directly to billing URL
        self.driver.get("https://ext-web.ltl-xpo.com/app/billing")
        self.logger.info("Waiting for invoices to load...")
        self._wait_for_rows()
        self.dismiss_popup()
        self.logger.info(f"Billing ready. URL: {self.driver.current_url}")

    # ─────────────────────────────────────── WAIT FOR li.invoice-item
    def _wait_for_rows(self):
        for i in range(20):  # up to 40s
            try:
                rows = self.driver.find_elements(By.CSS_SELECTOR, "li.invoice-item")
                if rows:
                    self.logger.info(f"  {len(rows)} rows loaded after {i*2}s")
                    time.sleep(1)
                    return
            except Exception:
                pass
            time.sleep(2)
        self.logger.warning("Rows not confirmed after 40s — proceeding.")

    # ─────────────────────────────────────── GET ALL ROWS ON PAGE
    def get_bill_rows(self) -> List[Dict]:
        self.logger.info("Reading invoice rows...")

        # Confirmed selector from DevTools
        rows = self.driver.find_elements(By.CSS_SELECTOR, "li.invoice-item")
        if not rows:
            for sel in ["//li[contains(@class,'invoice-item')]",
                        "//li[contains(@class,'invoice')]",
                        "//*[contains(@class,'invoice-item')]"]:
                rows = self.driver.find_elements(By.XPATH, sel)
                if rows:
                    break

        if not rows:
            self.logger.error(f"NO ROWS FOUND. URL={self.driver.current_url} Title={self.driver.title}")
            return []

        self.logger.info(f"Found {len(rows)} rows.")
        bills = []
        for idx, row in enumerate(rows):
            try:
                lines = [l.strip() for l in row.text.split('\n') if l.strip()]
                bills.append({
                    "row_index":   idx,
                    "element":     row,
                    "bill_number": self._get_invoice_num(lines),
                    "bill_date":   self._get_date(lines),
                    "raw_text":    " | ".join(lines[:6]),
                })
            except StaleElementReferenceException:
                continue
        self.logger.info(f"Parsed {len(bills)} bills.")
        return bills

    # ─────────────────────────────────────── DOWNLOAD BATCH OF 10
    def download_batch(self, batch: List[Dict], batch_num: int) -> List[Dict]:
        now          = datetime.now()
        files_before = set(self.session_download_dir.glob("*.*"))
        checked      = 0

        # ── Step 1: tick checkboxes
        self.logger.info(f"Batch {batch_num}: ticking {len(batch)} checkboxes...")
        for bill in batch:
            try:
                row = bill["element"]
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                time.sleep(0.2)
                if self._tick_checkbox(row):
                    checked += 1
                    self.logger.info(f"  ✓ {bill['bill_number']}")
                else:
                    self.logger.warning(f"  ✗ No checkbox: {bill['bill_number']}")
            except StaleElementReferenceException:
                self.logger.warning(f"  Stale: {bill['bill_number']}")
            except Exception as e:
                self.logger.warning(f"  Error {bill['bill_number']}: {e}")

        self.logger.info(f"  {checked}/{len(batch)} checked.")

        if checked == 0:
            self.logger.error("  Nothing checked — skipping batch.")
            return self._failed_results(batch, batch_num, now)

        # ── Step 2: click "Download Invoices" top button
        if not self._click_top_download_button():
            self.logger.warning("  Top button not found — trying per-row icons.")
            return self._download_per_row(batch, batch_num)

        # ── Step 3: click "Download" in confirmation dialog
        self._click_confirm_download()

        # ── Step 4: wait for file(s)
        new_files = self._wait_for_download(files_before)
        self.logger.info(f"  {len(new_files)} file(s) received.")

        results = []
        for i, bill in enumerate(batch):
            fp = new_files[i] if i < len(new_files) else None
            results.append({**bill,
                "success":           fp is not None,
                "file_path":         fp or "",
                "file_name":         Path(fp).name if fp else "",
                "downloaded_on":     now.strftime("%Y-%m-%d"),
                "downloaded_at":     now.strftime("%H:%M:%S"),
                "batch_num":         batch_num,
                "position_in_batch": i + 1,
            })
        return results

    # ─────────────────────────────────────── TICK CHECKBOX
    def _tick_checkbox(self, row) -> bool:
        # JS approach — tries all known checkbox patterns
        ok = self.driver.execute_script("""
            var row = arguments[0];
            var t = row.querySelector('span.mat-checkbox-inner-container') ||
                    row.querySelector('mat-checkbox label') ||
                    row.querySelector('mat-checkbox') ||
                    row.querySelector('[role="checkbox"]') ||
                    row.querySelector('input[type="checkbox"]');
            if (t) { t.click(); return true; }
            return false;
        """, row)
        if ok:
            time.sleep(0.15)
            return True
        # Selenium fallback
        for css in ["span.mat-checkbox-inner-container", "mat-checkbox label",
                    "mat-checkbox", "input[type='checkbox']"]:
            try:
                el = row.find_element(By.CSS_SELECTOR, css)
                self.driver.execute_script("arguments[0].click();", el)
                time.sleep(0.15)
                return True
            except (NoSuchElementException, StaleElementReferenceException):
                continue
        return False

    # ─────────────────────────────────────── CLICK TOP "Download Invoices"
    def _click_top_download_button(self) -> bool:
        time.sleep(1)  # let button become enabled after checkboxes ticked
        for sel in [
            (By.CSS_SELECTOR, "button#download-invoice"),
            (By.XPATH, "//button[@id='download-invoice']"),
            (By.XPATH, "//button[.//span[contains(@class,'mat-button-wrapper') and contains(normalize-space(.),'Download Invoices')]]"),
            (By.XPATH, "//button[contains(normalize-space(.),'Download Invoices')]"),
        ]:
            try:
                btn = WebDriverWait(self.driver, 6).until(EC.presence_of_element_located(sel))
                if btn.get_attribute("disabled"):
                    self.logger.warning("  'Download Invoices' button is disabled — checkboxes may not have registered.")
                    continue
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                time.sleep(0.3)
                self.driver.execute_script("arguments[0].click();", btn)
                self.logger.info("  ✓ Clicked 'Download Invoices'.")
                return True
            except (TimeoutException, StaleElementReferenceException):
                continue
            except Exception as e:
                self.logger.debug(f"  Btn selector failed: {e}")

        # Debug all visible buttons
        self.logger.warning("  'Download Invoices' NOT found. All visible buttons:")
        for b in self.driver.find_elements(By.XPATH, "//button"):
            try:
                if b.is_displayed():
                    self.logger.warning(f"    id='{b.get_attribute('id')}' text='{b.text.strip()[:40]}' disabled={b.get_attribute('disabled')}")
            except Exception:
                pass
        return False

    # ─────────────────────────────────────── CLICK CONFIRM "Download"
    def _click_confirm_download(self):
        """
        After clicking Download Invoices, XPO shows a dialog:
        'Do you wish to download selected 10 invoices as pdf?'
        with [Cancel] [Download] buttons.
        Click the red Download button.
        """
        self.logger.info("  Waiting for confirmation dialog...")
        time.sleep(2)

        for sel in [
            (By.CSS_SELECTOR, "mat-dialog-container button.mat-primary"),
            (By.CSS_SELECTOR, "mat-dialog-container button.mat-flat-button"),
            (By.XPATH, "//mat-dialog-container//button[normalize-space(.)='Download']"),
            (By.XPATH, "//mat-dialog-container//button[contains(normalize-space(.),'Download') and not(contains(.,'Invoices'))]"),
            (By.XPATH, "//button[contains(@class,'mat-primary') and normalize-space(.)='Download']"),
            (By.XPATH, "//button[contains(@class,'mat-flat-button') and normalize-space(.)='Download']"),
        ]:
            try:
                btn = WebDriverWait(self.driver, 6).until(EC.element_to_be_clickable(sel))
                txt = btn.text.strip()
                self.driver.execute_script("arguments[0].click();", btn)
                self.logger.info(f"  ✓ Confirmed dialog — clicked '{txt}'.")
                time.sleep(1)
                return
            except TimeoutException:
                continue
            except Exception as e:
                self.logger.debug(f"  Confirm selector failed: {e}")

        self.logger.warning("  Confirmation dialog not found — may have auto-started.")

    # ─────────────────────────────────────── WAIT FOR DOWNLOAD FILE
    def _get_all_watch_folders(self) -> List[Path]:
        """Return all folders to watch for new downloads."""
        import os
        folders = [self.session_download_dir]
        # Also watch Windows default Downloads folder
        for candidate in [
            Path.home() / "Downloads",
            Path("C:/Users") / os.environ.get("USERNAME", "") / "Downloads",
            Path(os.environ.get("USERPROFILE", "")) / "Downloads",
        ]:
            try:
                if candidate.exists() and candidate not in folders:
                    folders.append(candidate)
            except Exception:
                pass
        self.logger.info(f"  Watching folders: {[str(f) for f in folders]}")
        return folders

    def _wait_for_download(self, files_before: set, timeout: int = 90) -> List[str]:
        self.logger.info(f"  Waiting for download...")
        start            = time.time()
        original_handles = set(self.driver.window_handles)

        # Snapshot all watch folders before download
        watch_folders  = self._get_all_watch_folders()
        snapshots      = {}
        for folder in watch_folders:
            try:
                snapshots[folder] = set(folder.glob("*.*"))
            except Exception:
                snapshots[folder] = set()

        while time.time() - start < timeout:
            time.sleep(2)

            # Check ALL watch folders for new complete files
            for folder in watch_folders:
                try:
                    current  = set(folder.glob("*.*"))
                    before   = snapshots[folder]
                    new_done = [f for f in (current - before)
                                if not str(f).endswith((".crdownload", ".tmp", ".part"))]
                    if new_done:
                        # Move to our session folder if found elsewhere
                        final_paths = []
                        for src in new_done:
                            src = Path(src)
                            if src.parent != self.session_download_dir:
                                dst = self.session_download_dir / src.name
                                import shutil
                                shutil.move(str(src), str(dst))
                                self.logger.info(f"  Moved: {src.name} → session folder")
                                final_paths.append(str(dst))
                            else:
                                final_paths.append(str(src))
                        self.logger.info(f"  File(s) received: {[Path(p).name for p in final_paths]}")
                        time.sleep(1)
                        return final_paths
                except Exception as e:
                    self.logger.debug(f"  Folder check error: {e}")

            # Handle new tab (PDF opened in browser tab)
            new_tabs = set(self.driver.window_handles) - original_handles
            for handle in new_tabs:
                try:
                    self.driver.switch_to.window(handle)
                    url = self.driver.current_url
                    self.logger.info(f"  New tab detected: {url[:80]}")
                    # Trigger JS download
                    self.driver.execute_script("""
                        var a = document.createElement('a');
                        a.href = window.location.href;
                        a.download = 'invoice.pdf';
                        document.body.appendChild(a);
                        a.click();
                    """)
                    time.sleep(3)
                    self.driver.close()
                    self.driver.switch_to.window(list(self.driver.window_handles)[0])
                except Exception as e:
                    self.logger.debug(f"  Tab error: {e}")
                    try:
                        self.driver.switch_to.window(list(self.driver.window_handles)[0])
                    except Exception:
                        pass

        elapsed = int(time.time() - start)
        self.logger.warning(f"  Timed out after {elapsed}s.")
        for folder in watch_folders:
            try:
                files = list(folder.glob("*.*"))
                self.logger.warning(f"  {folder}: {[f.name for f in files[-5:]]}")
            except Exception:
                pass
        return []

    # ─────────────────────────────────────── FALLBACK: INDIVIDUAL ICONS
    def _download_per_row(self, batch: List[Dict], batch_num: int) -> List[Dict]:
        """Fallback: click the download icon in each row individually."""
        results = []
        for pos, bill in enumerate(batch, 1):
            now          = datetime.now()
            files_before = set(self.session_download_dir.glob("*.*"))
            fp           = None
            try:
                row = bill["element"]
                self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", row)
                time.sleep(0.3)

                # Find download icon button inside the row
                dl_btn = self.driver.execute_script("""
                    var row = arguments[0];
                    var b = row.querySelector('button#download-invoice');
                    if (b) return b;
                    var all = row.querySelectorAll('button');
                    for (var el of all) {
                        var id  = (el.id || '').toLowerCase();
                        var cls = (el.className || '').toLowerCase();
                        var htm = el.innerHTML.toLowerCase();
                        if (id.includes('download') || cls.includes('download') ||
                            htm.includes('cloud_download') || htm.includes('save_alt'))
                            return el;
                    }
                    return null;
                """, row)

                if dl_btn:
                    self.driver.execute_script("arguments[0].click();", dl_btn)
                    self.logger.info(f"  ✓ Icon clicked: {bill['bill_number']}")
                    self._click_confirm_download()
                    files = self._wait_for_download(files_before, timeout=30)
                    fp    = files[0] if files else None
                else:
                    self.logger.warning(f"  No icon: {bill['bill_number']}")

            except StaleElementReferenceException:
                self.logger.warning(f"  Stale: {bill['bill_number']}")
            except Exception as e:
                self.logger.error(f"  Error {bill['bill_number']}: {e}")

            results.append({**bill,
                "success":           fp is not None,
                "file_path":         fp or "",
                "file_name":         Path(fp).name if fp else "",
                "downloaded_on":     now.strftime("%Y-%m-%d"),
                "downloaded_at":     now.strftime("%H:%M:%S"),
                "batch_num":         batch_num,
                "position_in_batch": pos,
            })
        return results

    # ─────────────────────────────────────── NEXT PAGE
    def next_page(self) -> bool:
        for sel in [
            (By.XPATH,        "//button[@aria-label='Next page']"),
            (By.XPATH,        "//button[@aria-label='Next']"),
            (By.XPATH,        "//button[@aria-label='next']"),
            (By.CSS_SELECTOR, "button[aria-label='Next page']:not([disabled])"),
            (By.XPATH,        "//div[contains(@class,'invoice-pag')]//button[last()]"),
            (By.XPATH,        "//div[contains(@class,'paginator')]//button[last()]"),
        ]:
            try:
                btn = WebDriverWait(self.driver, 4).until(EC.element_to_be_clickable(sel))
                if btn.get_attribute("disabled"):
                    return False
                self.driver.execute_script("arguments[0].click();", btn)
                self.logger.info("  ✓ Next page.")
                time.sleep(2)
                self._wait_for_rows()
                return True
            except (TimeoutException, NoSuchElementException, StaleElementReferenceException):
                continue
        return False

    # ─────────────────────────────────────── HELPERS
    def uncheck_all(self):
        try:
            self.driver.execute_script("""
                document.querySelectorAll('input[type="checkbox"]:checked')
                    .forEach(cb => { if (!cb.indeterminate) cb.click(); });
            """)
            time.sleep(0.3)
        except Exception:
            pass

    def _failed_results(self, batch, batch_num, now):
        return [{**bill, "success": False, "file_path": "", "file_name": "",
                 "downloaded_on": now.strftime("%Y-%m-%d"), "downloaded_at": now.strftime("%H:%M:%S"),
                 "batch_num": batch_num, "position_in_batch": i+1}
                for i, bill in enumerate(batch)]

    def _get_invoice_num(self, lines):
        import re
        for t in lines:
            m = re.search(r'\b(\d{9,})\b', t)
            if m: return m.group(1)
            m = re.search(r'INVOICE\s+(\S+)', t, re.IGNORECASE)
            if m: return m.group(1)
        return lines[0][:30] if lines else "UNKNOWN"

    def _get_date(self, lines):
        import re
        for t in lines:
            m = re.search(r'\b(\d{2}/\d{2}/\d{4})\b', t)
            if m: return m.group(1)
        return datetime.now().strftime("%m/%d/%Y")

    def close(self):
        if self.driver:
            try: self.driver.quit()
            except Exception: pass
            self.logger.info("Browser closed.")
