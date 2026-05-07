"""
RPA Runner — popup watcher runs in background throughout entire session.
Downloads all available invoices across all pages in batches of 10.
"""
import time, threading, logging
from pathlib import Path
from typing import List, Dict
from selenium.webdriver.common.by import By
from rpa.config import Config
from rpa.browser import XPOBrowser
from rpa.report import ReportGenerator


class RPARunner:
    def __init__(self, config: Config, logger: logging.Logger):
        self.config  = config
        self.logger  = logger
        self.browser = XPOBrowser(config.DOWNLOAD_DIR, logger)
        self.results: List[Dict] = []
        self._watching = False

    def run(self):
        try:
            self.logger.info("Step 1/4 — Setting up browser")
            self.browser.setup()

            self.logger.info("Step 2/4 — Logging in")
            self.browser.login(self.config.XPO_URL, self.config.login_id, self.config.password)

            # Start popup watcher background thread
            self._start_watcher()
            time.sleep(2)

            self.logger.info("Step 3/4 — Navigating to Billing")
            self.browser.navigate_to_billing()

            self.logger.info("Step 4/4 — Downloading all invoices")
            self._download_all()

            self.logger.info("Step 5/4 — Generating Excel report")
            report = ReportGenerator(self.config.DOWNLOAD_DIR, self.logger).generate(self.results)
            self._summary(report)

        finally:
            self._stop_watcher()
            self.browser.close()

    # ── Background popup watcher ──────────────────────────────────
    def _start_watcher(self):
        self._watching = True
        threading.Thread(target=self._watch_loop, daemon=True).start()
        self.logger.info("Popup watcher started.")

    def _stop_watcher(self):
        self._watching = False

    def _watch_loop(self):
        while self._watching:
            try:
                self.browser.dismiss_popup()
            except Exception:
                pass
            time.sleep(2)

    # ── Main download loop ────────────────────────────────────────
    def _download_all(self):
        batch_size = self.config.BATCH_SIZE
        batch_num  = 0
        page_num   = 1
        total_ok   = 0

        while True:
            self.logger.info(f"\n=== Page {page_num} ===")
            self.browser.dismiss_popup()
            time.sleep(1)

            rows = self.browser.get_bill_rows()
            if not rows:
                self.logger.info("No rows — stopping.")
                break

            self.logger.info(f"  {len(rows)} invoices on page {page_num}")

            for start in range(0, len(rows), batch_size):
                chunk     = rows[start: start + batch_size]
                batch_num += 1
                self.logger.info(f"\n--- Batch {batch_num} | rows {start+1}–{start+len(chunk)} ---")

                self.browser.dismiss_popup()
                self.browser.uncheck_all()

                batch_results = self.browser.download_batch(chunk, batch_num)
                self.results.extend(batch_results)

                ok      = sum(1 for r in batch_results if r.get("success"))
                total_ok += ok
                self.logger.info(f"    {ok}/{len(chunk)} downloaded | running total: {total_ok}")
                time.sleep(1)

            self.logger.info(f"Page {page_num} done.")
            self.browser.dismiss_popup()
            if not self.browser.next_page():
                self.logger.info("No more pages.")
                break
            page_num += 1
            time.sleep(2)

        self.logger.info(f"\nAll pages done. Total downloaded: {total_ok}")

    def _summary(self, report: Path):
        total = len(self.results)
        ok    = sum(1 for r in self.results if r.get("success"))
        self.logger.info("\n" + "="*55)
        self.logger.info("COMPLETED")
        self.logger.info(f"  Total invoices : {total}")
        self.logger.info(f"  Downloaded     : {ok}")
        self.logger.info(f"  Failed         : {total - ok}")
        self.logger.info(f"  Saved to       : {self.browser.session_download_dir}")
        self.logger.info(f"  Excel report   : {report}")
        self.logger.info("="*55)
