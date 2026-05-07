"""
Run: python diagnose.py
Logs in, opens billing, dumps EXACT HTML structure of the invoice table.
Paste the terminal output here to fix row selectors permanently.
"""
import sys, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

from rpa.config import Config
from rpa.logger import setup_logger
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

logger = setup_logger("DIAGNOSE")
config = Config()
config.validate()

opts = Options()
opts.add_argument("--start-maximized")

try:
    driver = webdriver.Chrome(options=opts)
except Exception:
    from webdriver_manager.chrome import ChromeDriverManager
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

try:
    # Login
    driver.get(config.XPO_URL)
    time.sleep(4)
    for sel in [(By.XPATH,"//input[@type='email']"),(By.XPATH,"//input[@type='text']"),(By.ID,"username")]:
        try: f=driver.find_element(*sel); f.clear(); f.send_keys(config.login_id); break
        except: continue
    driver.find_element(By.XPATH,"//input[@type='password']").send_keys(config.password)
    for sel in [(By.XPATH,"//button[@type='submit']"),(By.XPATH,"//input[@type='submit']")]:
        try: driver.find_element(*sel).click(); break
        except: continue
    time.sleep(5)

    # Go to billing
    driver.get("https://ext-web.ltl-xpo.com/app/home/billing")
    print("Waiting 8 seconds for page to load...")
    time.sleep(8)
    print(f"URL: {driver.current_url}\n")

    # Try every row selector
    print("="*60)
    print("TESTING ROW SELECTORS:")
    print("="*60)
    selectors = [
        "//table//tbody/tr",
        "//mat-row",
        "//tr[contains(@class,'mat-row')]",
        "//div[contains(@class,'mat-row')]",
        "//*[@role='row'][not(@role='columnheader')]",
        "//cdk-row",
        "//tr[@mat-row]",
        "//tbody/tr",
    ]
    for sel in selectors:
        els = driver.find_elements(By.XPATH, sel)
        print(f"  {len(els):3d} results — {sel}")

    print("\n" + "="*60)
    print("OUTER HTML of <tbody> (first 2000 chars):")
    print("="*60)
    try:
        tbody = driver.find_element(By.XPATH, "//tbody")
        html  = driver.execute_script("return arguments[0].outerHTML;", tbody)
        print(html[:2000])
    except:
        print("  No <tbody> found!")
        print("\nFull body tag names at depth 3:")
        els = driver.find_elements(By.XPATH, "//*")
        tags = {}
        for e in els:
            try:
                t = e.tag_name.lower()
                tags[t] = tags.get(t, 0) + 1
            except: pass
        for k,v in sorted(tags.items(), key=lambda x: -x[1])[:30]:
            print(f"  <{k}> x{v}")

    print("\n" + "="*60)
    print("ALL BUTTONS text:")
    print("="*60)
    for b in driver.find_elements(By.XPATH, "//button"):
        try:
            t = b.text.strip().replace("\n"," ")
            if t: print(f"  '{t}' | disabled={b.get_attribute('disabled')}")
        except: pass

    input("\nDone. Press ENTER to close...")
finally:
    driver.quit()
