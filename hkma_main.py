import sys
import traceback
import time
import re
import pandas as pd
import os
import ddddocr
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def safe_save_workbook(wb, filename):
    """Wait for user to close Excel if a permission error occurs."""
    while True:
        try:
            wb.save(filename)
            return True
        except PermissionError:
            print(f"\n⚠️  PERMISSION DENIED: Please CLOSE '{filename}' in Excel!")
            time.sleep(5)
        except Exception as e:
            print(f"❌ Save Error: {e}")
            return False

class HKMABot:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        
        service = Service(ChromeDriverManager().install())
        self.ocr = ddddocr.DdddOcr(show_ad=False)
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        self.opened_urls = set()

    def handle_landing_page(self, lang_path):
        search_url = f"https://apps.hkma.gov.hk/{lang_path}/index.php?c=search&m=search_by_name"
        self.driver.get(f"https://apps.hkma.gov.hk/{lang_path}/index.php?")
        try:
            btn = self.wait.until(EC.element_to_be_clickable((By.ID, "submitOK")))
            self.driver.execute_script("arguments[0].click();", btn)
            time.sleep(1)
        except:
            if self.driver.current_url != search_url:
                self.driver.get(search_url)

    def solve_and_search_persistent(self, is_chi, sn, fn):
        while True:
            try:
                # Re-input name fields
                if is_chi:
                    field = self.wait.until(EC.presence_of_element_located((By.ID, "optional_name_ch_name")))
                    field.clear()
                    field.send_keys(str(sn))
                else:
                    s_field = self.wait.until(EC.presence_of_element_located((By.ID, "optional_name_surname")))
                    f_field = self.driver.find_element(By.ID, "optional_name_forename")
                    s_field.clear()
                    s_field.send_keys(str(sn))
                    f_field.clear()
                    f_field.send_keys(str(fn))

                # Solve Captcha
                img = self.wait.until(EC.visibility_of_element_located((By.ID, "captcha_search_by_name")))
                img.screenshot("captcha_current.png")
                with open("captcha_current.png", 'rb') as f:
                    res = self.ocr.classification(f.read()).upper()
                
                c_field = self.driver.find_element(By.ID, "captcha_code_field_by_name")
                c_field.clear()
                c_field.send_keys(res)
                
                # Use JavaScript Click for Submit to avoid interception
                submit_btn = self.driver.find_element(By.CSS_SELECTOR, "input.submit")
                self.driver.execute_script("arguments[0].click();", submit_btn)
                time.sleep(2.5) 
                
                error_msgs = self.driver.find_elements(By.ID, "msg_search_by_name_captchya")
                if not any(m.is_displayed() and ("Incorrect" in m.text or "不正確" in m.text) for m in error_msgs):
                    return True 
                
                print(f"❌ Captcha incorrect. Retrying...")
                # Use JavaScript Click for Refresh to avoid interception
                refresh_btn = self.driver.find_element(By.CSS_SELECTOR, "a[onclick*='captcha']")
                self.driver.execute_script("arguments[0].click();", refresh_btn)
                time.sleep(1.5)
            except Exception as e:
                print(f"⚠️ Navigation Error: {e}. Refreshing page...")
                self.driver.refresh()
                time.sleep(2)

    def normalize_to_set(self, text):
        if not text or pd.isna(text) or str(text).lower() in ["nan", ""]: return set()
        clean = re.sub(r'[^A-Z0-9\u4e00-\u9fff\s]', ' ', str(text).upper())
        return set(clean.split())

    def run_task(self, is_chi, sn, fn, target_parts):
        self.handle_landing_page("chi" if is_chi else "eng")
        status, count = "N", 0
        
        if self.solve_and_search_persistent(is_chi, sn, fn):
            time.sleep(2)
            
            # Extract Count
            try:
                count_element = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Records:') or contains(text(), '項記錄')]")
                numbers = re.findall(r'\d+', count_element.text)
                if numbers:
                    count = int(numbers[-1])
            except:
                count = 0

            # Process Results
            links = self.driver.find_elements(By.CSS_SELECTOR, ".search_result_td a")
            if len(links) > 0:
                status = "NRH"
                for link in links:
                    if self.normalize_to_set(link.text) == target_parts:
                        status = "Y"
                        url = link.get_attribute("href")
                        if url not in self.opened_urls:
                            self.driver.execute_script(f"window.open('{url}', '_blank');")
                            self.opened_urls.add(url)
                            self.driver.switch_to.window(self.driver.window_handles[0])
                            
        return status, count

def main():
    file = "Search_List.xlsx"
    if not os.path.exists(file): return

    wb = openpyxl.load_workbook(file)
    ws1 = wb["Sheet1"]
    df_s1 = pd.read_excel(file, sheet_name="Sheet1").fillna("")
    df_s2 = pd.read_excel(file, sheet_name="Sheet2").fillna("")
    bot = HKMABot()

    print(f"🚀 Running Sequential Search...")

    for idx, row in df_s1.iterrows():
        master_name = str(row.iloc[0]).strip()
        if not master_name: continue
        
        target_parts = bot.normalize_to_set(master_name)
        match_found = False
        
        for _, s2_row in df_s2.iterrows():
            sn2, fn2, cn2 = str(s2_row.iloc[0]).strip(), str(s2_row.iloc[1]).strip(), str(s2_row.iloc[2]).strip()
            
            if bot.normalize_to_set(cn2) == target_parts:
                status, count = bot.run_task(True, cn2, "", target_parts)
                match_found = True
                break
            elif bot.normalize_to_set(f"{fn2} {sn2}") == target_parts:
                status, count = bot.run_task(False, sn2, fn2, target_parts)
                match_found = True
                break
        
        if match_found:
            ws1.cell(row=idx + 2, column=6).value = status
            print(f"✅ [{idx+1}] {master_name} -> {status} ({count} results)")
            safe_save_workbook(wb, file)

    print("\nALL TASKS COMPLETE")
    input("Press Enter to close...")

if __name__ == "__main__":
    try:
        main()
    except Exception:
        traceback.print_exc()
        input("Press Enter to exit...")