import time
import re
import unicodedata
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

class HKSFCSearcher: 
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--incognito")
        
        prefs = {"profile.managed_default_content_settings.images": 2}
        chrome_options.add_experimental_option("prefs", prefs)

        self.driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )
        self.wait = WebDriverWait(self.driver, 15)
        self.opened_urls = set()

    def is_chinese(self, val):
        """Checks if the string contains Chinese characters."""
        return any('\u4e00' <= char <= '\u9fff' for char in str(val))

    def normalize_english(self, s):
        s = re.sub(r"[^a-zA-Z0-9\s]", " ", str(s))
        s = re.sub(r"\s+", " ", s).strip().lower()
        return " ".join(sorted(s.split()))

    def normalize_chinese(self, s):
        if not s or pd.isna(s): return ""
        s = unicodedata.normalize("NFKC", str(s))
        s = re.sub(r'[\s\u3000\xa0]+', '', s)
        return s

    def smart_click(self, label_text):
        try:
            xpath = f"//label[contains(text(), '{label_text}')]/preceding-sibling::input | //label[contains(text(), '{label_text}')]"
            element = self.wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
            self.driver.execute_script("arguments[0].click();", element)
            time.sleep(0.3) 
        except: pass

    def perform_search(self, query_name, all_variants):
        try:
            is_chi_query = self.is_chinese(query_name)
            print(f"\n🔍 HKSFC Searching ({'Chinese' if is_chi_query else 'English'}): {query_name}")
            
            # Use the same tab for every search
            self.driver.get("https://apps.sfc.hk/publicregWeb/searchByName?locale=en")

            # Navigation
            self.smart_click("Active and inactive")
            self.smart_click("SFO licence and/or AMLO licence")
            self.smart_click("Individual")
            
            if is_chi_query:
                self.smart_click("Chinese name")
            else:
                self.smart_click("English name")

            name_input = self.wait.until(EC.element_to_be_clickable((By.ID, "searchtextname-inputEl")))
            name_input.clear()
            name_input.send_keys(query_name)

            search_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Search']/ancestor::button")))
            self.driver.execute_script("arguments[0].click();", search_btn)

            # CRITICAL FIX: Wait for the AJAX grid to actually update
            time.sleep(1.5) 

            try:
                self.wait.until(EC.presence_of_element_located((By.XPATH, "//tr[contains(@class,'x-grid-row')] | //div[contains(text(), 'no name matched')]")))
            except: pass 

            total_hits = 0
            try:
                # Check for the items count in the toolbar
                toolbar_text = self.driver.find_element(By.XPATH, "//*[contains(text(), 'Item 1 -') or contains(text(), 'Item 0 -')]").text
                match = re.search(r"of (\d+)", toolbar_text)
                if match: total_hits = int(match.group(1))
            except:
                if "no name matched" in self.driver.page_source.lower(): 
                    print(f"RESULT: 0(0) -> N")
                    return "N"

            norm_all_eng = [self.normalize_english(v) for v in all_variants if not self.is_chinese(v)]
            norm_all_chi = [self.normalize_chinese(v) for v in all_variants if self.is_chinese(v)]
            
            exact_count = 0
            rows = self.driver.find_elements(By.XPATH, "//tr[contains(@class,'x-grid-row')]")
            
            for row in rows:
                try:
                    eng_portal = row.find_element(By.XPATH, ".//td[contains(@class,'gridcolumn-1035')]//div").text.strip()
                    chi_portal = row.find_element(By.XPATH, ".//td[contains(@class,'gridcolumn-1036')]//div").text.strip()
                    details_link = row.find_element(By.XPATH, ".//a[contains(text(),'details')]").get_attribute("href")
                    
                    norm_row_eng = self.normalize_english(eng_portal)
                    norm_row_chi = self.normalize_chinese(chi_portal)

                    # Counting exact matches for the specific query
                    if is_chi_query:
                        if norm_row_chi == self.normalize_chinese(query_name): exact_count += 1
                    else:
                        if norm_row_eng == self.normalize_english(query_name): exact_count += 1

                    # TAB OPENING LOGIC
                    is_match = False
                    if norm_row_eng in norm_all_eng:
                        if norm_all_chi:
                            if norm_row_chi:
                                if norm_row_chi in norm_all_chi: is_match = True
                            else: is_match = True
                        else: is_match = True

                    # Only open a new tab for the Licence Record if it's a match
                    if is_match and details_link not in self.opened_urls:
                        print(f"🔗 OPENING LINK: {details_link}")
                        self.opened_urls.add(details_link)
                        main_tab = self.driver.current_window_handle
                        self.driver.execute_script("window.open(arguments[0]);", details_link)
                        self.driver.switch_to.window(self.driver.window_handles[-1])
                        try:
                            btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@onclick,'licenceRecord')]")))
                            self.driver.execute_script("arguments[0].click();", btn)
                        except: pass
                        self.driver.switch_to.window(main_tab)
                except: continue

            # STATUS LOGIC
            status = "N"
            if total_hits > 0:
                status = "Y" if exact_count > 0 else "NRH"
            
            print(f"RESULT: {exact_count}({total_hits}) -> {status}")
            return status

        except Exception as e:
            print(f"Error: {e}")
            return "ERROR"

if __name__ == "__main__":
    # Updated file path to the standardized name in the same folder
    file_path = "Search_List.xlsx"
    if not os.path.exists(file_path):
        print(f"Error: {file_path} not found.")
        time.sleep(5)
        exit()

    try:
        # Load from Sheet1
        df = pd.read_excel(file_path, sheet_name="Sheet1")
        combinations = df["Combinations"].dropna().astype(str).tolist()
        bot = HKSFCSearcher()
        
        results_for_excel = []
        for name in combinations:
            res_status = bot.perform_search(name, combinations)
            results_for_excel.append(res_status)
        
        # Save results to the HKSFC column (Column B)
        df["HKSFC"] = results_for_excel
        
        saved = False
        while not saved:
            try:
                # Save specifically back to Sheet1 while preserving the file
                with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df.to_excel(writer, sheet_name="Sheet1", index=False)
                print("\n" + "="*45 + "\nSUCCESS: EXCEL FILE UPDATED\n" + "="*45)
                saved = True
            except PermissionError:
                input("\n⚠️  PERMISSION ERROR: Close Excel and press ENTER to retry...")

        print("\n[INFO] Complete. Console summary finished. Windows remain open.")
        while True: time.sleep(10)
    except Exception as e:
        print(f"Critical Error: {e}")
        input("Press Enter to close...")
print("SCRIPT_DONE")