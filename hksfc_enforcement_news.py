import os
import pandas as pd
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

class SFCAutomation:
    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True) 
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        
        self.driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
        self.wait = WebDriverWait(self.driver, 15)
        self.opened_urls = set() # Global deduplication
        self.search_tab_handle = None 

    def normalize(self, text):
        """Standardizes text for comparison."""
        if not text: return ""
        return " ".join(str(text).strip().lower().split())

    def is_exact_name_in_text(self, search_name, headline_text):
        """
        Checks if the search_name appears as a distinct phrase within the headline.
        This handles cases like 'SFC bans Lui Pak Tong for life...'
        """
        search_norm = self.normalize(search_name)
        headline_norm = self.normalize(headline_text)
        
        if not search_norm: return False
        
        # Use regex to find the name as a whole phrase (word boundaries)
        pattern = r'\b' + re.escape(search_norm) + r'\b'
        return bool(re.search(pattern, headline_norm))

    def perform_search(self, query_name):
        name = str(query_name).strip()
        if not name or name.lower() == 'nan': return "N", 0
        
        try:
            if not self.search_tab_handle:
                self.search_tab_handle = self.driver.current_window_handle

            self.driver.get("https://apps.sfc.hk/edistributionWeb/gateway/EN/news-and-announcements/news/enforcement-news/")
            
            # Navigate to search
            go_btn = self.wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Go to name search']")))
            self.driver.execute_script("arguments[0].click();", go_btn)

            # Input and Search
            name_input = self.wait.until(EC.visibility_of_element_located((By.NAME, "search-content")))
            name_input.clear()
            name_input.send_keys(name)

            search_btn = self.wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input.btn-primary[value='Search']")))
            self.driver.execute_script("arguments[0].click();", search_btn)
            
            time.sleep(3) 

            # Check for results
            result_links = self.driver.find_elements(By.CSS_SELECTOR, "table tbody tr td a")
            if not result_links:
                return "N", 0

            found_match = False
            tabs_opened_this_search = 0

            for link in result_links:
                headline = link.text
                if self.is_exact_name_in_text(name, headline):
                    found_match = True
                    link_url = link.get_attribute("href")
                    
                    # Deduplication Check
                    if link_url and link_url not in self.opened_urls:
                        self.opened_urls.add(link_url)
                        
                        # Open in new tab using Control + Click
                        actions = ActionChains(self.driver)
                        actions.key_down(Keys.CONTROL).click(link).key_up(Keys.CONTROL).perform()
                        
                        # Return focus to main search window
                        self.driver.switch_to.window(self.search_tab_handle)
                        tabs_opened_this_search += 1
                        time.sleep(0.5)

            status = "Y" if found_match else "N"
            return status, tabs_opened_this_search

        except Exception as e:
            print(f"⚠️ Error searching '{name}': {e}")
            return "ERROR", 0

def main():
    # Standardized file path for the HK FRC Automation folder
    file_path = 'Search_List.xlsx'

    if not os.path.exists(file_path):
        print(f"❌ File not found: {file_path}")
        return

    try:
        # Load specifically from Sheet1
        full_df = pd.read_excel(file_path, header=None, sheet_name="Sheet1")
        
        # Find where 'combination' header is
        header_row_idx, name_col_idx = None, None
        for r_idx, row in full_df.iterrows():
            for c_idx, val in enumerate(row):
                if "combination" in str(val).lower():
                    header_row_idx, name_col_idx = r_idx, c_idx
                    break
            if header_row_idx is not None: break

        bot = SFCAutomation()
        
        for idx in range(header_row_idx + 1, len(full_df)):
            name_val = full_df.iloc[idx, name_col_idx]
            if pd.isna(name_val): continue
            
            print(f"🔍 Checking: {name_val}")
            status, opened = bot.perform_search(name_val)
            
            # Write into Column C (name_col_idx + 2)
            full_df.iloc[idx, name_col_idx + 2] = status
            
            # Save results while preserving Sheet1 name
            with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                full_df.to_excel(writer, sheet_name="Sheet1", index=False, header=False)
                
            print(f"   Done. Status: {status} | New Tabs: {opened}")

    except Exception as e:
        print(f"❌ Fatal Error: {e}")
    
    input("\nTask Complete. Press Enter to exit...")

if __name__ == "__main__":
    main()
print("SCRIPT_DONE")