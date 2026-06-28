import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import traceback

# --- AUTOMATIC FILE PATH SETUP ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_NAME = os.path.join(BASE_DIR, 'Search_List.xlsx')
SEARCH_URL = "https://apps.sfc.hk/publicregWeb/searchByName?locale=en#"

def search_process():
    if not os.path.exists(FILE_NAME):
        print(f"CRITICAL ERROR: Could not find '{os.path.basename(FILE_NAME)}'")
        input("Press Enter to exit...")
        return

    try:
        # 1. LOAD EXCEL (Specifically Sheet 1 / Index 0)
        # We use None for sheet_name to load all or specify the sheet to edit.
        # To ensure we don't lose data, we read the specific sheet we want to change.
        df = pd.read_excel(FILE_NAME, sheet_name=0, dtype=str)
        
        # Ensure Column E exists for the status (Index 4)
        while df.shape[1] < 5:
            df[f"New_Col_{df.shape[1]}"] = ""
            
    except Exception as e:
        print(f"Error reading Excel: {e}")
        input("Press Enter to exit...")
        return

    # --- UPDATED DRIVER INITIALIZATION ---
    driver_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=driver_service)
    driver.maximize_window()
    
    try:
        wait = WebDriverWait(driver, 20)

        for index in range(len(df)):
            name = str(df.iloc[index, 0]).strip()
            
            if not name or name.lower() in ['nan', 'none', 'combinations']:
                continue

            print(f"--- Processing Row {index + 2}: {name} ---")

            try:
                driver.get(SEARCH_URL)
            except Exception as e:
                print(f"Network error loading page: {e}. Retrying...")
                time.sleep(3)
                driver.get(SEARCH_URL)

            try:
                time.sleep(2)

                # STEP 1: Open Search Overlay
                search_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "//i[contains(@class, 'fa-search')]")))
                search_icon.click()
                
                time.sleep(3)
                
                if len(driver.find_elements(By.ID, "popup-advanced-search")) == 0:
                    search_icon.click()
                    time.sleep(2)

                # STEP 2: Input Name
                search_input = wait.until(EC.visibility_of_element_located((By.ID, "popup-advanced-search")))
                search_input.clear()
                search_input.send_keys(name)
                
                time.sleep(1.5)

                # STEP 3: Click Search 
                submit_btn = wait.until(EC.element_to_be_clickable((By.ID, "basic-popup-search")))
                submit_btn.click()

                # STEP 4: Results Logic
                time.sleep(6) 
                
                try:
                    count_element = driver.find_element(By.CLASS_NAME, "searchResultTotal")
                    count_text = count_element.text.strip()
                    result_count = int(count_text.split()[-1]) if count_text else 0
                except:
                    result_count = 0

                if result_count == 0:
                    status = "N"
                else:
                    status = "NRH"
                
                df.iloc[index, 4] = status
                print(f"RESULT FOR SCRIPT: {result_count} records found | STATUS SAVED TO EXCEL: {status}")

            except Exception as e:
                print(f"Error during search for {name}: {type(e).__name__}") 
                df.iloc[index, 4] = "Error"

        # 3. SAVE RESULTS WITHOUT DELETING OTHER SHEETS
        # We use 'overlay' mode to update only the specific sheet while preserving others
        with pd.ExcelWriter(FILE_NAME, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            # We target the first sheet (usually index 0)
            # If your first sheet has a specific name, replace sheet_name=0 with sheet_name='Sheet1'
            df.to_excel(writer, sheet_name=writer.book.sheetnames[0], index=False)

        print("\n" + "="*40)
        print("SUCCESS: Excel updated. Sheet 2 and others have been preserved.")
        print("="*40)

    except BaseException as e:
        print(f"\nA major error occurred: {e}")
        traceback.print_exc()

    finally:
        print("\nProcess finished. Press ENTER to close...")
        input()
        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    search_process()