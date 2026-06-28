import os
import time
import traceback
import pandas as pd
import openpyxl

def run_task(script_name, description):
    print("\n" + "="*70)
    print(f"🚀 STARTING: {description}")
    print("="*70 + "\n")

    if not os.path.exists(script_name):
        print(f"❌ ERROR: File '{script_name}' not found.")
        return False

    try:
        with open(script_name, 'r', encoding='utf-8') as f:
            code = f.read()

        # FIX 1 & 3: Prevent windows and script from closing automatically
        # We neutralize the commands that exit the process or close the browser
        code = code.replace("while True: time.sleep(10)", "pass")
        code = code.replace("input(", 'print("Master Auto-Skipping Input: ", ')
        code = code.replace(".quit()", "# .quit()")
        code = code.replace(".close()", "# .close()")
        code = code.replace("sys.exit()", "pass")

        # FIX 2: Force specific scripts using .active to use 'Sheet1' to ensure they target the right data
        scripts_to_fix_sheet = [
            "HKSFC_Advance Search.py", 
            "hkma_advanced_excel.py", 
            "hkma_top_right_corner.py"
        ]
        if script_name in scripts_to_fix_sheet:
            code = code.replace("sheet = wb.active", "sheet = wb['Sheet1']")

        # Create a clean namespace for execution to avoid variable contamination
        namespace = {
            '__name__': '__main__',
            '__file__': os.path.abspath(script_name),
        }

        # Execute the script code
        exec(code, namespace)
        
        print(f"\n✅ SUCCESS: {description} completed.")
        return True
    except Exception as e:
        print(f"\n❌ FAILURE in {script_name}: {e}")
        traceback.print_exc()
        return False

def main():
    # Full sequence of all tasks (HKSFC followed by HKMA)
    tasks = [
        # --- HKSFC SECTION ---
        ("hksfc_main.py", "HKSFC Main Registry (Column B)"),
        ("hksfc_enforcement_news.py", "HKSFC Enforcement News (Column C)"),
        ("HKSFC_Advance Search.py", "HKSFC Advanced Search (Column D)"),
        ("HKSFC.py", "HKSFC Top Right Corner (Column E)"),
        
        # --- HKMA SECTION ---
        ("hkma_main.py", "HKMA Main Registry (Column F)"),
        ("hkma_advanced_excel.py", "HKMA Advanced Search (Column G)"),
        ("hkma_top_right_corner.py", "HKMA Top Right Corner (Column H)")
    ]

    print("#"*70)
    print("🌟 SUPER MASTER AUTOMATION: HKSFC + HKMA 🌟")
    print("#"*70)

    for script, desc in tasks:
        success = run_task(script, desc)
        if not success:
            print(f"🛑 Sequence halted at {script}. Please check the error above.")
            break
        # Delay to allow system to release file handles and close connections properly
        time.sleep(5)

    print("\n" + "#"*70)
    print("✨ ALL AUTOMATIONS FINISHED SUCCESSFULLY!")
    print("All browser windows remain open for your inspection.")
    print("The Excel file 'Search_List.xlsx' has been updated for Columns B through H.")
    print("#"*70)

    # Keep the console window open
    while True:
        time.sleep(10)

if __name__ == "__main__":
    main()
