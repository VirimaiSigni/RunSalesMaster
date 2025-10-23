


import pyautogui
import pygetwindow as gw
import schedule
import time

pyautogui.FAILSAFE = True

def run_sales_task():
    # Step 1: Open Start Menu
    pyautogui.press('win')
    time.sleep(1)

    # Step 2: Search and launch Sales Master
    pyautogui.write('Sales Master', interval=0.1)
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)  # Wait for the app to launch

    # Step 3: Press Enter to trigger Excel export
    pyautogui.press('enter')
    time.sleep(5)  # Wait for Excel to open

    # Step 4: Bring Excel to front

    excel_windows = gw.getWindowsWithTitle('Excel')
    if excel_windows:
        excel_windows[0].restore()
        excel_windows[0].activate()
        time.sleep(2)

        # Step 5: Select All and Save
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(1)
        pyautogui.hotkey('ctrl', 's')
        time.sleep(30)
    else:
        print("Excel window not found!")
schedule.every(30).seconds.do(run_sales_task)

print("Automation started. Press Ctrl+C to stop.")


while True:
    schedule.run_pending()
    time.sleep(1)

