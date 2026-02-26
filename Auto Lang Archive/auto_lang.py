import ctypes
import time
import pygetwindow as gw # התקנה: pip install pygetwindow

# הגדרות שפה (Locale IDs)
HEBREW = 0x040D
ENGLISH = 0x0409

# רשימות אפליקציות לפי שפה
ENGLISH_APPS = ["Visual Studio Code", "Terminal", "PyCharm", "Command Prompt", "Slack", "Outlook", "Excel", "PowerPoint","ASG"]
HEBREW_APPS = ["WhatsApp", "Telegram", "Word", "Team"]

def set_keyboard_language(lang_id):
    user32 = ctypes.WinDLL('user32', use_last_error=True)
    curr_window = user32.GetForegroundWindow()
    user32.PostMessageW(curr_window, 0x50, 0, lang_id)

def main():
    print("הסקריפט מנטר חלונות... (לסיום לחץ Ctrl+C)")
    last_active_window = ""

    while True:
        try:
            active_window = gw.getActiveWindow()
            if active_window and active_window.title != last_active_window:
                title = active_window.title
                last_active_window = title
                
                # בדיקה מול רשימת האנגלית
                if any(app.lower() in title.lower() for app in ENGLISH_APPS):
                    set_keyboard_language(ENGLISH)
                    print(f"החלפת לאנגלית עבור: {title}")
                
                # בדיקה מול רשימת העברית
                elif any(app.lower() in title.lower() for app in HEBREW_APPS):
                    set_keyboard_language(HEBREW)
                    print(f"החלפת לעברית עבור: {title}")
                    
        except Exception:
            pass
        
        time.sleep(0.5)

if __name__ == "__main__":
    main()
