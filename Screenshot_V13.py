import sys
import os
import re
import socket
import datetime
import subprocess
import keyboard
from PyQt5.QtWidgets import QApplication, QSystemTrayIcon, QMenu, QMessageBox
from PyQt5.QtGui import QIcon
import logging
import qtmodern
import qtmodern.styles

if getattr(sys, 'frozen', False):
    application_path = sys._MEIPASS
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(application_path, 'format.ico')

class LogWriter:
    LOG_FOLDER = r"M:\QA_Program_Raw_Data\Log History\Screenshot"

    @staticmethod
    def save_log():
        try:
            username = get_username()
            current_datetime = get_current_datetime()
            log_message = f"{current_datetime} {username} Open\n"

            log_folder_path = os.path.join(LogWriter.LOG_FOLDER)
            os.makedirs(log_folder_path, exist_ok=True)

            log_file_path = os.path.join(log_folder_path, f"{username}.txt")
            with open(log_file_path, 'a') as log_file:
                log_file.write(log_message)
        except Exception as e:
            print(f"Error saving log: {e}")

def get_username():
    try:
        hostname = socket.gethostname()
        username = hostname.split('.')[0]

        restricted_usernames = ["A005239", "A005359", "A005492", "A005573", "A005715", "A005721", "A005844", "A005943", "A005950", "A005958", "A005959", "A005965",
                                "A005980", "A005986", "A006098", "A006149", "A006172", "A006204", "A006209", "A0733", "A0888", "A2830", "A3003", "A3004", "A3211", "A3933", "A4505", "A4895", 
                                "A4975", "A4987", "A4997", "A5065", "VDI"]
        
        for restricted_username in restricted_usernames:
            if restricted_username in username:
                if restricted_username == "-" and username.endswith("-NB"):
                    break
                else:
                    show_message_box("Warning", "Activation authority not obtained, please contact #1082 Racky")
                    sys.exit(0)

        return username
    except Exception:
        return "Unknown"

def get_current_datetime():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

class VersionChecker:
    APP_DIR = r"M:\QA_Program_Raw_Data\Apps"

    @staticmethod
    def check_version():
        try:
            if not os.path.exists(VersionChecker.APP_DIR):
                show_message_box("Warning", "未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky")
                return False

            exe_files = [f for f in os.listdir(VersionChecker.APP_DIR) if f.startswith("Screenshot_V") and f.endswith(".exe")]
            
            if not exe_files:
                show_message_box("Warning", "未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky")
                return False

            latest_version = max(int(re.search(r'_V(\d+)\.exe', f).group(1)) for f in exe_files)

            current_version_match = re.search(r'_V(\d+)\.exe', os.path.basename(sys.executable))
            if current_version_match:
                current_version = int(current_version_match.group(1))
            else:
                current_version = 0

            if current_version < latest_version:
                show_message_box("Warning", "請更新至最新版本")
                VersionChecker.open_directory(VersionChecker.APP_DIR)
                return False

            hostname = socket.gethostname()
            match = re.search(r'^(.+)', hostname)
            if match:
                username = match.group(1)
                restricted_usernames = ["A005239", "A005359", "A005492", "A005573", "A005715", "A005721", "A005844", "A005943", "A005950", "A005958", "A005959", "A005965",
                                        "A005980", "A005986", "A006098", "A006149", "A006172", "A006204", "A006209", "A0733", "A0888", "A2830", "A3003", "A3004", "A3211", "A3933", "A4505", "A4895", 
                                        "A4975", "A4987", "A4997", "A5065", "VDI"]
                
                if any(restricted_username in username for restricted_username in restricted_usernames):
                    if not (username.endswith("-NB") and "-" in username):
                        show_message_box("Warning", "未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky")
                        return False
            else:
                show_message_box("Warning", "未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky")
                return False

            return True

        except Exception as e:
            print(f"Error checking version: {e}")
            show_message_box("Warning", "未獲取啟動權限, 請申請M:\QA_Program_Raw_Data權限, 並聯絡#1082 Racky")
            return False

    @staticmethod
    def open_directory(dir_path):
        try:
            if os.name == 'nt':  # Windows
                os.startfile(dir_path)
            elif os.name == 'posix':  # macOS and Linux
                subprocess.call(['open', dir_path])
        except Exception as e:
            print(f"Error opening directory: {e}")

def show_message_box(title, message):
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Information)
    msg_box.setText(message)
    msg_box.setWindowTitle(title)
    msg_box.setStandardButtons(QMessageBox.Ok)
    msg_box.exec_()

def on_hotkey():
    subprocess.Popen(["explorer.exe", "ms-screenclip:"])

class ScreenshotTool(QSystemTrayIcon):
    def __init__(self, icon, parent=None):
        super().__init__(icon, parent)
        self.setToolTip('Screenshot Tool')
        
        menu = QMenu(parent)
        exit_action = menu.addAction("Exit")
        exit_action.triggered.connect(self.exit_app)
        
        self.setContextMenu(menu)
        self.activated.connect(self.on_tray_activated)

        # Register global hotkey
        keyboard.add_hotkey('F2', on_hotkey)

    def on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.DoubleClick:
            on_hotkey()

    def exit_app(self):
        keyboard.unhook_all()  # Remove all hotkeys
        QApplication.quit()

def main():
    app = QApplication(sys.argv)
    app_icon = QIcon(icon_path)
    app.setWindowIcon(app_icon)
    
    qtmodern.styles.dark(app)
    
    try:
        if not VersionChecker.check_version():
            return

        LogWriter.save_log()

        tray = ScreenshotTool(app_icon) 
        tray.show()

        sys.exit(app.exec_())
    except Exception as e:
        logging.exception("Error in main function")
        show_message_box("Error", f"程序發生錯誤: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()