import sys
import os
import ctypes
import random
import json
import threading
import time
import winreg
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QLineEdit,
    QFileDialog, QComboBox, QSpinBox, QHBoxLayout, QVBoxLayout,
    QMessageBox, QSystemTrayIcon, QMenu
)
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
from PIL import Image

# Constants
SETTINGS_FILE = 'settings.json'
TRAY_ICON_PATH = os.path.join('assets', 'tray_icon.ico')

# Helper function to set wallpaper
def set_wallpaper(image_path, style):
    # Constants for SystemParametersInfo
    SPI_SETDESKWALLPAPER = 20
    SPIF_UPDATEINIFILE = 0x01
    SPIF_SENDCHANGE = 0x02

    # Set wallpaper style
    # 0 = Center, 2 = Stretch, 6 = Fit, 10 = Fill, 22 = Span
    style_mapping = {
        "Fill Screen": "10",
        "Fit Screen": "6",
        "Stretch": "2",
        "Center": "0"
    }
    style_value = style_mapping.get(style, "10")  # Default to Fill

    # Accessing the registry to set wallpaper style
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            "Control Panel\\Desktop",
            0, winreg.KEY_SET_VALUE
        )
        winreg.SetValueEx(key, "WallpaperStyle", 0, winreg.REG_SZ, style_value)
        winreg.SetValueEx(key, "TileWallpaper", 0, winreg.REG_SZ, "0")
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Failed to set wallpaper style in registry: {e}")

    # Set the wallpaper
    try:
        ctypes.windll.user32.SystemParametersInfoW(
            SPI_SETDESKWALLPAPER, 0, image_path, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE
        )
    except Exception as e:
        print(f"Failed to set wallpaper: {e}")

class BackgroundSwitcherThread(QThread):
    # Signal to log messages (optional)
    log_signal = pyqtSignal(str)

    def __init__(self, get_settings_callback):
        super().__init__()
        self.get_settings = get_settings_callback
        self.is_running = True

    def run(self):
        while self.is_running:
            settings = self.get_settings()
            directory = settings.get("directory")
            interval = settings.get("interval", 60)
            order = settings.get("order", "Random")
            display_mode = settings.get("display_mode", "Fill Screen")

            if not directory or not os.path.isdir(directory):
                self.log_signal.emit("Invalid or no directory set. Waiting...")
                time.sleep(interval)
                continue

            images = [file for file in os.listdir(directory) if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif'))]
            if not images:
                self.log_signal.emit("No images found in the directory. Waiting...")
                time.sleep(interval)
                continue

            if order == "Random":
                image_file = random.choice(images)
            elif order == "Ascending Order":
                images.sort()
                image_file = images[0]
                images = images[1:] + [image_file]
            elif order == "Descending Order":
                images.sort(reverse=True)
                image_file = images[0]
                images = images[1:] + [image_file]
            else:
                image_file = random.choice(images)

            image_path = os.path.join(directory, image_file)
            if os.path.exists(image_path) and self.is_valid_image(image_path):
                set_wallpaper(image_path, display_mode)
                self.log_signal.emit(f"Wallpaper set to: {image_file}")
            else:
                self.log_signal.emit(f"Invalid or missing image: {image_file}")

            time.sleep(interval)

    def is_valid_image(self, image_path):
        try:
            with Image.open(image_path) as img:
                img.verify()
            return True
        except Exception:
            return False

    def stop(self):
        self.is_running = False
        self.quit()
        self.wait()

class BackgroundSwitcher(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Super Lightweight Background Switcher")
        self.setWindowIcon(QIcon(TRAY_ICON_PATH))
        self.setFixedSize(400, 300)
        self.init_ui()
        self.load_settings()
        self.init_tray()
        self.start_background_switcher()

    def init_ui(self):
        layout = QVBoxLayout()

        # Directory selection
        dir_layout = QHBoxLayout()
        self.dir_label = QLabel("Image Directory:")
        self.dir_input = QLineEdit()
        self.dir_browse = QPushButton("Browse")
        self.dir_browse.clicked.connect(self.browse_directory)
        dir_layout.addWidget(self.dir_label)
        dir_layout.addWidget(self.dir_input)
        dir_layout.addWidget(self.dir_browse)
        layout.addLayout(dir_layout)

        # Interval setting
        interval_layout = QHBoxLayout()
        self.interval_label = QLabel("Interval (seconds):")
        self.interval_input = QSpinBox()
        self.interval_input.setRange(10, 86400)  # 10 seconds to 24 hours
        interval_layout.addWidget(self.interval_label)
        interval_layout.addWidget(self.interval_input)
        layout.addLayout(interval_layout)

        # Order selection
        order_layout = QHBoxLayout()
        self.order_label = QLabel("Image Order:")
        self.order_combo = QComboBox()
        self.order_combo.addItems(["Random", "Ascending Order", "Descending Order"])
        order_layout.addWidget(self.order_label)
        order_layout.addWidget(self.order_combo)
        layout.addLayout(order_layout)

        # Display mode
        display_layout = QHBoxLayout()
        self.display_label = QLabel("Display Mode:")
        self.display_combo = QComboBox()
        self.display_combo.addItems(["Fill Screen", "Fit Screen", "Stretch", "Center"])
        display_layout.addWidget(self.display_label)
        display_layout.addWidget(self.display_combo)
        layout.addLayout(display_layout)

        # Save button
        self.save_button = QPushButton("Save Settings")
        self.save_button.clicked.connect(self.save_settings)
        layout.addWidget(self.save_button)

        self.setLayout(layout)

    def browse_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Image Directory")
        if directory:
            if os.path.isdir(directory):
                self.dir_input.setText(directory)
            else:
                QMessageBox.warning(self, "Invalid Directory", "The selected directory is invalid.")

    def save_settings(self):
        settings = {
            "directory": self.dir_input.text(),
            "interval": self.interval_input.value(),
            "order": self.order_combo.currentText(),
            "display_mode": self.display_combo.currentText()
        }
        try:
            with open(SETTINGS_FILE, 'w') as f:
                json.dump(settings, f, indent=4)
            QMessageBox.information(self, "Settings Saved", "Your settings have been saved successfully.")
            # After saving settings, minimize to tray
            self.hide()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {e}")

    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    settings = json.load(f)
            except json.JSONDecodeError:
                QMessageBox.warning(self, "Error", "Settings file is corrupted. Loading default settings.")
                settings = {}
        else:
            settings = {}

        self.dir_input.setText(settings.get("directory", ""))
        self.interval_input.setValue(settings.get("interval", 60))
        order = settings.get("order", "Random")
        index = self.order_combo.findText(order)
        if index != -1:
            self.order_combo.setCurrentIndex(index)
        display_mode = settings.get("display_mode", "Fill Screen")
        index = self.display_combo.findText(display_mode)
        if index != -1:
            self.display_combo.setCurrentIndex(index)

    def init_tray(self):
        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon(TRAY_ICON_PATH))
        self.tray_icon.setVisible(True)

        # Create tray menu
        tray_menu = QMenu()

        open_action = tray_menu.addAction("Open")
        open_action.triggered.connect(self.show_window)

        exit_action = tray_menu.addAction("Exit")
        exit_action.triggered.connect(self.exit_app)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_icon_activated)

    def on_tray_icon_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            self.show_window()

    def show_window(self):
        self.showNormal()
        self.activateWindow()

    def exit_app(self):
        self.background_thread.stop()
        self.tray_icon.hide()
        QApplication.quit()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "Background Switcher",
            "Application was minimized to Tray. To terminate the application, choose 'Exit' in the Tray Menu.",
            QSystemTrayIcon.Information,
            2000
        )

    def start_background_switcher(self):
        self.background_thread = BackgroundSwitcherThread(self.load_current_settings)
        self.background_thread.log_signal.connect(self.log_message)  # Optional: Connect to a logging method
        self.background_thread.start()

    def load_current_settings(self):
        settings = {
            "directory": self.dir_input.text(),
            "interval": self.interval_input.value(),
            "order": self.order_combo.currentText(),
            "display_mode": self.display_combo.currentText()
        }
        return settings

    def log_message(self, message):
        print(message)  # For debugging purposes; you can implement a logging UI or write to a file

def add_to_startup():
    import sys
    if not sys.argv[-1] == "startup":
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0, winreg.KEY_SET_VALUE
            )
            exe_path = sys.executable
            script_path = os.path.abspath(sys.argv[0])
            # Command to run the script with Python interpreter
            cmd = f'"{exe_path}" "{script_path}"'
            winreg.SetValueEx(key, "Super Lightweight Background Switcher", 0, winreg.REG_SZ, cmd)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Failed to add to startup: {e}")

if __name__ == "__main__":
    app = QApplication(sys.argv)

    # Add to startup
    add_to_startup()

    switcher = BackgroundSwitcher()
    
    # Show the GUI if no directory is set
    if not switcher.dir_input.text() or not os.path.isdir(switcher.dir_input.text()):
        switcher.show()
    else:
        switcher.hide()  # Start minimized to tray if settings are valid

    sys.exit(app.exec_())
