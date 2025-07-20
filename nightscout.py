import requests
import pystray
from PIL import Image, ImageDraw, ImageFont, Image as PILImage
import threading
import time
import os
import json
from datetime import datetime
import tkinter as tk
from tkinter import messagebox  # <-- FIX: import messagebox correctly
import sys
import winreg

CONFIG_FILE = os.path.expanduser("~\\AppData\\Local\\Temp\\nightscout_tray_config.json")
CACHE_FILE = os.path.expanduser("~\\AppData\\Local\\Temp\\nightscout_tray_cache.json")
CHECK_INTERVAL = 30  # seconds (for more frequent, lightweight polling)
API_PATH = "/api/v1/entries.json"  # fallback to entries.json, as your status.json does not have glucose!
APP_ICON_FILE = "nightscout.ico"

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

APP_ICON_PATH = resource_path(APP_ICON_FILE)

ARROW_TEXT = {
    'DoubleUp': '↑↑',
    'SingleUp': '↑',
    'FortyFiveUp': '↗',
    'Flat': '→',
    'FortyFiveDown': '↘',
    'SingleDown': '↓',
    'DoubleDown': '↓↓',
    'NONE': '',
    'NOT COMPUTABLE': '',
    'RATE OUT OF RANGE': '',
    None: '',
    '': ''
}

def is_windows_dark_mode():
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(
            registry,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0
    except Exception:
        return False  # Default to light if detection fails

def get_arrow_text(direction):
    return ARROW_TEXT.get(direction, '')

def get_glucose_class(sgv):
    try:
        sgv = int(sgv)
        if sgv < 70:
            return "critical-low"
        elif sgv > 180:
            return "critical-high"
        else:
            return "normal"
    except Exception:
        return "error"

def get_glucose_age(utc_time):
    if not utc_time or utc_time == "null":
        return 999
    now = time.time()
    try:
        if isinstance(utc_time, int) or str(utc_time).isdigit():
            entry_epoch = int(utc_time) / 1000
        else:
            dt = datetime.fromisoformat(str(utc_time).replace("Z", "+00:00"))
            entry_epoch = dt.timestamp()
        diff_minutes = int((now - entry_epoch) / 60)
        return diff_minutes
    except Exception:
        return 999

def load_arrow_font(size):
    for font_name in ["seguisym.ttf", "arialbd.ttf", "DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(font_name, size)
        except IOError:
            continue
    return ImageFont.load_default()

def create_text_icon(value, arrow_direction, glucose_class="normal"):
    size = 32
    img = PILImage.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    text_color = (255,255,255,255) if is_windows_dark_mode() else (0,0,0,255)
    for font_size in [18, 16, 14]:
        try:
            value_font = ImageFont.truetype("arialbd.ttf", font_size)
            break
        except:
            value_font = ImageFont.load_default()
    arrow_font = load_arrow_font(22)
    value_text = value if value else "?"
    arrow_text = get_arrow_text(arrow_direction)
    try:
        value_bbox = draw.textbbox((0, 0), value_text, font=value_font)
        value_w, value_h = value_bbox[2] - value_bbox[0], value_bbox[3] - value_bbox[1]
    except AttributeError:
        value_w, value_h = draw.textsize(value_text, font=value_font)
    try:
        arrow_bbox = draw.textbbox((0, 0), arrow_text, font=arrow_font)
        arrow_w, arrow_h = arrow_bbox[2] - arrow_bbox[0], arrow_bbox[3] - arrow_bbox[1]
    except AttributeError:
        arrow_w, arrow_h = draw.textsize(arrow_text, font=arrow_font)
    gap = 2
    arrow_y = (size - arrow_h) // 2
    value_y = arrow_y - value_h - gap if arrow_h + value_h + gap <= size else 0
    draw.text(((size - value_w)//2, value_y), value_text, font=value_font, fill=text_color)
    if arrow_text:
        draw.text(((size - arrow_w)//2, arrow_y), arrow_text, font=arrow_font, fill=text_color)
    return img

def load_config():
    if os.path.isfile(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                cfg = json.load(f)
                address = cfg.get("nightscout_address", "")
                token = cfg.get("token", "")
                return address, token
        except Exception:
            pass
    return "", ""

def save_config(address, token):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"nightscout_address": address, "token": token}, f)

def set_autostart(enable):
    app_name = "NightscoutTray"
    exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
    if exe_path.endswith(".py"):
        pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
        exe_path = f'"{pythonw}" "{exe_path}"'
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_SET_VALUE
        )
        if enable:
            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, exe_path)
        else:
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Failed to set autostart: {e}")
        return False

def get_autostart_status():
    app_name = "NightscoutTray"
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0, winreg.KEY_READ
        )
        try:
            val = winreg.QueryValueEx(key, app_name)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            winreg.CloseKey(key)
            return False
    except Exception:
        return False

def load_app_icon():
    if os.path.isfile(APP_ICON_PATH):
        try:
            return PILImage.open(APP_ICON_PATH)
        except Exception as e:
            print(f"Error loading app icon: {e}")
    return None

class NightscoutTray:
    def __init__(self):
        self.value = "?"
        self.arrow_direction = ""
        self.tooltip = ""
        self.glucose_class = "normal"
        self.icon = None
        self.update_thread = threading.Thread(target=self.update_loop, daemon=True)
        self.stop_flag = False
        self.cache = []
        self.nightscout_address, self.token = load_config()
        self.autostart_enabled = get_autostart_status()
        self.default_icon = load_app_icon()

    def get_api_url(self):
        base_address = self.nightscout_address.strip()
        token = self.token.strip()
        if not base_address or not token:
            return None
        api_url = base_address
        if api_url.endswith("/"):
            api_url = api_url[:-1]
        api_url = f"{api_url}{API_PATH}?token={token}&count=2"
        return api_url

    def fetch_data(self):
        url = self.get_api_url()
        if not url:
            self.cache = []
            return False
        try:
            resp = requests.get(url, timeout=10)
            resp.raise_for_status()
            data = resp.json()
            with open(CACHE_FILE, "w") as f:
                json.dump(data, f)
            self.cache = data
            return True
        except Exception as e:
            print("Error fetching Nightscout entries, using cache if available:", e)
            if os.path.isfile(CACHE_FILE):
                with open(CACHE_FILE, "r") as f:
                    data = json.load(f)
                self.cache = data
            else:
                self.cache = []
            return False

    def process_data(self):
        data = self.cache
        if isinstance(data, list) and len(data) > 0:
            latest = sorted(data, key=lambda x: x.get('date', 0), reverse=True)[0]
            self.value = str(latest.get('sgv', '?'))
            direction = latest.get('direction', '')
            self.arrow_direction = direction
            date_value = latest.get('date', latest.get('dateString', None))
            glucose_age_min = get_glucose_age(date_value)
            self.glucose_class = get_glucose_class(self.value)
            # Calculate delta if possible
            delta_text = ""
            if len(data) > 1:
                prev = sorted(data, key=lambda x: x.get('date', 0), reverse=True)[1]
                try:
                    delta = int(latest.get('sgv', 0)) - int(prev.get('sgv', 0))
                    delta_text = f"Delta: {delta:+d} mg/dL"
                except Exception:
                    delta_text = "Delta: ?"
            else:
                delta_text = "Delta: ?"
            try:
                if str(date_value).isdigit():
                    date_seconds = int(date_value) // 1000
                    local_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(date_seconds))
                else:
                    dt = datetime.fromisoformat(str(date_value).replace("Z", "+00:00"))
                    local_time = dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                local_time = "Unknown"
            arrow_display = get_arrow_text(direction)
            self.tooltip = (
                f"Glucose: {self.value} mg/dL\n"
                f"{delta_text}\n"
                f"Direction: {direction} ({arrow_display})\n"
                f"Updated: {local_time}\n"
                f"Age: {glucose_age_min} min ago"
            )
        else:
            self.value = "?"
            self.arrow_direction = ""
            self.glucose_class = "error"
            if not self.nightscout_address or not self.token:
                self.tooltip = "Nightscout address and/or token not configured"
            else:
                self.tooltip = "No valid Nightscout data"

    def update_icon(self):
        if self.icon:
            image = create_text_icon(self.value, self.arrow_direction, self.glucose_class)
            self.icon.icon = image
            self.icon.title = self.tooltip

    def update_once(self):
        ok = self.fetch_data()
        self.process_data()
        if self.icon:
            self.update_icon()
        return ok

    def update_loop(self):
        while not self.stop_flag:
            self.update_once()
            time.sleep(CHECK_INTERVAL)

    def configure_dialog(self):
        def show_dialog():
            root = tk.Tk()
            root.withdraw()
            bg_color = "#2d2d2d" if is_windows_dark_mode() else "#f0f0f0"
            fg_color = "#ffffff" if is_windows_dark_mode() else "#000000"

            dialog = tk.Toplevel(root)
            dialog.title("Configure Nightscout")
            dialog.geometry("520x220")
            dialog.resizable(False, False)
            dialog.configure(bg=bg_color)
            try:
                dialog.iconbitmap(APP_ICON_PATH)
            except Exception:
                try:
                    photo = tk.PhotoImage(file=APP_ICON_PATH)
                    dialog.iconphoto(False, photo)
                except Exception as e:
                    print(f"Could not set window icon: {e}")

            title_label = tk.Label(dialog, text="Nightscout Configuration", font=("Segoe UI", 13, "bold"), fg=fg_color, bg=bg_color)
            title_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(18, 6))

            tk.Label(dialog, text="Nightscout Address:", font=("Segoe UI", 10), fg=fg_color, bg=bg_color).grid(row=1, column=0, sticky="e", padx=(20, 8), pady=(8, 8))
            address_entry = tk.Entry(dialog, width=48, font=("Segoe UI", 10), fg=fg_color, bg=bg_color, insertbackground=fg_color)
            address_entry.insert(0, self.nightscout_address)
            address_entry.grid(row=1, column=1, padx=(0, 20), pady=(8, 8), sticky="w")

            tk.Label(dialog, text="Token:", font=("Segoe UI", 10), fg=fg_color, bg=bg_color).grid(row=2, column=0, sticky="e", padx=(20, 8), pady=(8, 8))
            token_entry = tk.Entry(dialog, width=48, font=("Segoe UI", 10), fg=fg_color, bg=bg_color, insertbackground=fg_color)
            token_entry.insert(0, self.token)
            token_entry.grid(row=2, column=1, padx=(0, 20), pady=(8, 8), sticky="w")

            def on_save():
                address = address_entry.get().strip()
                token = token_entry.get().strip()
                save_config(address, token)
                self.nightscout_address = address
                self.token = token
                ok = self.update_once()
                if not ok:
                    messagebox.showerror("Error", "Unable to contact Nightscout server with provided settings.")
                else:
                    messagebox.showinfo("Configuration", "Configuration saved and Nightscout contacted successfully.")
                dialog.destroy()
                root.destroy()

            def on_cancel():
                dialog.destroy()
                root.destroy()

            button_frame = tk.Frame(dialog, bg=bg_color)
            button_frame.grid(row=3, column=0, columnspan=2, sticky="e", padx=20, pady=(12, 16))

            save_btn = tk.Button(button_frame, text="Save", width=10, command=on_save, fg=fg_color, bg=bg_color)
            save_btn.pack(side="right", padx=(0, 8))
            cancel_btn = tk.Button(button_frame, text="Cancel", width=10, command=on_cancel, fg=fg_color, bg=bg_color)
            cancel_btn.pack(side="right")

            dialog.protocol("WM_DELETE_WINDOW", on_cancel)
            dialog.mainloop()

        threading.Thread(target=show_dialog, daemon=True).start()

    def toggle_autostart(self):
        new_status = not self.autostart_enabled
        set_autostart(new_status)
        self.autostart_enabled = get_autostart_status()

    def build_menu(self):
        return pystray.Menu(
            pystray.MenuItem("Configure", lambda icon, item: self.configure_dialog()),
            pystray.MenuItem(
                "Autostart on boot",
                lambda icon, item: self.toggle_autostart(),
                checked=lambda item: self.autostart_enabled,
            ),
            pystray.MenuItem("Quit", lambda icon, item: self.quit())
        )

    def run(self):
        self.update_once()
        image = self.default_icon if self.default_icon else create_text_icon(self.value, self.arrow_direction, self.glucose_class)
        self.icon = pystray.Icon("Nightscout", image, "Nightscout", self.build_menu())
        self.update_thread.start()
        self.icon.run()

    def quit(self):
        self.stop_flag = True
        self.icon.stop()

if __name__ == "__main__":
    NightscoutTray().run()