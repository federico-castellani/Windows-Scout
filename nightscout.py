import requests
import pystray
from PIL import Image, ImageDraw, ImageFont, Image as PILImage
import threading
import time
import os
import json
from datetime import datetime
import tkinter as tk
from tkinter import messagebox
import sys
import socket
import hashlib
import base64
import zlib
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad

CONFIG_FILE = os.path.expanduser("~\\AppData\\Local\\Temp\\nightscout_tray_config.json")
CACHE_FILE = os.path.expanduser("~\\AppData\\Local\\Temp\\nightscout_tray_cache.json")
APP_ICON_FILE = "nightscout.ico"

HOST = "jamcm3749021.bluejay.website"
PORT = 25935

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

APP_ICON_PATH = resource_path(APP_ICON_FILE)

# --- Miniature Pure-Python Protobuf Parser ---
def read_varint(data, offset=0):
    result = 0
    shift = 0
    while True:
        if offset >= len(data): break
        byte = data[offset]
        offset += 1
        result |= (byte & 0x7f) << shift
        shift += 7
        if not (byte & 0x80): break
    return result, offset

def parse_protobuf(data):
    fields = {}
    offset = 0
    while offset < len(data):
        tag, offset = read_varint(data, offset)
        field_num = tag >> 3
        wire_type = tag & 0x07

        if wire_type == 0:
            val, offset = read_varint(data, offset)
            fields[field_num] = val
        elif wire_type == 1: offset += 8
        elif wire_type == 2:
            length, offset = read_varint(data, offset)
            fields[field_num] = data[offset:offset+length]
            offset += length
        elif wire_type == 5: offset += 4
        else: break
    return fields

def get_identity(sync_key):
    return hashlib.sha256(sync_key.encode('utf-8')).hexdigest()[:32].lower()

def decrypt_payload(sync_key, inbytes):
    try:
        key_constant = "ebe5c0df162a50ba232d2d721ea8e3e1c5423bb0-12bd-48c3-8932-c93883dfcf1f"
        sync_key_md5 = hashlib.md5(sync_key.encode('utf-8')).hexdigest().lower()
        full_key_string = key_constant + sync_key_md5
        aes_key = hashlib.md5(full_key_string.encode('utf-8')).digest()
        
        try:
            inbytes = base64.b64decode(inbytes.decode('ascii'))
        except:
            pass

        if len(inbytes) < 16:
            return inbytes.decode('utf-8', errors='replace')
            
        iv = inbytes[:16]
        cipher_data = inbytes[16:]
        
        cipher = AES.new(aes_key, AES.MODE_CBC, iv)
        try:
            decrypted = unpad(cipher.decrypt(cipher_data), AES.block_size)
        except ValueError:
            decrypted = cipher.decrypt(cipher_data)
        
        if len(decrypted) > 8 and decrypted[0] == 0x1F and decrypted[1] == 0x8B:
            try:
                decrypted = zlib.decompress(decrypted, 16 + zlib.MAX_WBITS)
            except zlib.error:
                decrypted = zlib.decompress(decrypted, -zlib.MAX_WBITS)
            
        return decrypted.decode('utf-8')
    except Exception:
        return None

def send_msg(s, cmd, payload=None, param=None):
    header = cmd
    l = len(payload) if payload else 0
    if l != 0 or param is not None: header += f" {l}"
    if param is not None: header += f" {param}"
    s.sendall(header.encode('utf-8') + b'\n')
    if payload: s.sendall(payload)

def is_windows_dark_mode():
    try:
        import winreg
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(
            registry,
            r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        )
        value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        winreg.CloseKey(key)
        return value == 0
    except Exception:
        return False

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

def load_arrow_font(size):
    for font_name in ["seguisym.ttf", "arialbd.ttf", "DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(font_name, size)
        except IOError:
            continue
    return ImageFont.load_default()

def create_text_icon(value, arrow_text, glucose_class="normal"):
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
    value_text = str(value) if value else "?"
    if not arrow_text: arrow_text = ""
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
                sync_key = cfg.get("sync_key", "")
                if not sync_key and "token" in cfg:
                    sync_key = cfg.get("token", "")
                return sync_key
        except Exception:
            pass
    return ""

def save_config(sync_key):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"sync_key": sync_key}, f)

def get_startup_shortcut_path():
    return os.path.join(os.environ["APPDATA"], "Microsoft", "Windows", "Start Menu", "Programs", "Startup", "NightscoutTray.lnk")

def set_autostart(enable):
    shortcut_path = get_startup_shortcut_path()
    if enable:
        exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(sys.argv[0])
        work_dir = os.path.dirname(exe_path)
        
        if exe_path.endswith(".py"):
            pythonw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
            target = pythonw
            args = f'"{exe_path}"'
        else:
            target = exe_path
            args = ""
        
        vbs_script = f"""
Set ws = WScript.CreateObject("WScript.Shell")
Set shortcut = ws.CreateShortcut("{shortcut_path}")
shortcut.TargetPath = "{target}"
shortcut.Arguments = "{args}"
shortcut.WorkingDirectory = "{work_dir}"
shortcut.Save
"""
        vbs_path = os.path.join(os.environ["TEMP"], "create_ns_shortcut.vbs")
        try:
            with open(vbs_path, "w") as f:
                f.write(vbs_script)
            os.system(f'cscript //nologo "{vbs_path}"')
            os.remove(vbs_path)
            return True
        except Exception as e:
            print(f"Failed to set autostart: {e}")
            return False
    else:
        if os.path.exists(shortcut_path):
            try:
                os.remove(shortcut_path)
                return True
            except:
                return False
        return True

def get_autostart_status():
    return os.path.exists(get_startup_shortcut_path())

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
        self.arrow_text = ""
        self.tooltip = "Initializing..."
        self.glucose_class = "normal"
        self.icon = None
        self.stop_flag = False
        self.sync_key = load_config()
        self.autostart_enabled = get_autostart_status()
        self.default_icon = load_app_icon()
        self.last_bg = None
        self.socket_thread = None
        self.s = None

    def update_icon(self):
        if self.icon:
            image = create_text_icon(self.value, self.arrow_text, self.glucose_class)
            self.icon.icon = image
            self.icon.title = self.tooltip

    def process_payload(self, payload_str):
        try:
            data = json.loads(payload_str)
            sgv = data.get("calculated_value", 0)
            slope = data.get("calculated_value_slope", 0)
            
            arrow = "→"
            if slope > 2: arrow = "⇈"
            elif slope > 1: arrow = "↑"
            elif slope > 0.5: arrow = "↗"
            elif slope < -2: arrow = "⇊"
            elif slope < -1: arrow = "↓"
            elif slope < -0.5: arrow = "↘"
            
            self.glucose_class = get_glucose_class(sgv)
            
            delta_text = "No change"
            if self.last_bg is not None:
                delta = sgv - self.last_bg
                delta_text = f"{delta:+d} mg/dL"
            self.last_bg = sgv
            
            self.value = str(sgv)
            self.arrow_text = arrow
            
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.tooltip = f"Glucose: {sgv} mg/dL\nChange: {delta_text}\nSlope: {slope:.2f}\nUpdated: {now_str}"
            self.update_icon()
        except json.JSONDecodeError:
            pass

    def handle_message(self, binary, parameter, sync_key):
        try:
            trans = parse_protobuf(binary)
            if trans.get(3) == 2:
                sync_payload_bytes = trans.get(5, b'')
                msg = parse_protobuf(sync_payload_bytes)
                action = msg.get(2, b'').decode('utf-8', errors='replace')
                raw_payload_bytes = msg.get(3, b'')
                
                if action == 'bgs':
                    payload = decrypt_payload(sync_key, raw_payload_bytes)
                    if payload:
                        self.process_payload(payload)
        except Exception:
            pass

    def connection_loop(self):
        while not self.stop_flag:
            if not self.sync_key:
                self.value = "?"
                self.arrow_text = ""
                self.tooltip = "Sync Key not configured"
                self.update_icon()
                time.sleep(5)
                continue

            self.value = "..."
            self.arrow_text = ""
            self.tooltip = "Connecting..."
            self.update_icon()

            my_identity = get_identity(self.sync_key)
            token = "pc_receiver_" + hashlib.sha256(str(time.time()).encode()).hexdigest()[:20]

            self.s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.s.settimeout(60.0)
            
            try:
                self.s.connect((HOST, PORT))
            except Exception as e:
                self.value = "!"
                self.tooltip = "Connection failed"
                self.update_icon()
                time.sleep(10)
                continue
            
            def send_line(msg):
                self.s.sendall(msg.encode('utf-8') + b'\n')

            send_line("1,1")
            send_line(token)
            send_line(my_identity)

            def read_line():
                line = bytearray()
                while True:
                    c = self.s.recv(1)
                    if not c: return None
                    if c == b'\n': return line.decode('utf-8').strip()
                    line += c

            try:
                resp = read_line()
            except Exception:
                resp = None

            if resp != "OK":
                self.value = "!"
                self.tooltip = "Auth failed"
                self.update_icon()
                self.s.close()
                time.sleep(10)
                continue

            self.tooltip = "Connected, waiting for data..."
            self.update_icon()

            def ping_thread(sock, stop_event):
                while not stop_event.is_set():
                    time.sleep(30)
                    try:
                        send_msg(sock, "P")
                    except:
                        break

            stop_ping = threading.Event()
            threading.Thread(target=ping_thread, args=(self.s, stop_ping), daemon=True).start()

            while not self.stop_flag:
                try:
                    line = read_line()
                    if line is None: break
                    parts = line.split(" ", 2)
                    cmd = parts[0]
                    size = int(parts[1]) if len(parts) > 1 else 0
                    param = parts[2] if len(parts) > 2 else ""

                    binary = b""
                    if size > 0:
                        while len(binary) < size:
                            chunk = self.s.recv(size - len(binary))
                            if not chunk: raise ConnectionError()
                            binary += chunk

                    if cmd == "O": send_msg(self.s, "K")
                    elif cmd == "M":
                        self.handle_message(binary, param, self.sync_key)
                        send_msg(self.s, "N", None, param)
                    elif cmd == "CEASE": break
                except socket.timeout:
                    try: send_msg(self.s, "P")
                    except: break
                except Exception:
                    break

            stop_ping.set()
            if self.s:
                try: self.s.close()
                except: pass
            
            if not self.stop_flag:
                self.value = "!"
                self.tooltip = "Disconnected, reconnecting..."
                self.update_icon()
                time.sleep(5)

    def configure_dialog(self):
        def show_dialog():
            root = tk.Tk()
            root.withdraw()
            bg_color = "#2d2d2d" if is_windows_dark_mode() else "#f0f0f0"
            fg_color = "#ffffff" if is_windows_dark_mode() else "#000000"

            dialog = tk.Toplevel(root)
            dialog.title("Configure Sync")
            dialog.geometry("520x150")
            dialog.resizable(False, False)
            dialog.configure(bg=bg_color)
            try:
                dialog.iconbitmap(APP_ICON_PATH)
            except Exception:
                pass

            title_label = tk.Label(dialog, text="Sync Configuration", font=("Segoe UI", 13, "bold"), fg=fg_color, bg=bg_color)
            title_label.grid(row=0, column=0, columnspan=2, sticky="w", padx=20, pady=(18, 6))

            tk.Label(dialog, text="Sync Key:", font=("Segoe UI", 10), fg=fg_color, bg=bg_color).grid(row=1, column=0, sticky="e", padx=(20, 8), pady=(8, 8))
            sync_entry = tk.Entry(dialog, width=48, font=("Segoe UI", 10), fg=fg_color, bg=bg_color, insertbackground=fg_color)
            sync_entry.insert(0, self.sync_key)
            sync_entry.grid(row=1, column=1, padx=(0, 20), pady=(8, 8), sticky="w")

            def on_save():
                new_key = sync_entry.get().strip()
                save_config(new_key)
                self.sync_key = new_key
                self.last_bg = None 
                if self.s:
                    try: self.s.close()
                    except: pass
                dialog.destroy()
                root.destroy()

            def on_cancel():
                dialog.destroy()
                root.destroy()

            button_frame = tk.Frame(dialog, bg=bg_color)
            button_frame.grid(row=2, column=0, columnspan=2, sticky="e", padx=20, pady=(12, 16))

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
            pystray.MenuItem("Configure Sync", lambda icon, item: self.configure_dialog()),
            pystray.MenuItem(
                "Autostart on boot",
                lambda icon, item: self.toggle_autostart(),
                checked=lambda item: self.autostart_enabled,
            ),
            pystray.MenuItem("Quit", lambda icon, item: self.quit())
        )

    def run(self):
        image = self.default_icon if self.default_icon else create_text_icon("...", "", "normal")
        self.icon = pystray.Icon("Nightscout", image, "Nightscout - Starting...", self.build_menu())
        
        self.socket_thread = threading.Thread(target=self.connection_loop, daemon=True)
        self.socket_thread.start()
        
        self.icon.run()

    def quit(self):
        self.stop_flag = True
        if self.s:
            try: self.s.close()
            except: pass
        self.icon.stop()

if __name__ == "__main__":
    NightscoutTray().run()
