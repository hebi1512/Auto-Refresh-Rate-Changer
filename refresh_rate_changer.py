import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import json
import os
import psutil
import threading
import time
import ctypes
from ctypes import wintypes
import pystray
from PIL import Image, ImageDraw

CONFIG_FILE = "refresh_config.json"

# --------- Phần thay đổi refresh rate ---------
class DEVMODE(ctypes.Structure):
    _fields_ = [
        ("dmDeviceName", wintypes.WCHAR * 32),
        ("dmSpecVersion", wintypes.WORD),
        ("dmDriverVersion", wintypes.WORD),
        ("dmSize", wintypes.WORD),
        ("dmDriverExtra", wintypes.WORD),
        ("dmFields", wintypes.DWORD),
        ("dmPosition_x", wintypes.LONG),
        ("dmPosition_y", wintypes.LONG),
        ("dmDisplayOrientation", wintypes.DWORD),
        ("dmDisplayFixedOutput", wintypes.DWORD),
        ("dmColor", wintypes.SHORT),
        ("dmDuplex", wintypes.SHORT),
        ("dmYResolution", wintypes.SHORT),
        ("dmTTOption", wintypes.SHORT),
        ("dmCollate", wintypes.SHORT),
        ("dmFormName", wintypes.WCHAR * 32),
        ("dmLogPixels", wintypes.WORD),
        ("dmBitsPerPel", wintypes.DWORD),
        ("dmPelsWidth", wintypes.DWORD),
        ("dmPelsHeight", wintypes.DWORD),
        ("dmDisplayFlags", wintypes.DWORD),
        ("dmDisplayFrequency", wintypes.DWORD),
        ("dmICMMethod", wintypes.DWORD),
        ("dmICMIntent", wintypes.DWORD),
        ("dmMediaType", wintypes.DWORD),
        ("dmDitherType", wintypes.DWORD),
        ("dmReserved1", wintypes.DWORD),
        ("dmReserved2", wintypes.DWORD),
        ("dmPanningWidth", wintypes.DWORD),
        ("dmPanningHeight", wintypes.DWORD),
    ]

ENUM_CURRENT_SETTINGS = -1
CDS_UPDATEREGISTRY = 0x01
DISP_CHANGE_SUCCESSFUL = 0
DM_DISPLAYFREQUENCY = 0x400000

user32 = ctypes.windll.user32

def set_refresh_rate(hz: int):
    devmode = DEVMODE()
    devmode.dmSize = ctypes.sizeof(DEVMODE)

    if user32.EnumDisplaySettingsW(None, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)) == 0:
        return False

    devmode.dmFields = DM_DISPLAYFREQUENCY
    devmode.dmDisplayFrequency = hz

    result = user32.ChangeDisplaySettingsW(ctypes.byref(devmode), CDS_UPDATEREGISTRY)
    return result == DISP_CHANGE_SUCCESSFUL

# --------- Phần quản lý cấu hình ---------
def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {}

def save_config(cfg):
    with open(CONFIG_FILE, "w") as f:
        json.dump(cfg, f, indent=4)

# --------- GUI ---------
class RefreshApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Refresh Rate Manager")

        self.config = load_config()

        self.icon_running = False

        tk.Label(root, text="Ứng dụng:").grid(row=0, column=0, padx=5, pady=5)
        self.app_entry = tk.Entry(root, width=40)
        self.app_entry.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(root, text="Chọn file .exe", command=self.choose_file).grid(row=0, column=2, padx=5)

        tk.Label(root, text="Refresh rate:").grid(row=1, column=0, padx=5, pady=5)
        self.rate_var = tk.StringVar()
        self.rate_combo = ttk.Combobox(root, textvariable=self.rate_var, values=["120", "144"])
        self.rate_combo.grid(row=1, column=1, padx=5, pady=5)

        tk.Button(root, text="Thêm cấu hình", command=self.add_config).grid(row=2, column=1, pady=10)

        self.listbox = tk.Listbox(root, width=60, height=10)
        self.listbox.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

        self.refresh_listbox()

        #Bắt sự kiện minimize
        self.root.protocol("WM_DELETE_WINDOW", self.on_exit)
        self.root.bind("<Unmap>", self.on_minimize)

        self.icon = self.create_tray_icon()

        # Thread giám sát
        self.monitor_thread = threading.Thread(target=self.monitor_apps, daemon=True)
        self.monitor_thread.start()

    def choose_file(self):
        file = filedialog.askopenfilename(filetypes=[("EXE files", "*.exe")])
        if file:
            self.app_entry.delete(0, tk.END)
            self.app_entry.insert(0, os.path.basename(file))

    def add_config(self):
        app = self.app_entry.get().strip()
        rate = self.rate_var.get().strip()

        if not app or not rate:
            messagebox.showwarning("Lỗi", "Vui lòng chọn ứng dụng và refresh rate")
            return

        self.config[app] = int(rate)
        save_config(self.config)
        self.refresh_listbox()

    def refresh_listbox(self):
        self.listbox.delete(0, tk.END)
        for app, rate in self.config.items():
            self.listbox.insert(tk.END, f"{app} -> {rate} Hz")

    def monitor_apps(self):
        default_rate = None
        current_rate = None

        # Lấy refresh rate mặc định
        devmode = DEVMODE()
        devmode.dmSize = ctypes.sizeof(DEVMODE)
        if user32.EnumDisplaySettingsW(None, ENUM_CURRENT_SETTINGS, ctypes.byref(devmode)):
            default_rate = devmode.dmDisplayFrequency

        while True:
            running = [p.name() for p in psutil.process_iter()]
            applied = False
            for app, rate in self.config.items():
                if app in running:
                    if current_rate != rate:
                        set_refresh_rate(rate)
                        current_rate = rate
                    applied = True
                    break

            if not applied and current_rate != default_rate:
                set_refresh_rate(default_rate)
                current_rate = default_rate

            time.sleep(3)
    
    def create_tray_icon(self):
        # Tạo icon cho system tray
        image = Image.new('RGB', (64, 64), color=(0, 100, 200))
        draw = ImageDraw.Draw(image)
        draw.rectangle((16, 16, 48, 48), fill=(255, 255, 255))

        menu = pystray.Menu(
            pystray.MenuItem('Mở lại', self.show_window),
            pystray.MenuItem('Thoát', self.on_exit)
        )
        return pystray.Icon("RefreshApp", image, "Auto Refresh Rate Manager", menu)
    
    def on_minimize(self, event=None):
        if self.root.state() == "iconic":
            self.root.withdraw()
            if not self.icon_running:
                threading.Thread(target=self.icon.run, daemon=True).start()
                self.icon_running = True

    def show_window(self, icon=None, item=None):
        self.root.after(0, self.root.deiconify)
        self.root.after(0, self.root.state, "normal")
        self.root.after(0, self.root.lift)


    def on_exit(self, icon=None, item=None):
        if self.icon_running:
            self.icon.stop()
            self.icon_running = False
        self.root.destroy()

# --------- Run app ---------
if __name__ == "__main__":
    root = tk.Tk()
    app = RefreshApp(root)
    root.mainloop()
