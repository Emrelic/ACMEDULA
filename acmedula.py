"""
ACS - Medula Oturum Koruma Programi
Medula sisteminden düşmemeyi sağlamak için otomatik tıklama yapar.
Sistem düşerse otomatik olarak tekrar giriş yapar.
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import threading
import time
import pyautogui
from datetime import datetime
import pystray
from PIL import Image, ImageDraw
import subprocess
import ctypes
from ctypes import wintypes

# PyAutoGUI güvenlik ayarları
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULT_CONFIG = {
    "click_points": [
        {"name": "e-Reçete Sorgu", "x": 59, "y": 197, "enabled": True},
        {"name": "Title Bar 1", "x": 665, "y": 14, "enabled": True},
        {"name": "Title Bar 2", "x": 768, "y": 8, "enabled": True}
    ],
    "interval_minutes": 1,
    "interval_seconds": 0,
    "click_delay_ms": 500,
    "start_minimized": False,
    "auto_start": False,
    "auto_relogin": True,
    "check_interval_seconds": 30,
    "login_settings": {
        "desktop_exe_x": 1850,
        "desktop_exe_y": 869,
        "username_x": 969,
        "username_y": 485,
        "password_x": 952,
        "password_y": 534,
        "login_button_x": 955,
        "login_button_y": 577,
        "window_title_contains": "MEDULA",
        "login_window_title": "BotanikEOS",
        "wait_after_exe_click": 5,
        "wait_after_login": 10
    }
}


def safe_typewrite(text):
    """Türkçe karakter destekli metin yazma"""
    for char in text:
        try:
            pyautogui.typewrite(char, interval=0.05)
        except Exception:
            # typewrite ASCII dışı karakterleri desteklemez, hotkey ile yaz
            try:
                import pyperclip
                old_clipboard = pyperclip.paste()
                pyperclip.copy(char)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.05)
                pyperclip.copy(old_clipboard)
            except Exception:
                # pyperclip yoksa ctypes ile clipboard kullan
                try:
                    ctypes.windll.user32.OpenClipboard(0)
                    ctypes.windll.user32.EmptyClipboard()
                    cf_unicode = 13
                    data = char.encode('utf-16-le') + b'\x00\x00'
                    h_mem = ctypes.windll.kernel32.GlobalAlloc(0x0042, len(data))
                    p_mem = ctypes.windll.kernel32.GlobalLock(h_mem)
                    ctypes.memmove(p_mem, data, len(data))
                    ctypes.windll.kernel32.GlobalUnlock(h_mem)
                    ctypes.windll.user32.SetClipboardData(cf_unicode, h_mem)
                    ctypes.windll.user32.CloseClipboard()
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.05)
                except Exception:
                    pass


class LoginDialog(tk.Toplevel):
    """Kullanıcı adı ve şifre giriş penceresi"""

    def __init__(self, parent):
        super().__init__(parent)
        self.title("ACS - Giriş Bilgileri")
        self.geometry("400x300")
        self.resizable(False, False)

        # Modal yap
        self.transient(parent)
        self.grab_set()

        # Ortala
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 400) // 2
        y = (self.winfo_screenheight() - 300) // 2
        self.geometry(f"+{x}+{y}")

        self.username = None
        self.password = None
        self.result = False

        self.create_widgets()

        # Kapatma olayı
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def create_widgets(self):
        # Başlık
        title_label = ttk.Label(self, text="Medula Giriş Bilgileri",
                                font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=20)

        info_label = ttk.Label(self, text="Bu bilgiler sadece bu oturum için hafızada tutulur.\n"
                                          "Sistem düştüğünde otomatik giriş için kullanılır.",
                               font=("Segoe UI", 9), foreground="gray")
        info_label.pack(pady=5)

        # Form
        form_frame = ttk.Frame(self)
        form_frame.pack(pady=20, padx=30, fill=tk.X)

        # Kullanıcı adı
        ttk.Label(form_frame, text="Kullanıcı Adı:", width=15).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(form_frame, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=0, column=1, pady=5)
        self.username_entry.focus()

        # Şifre
        ttk.Label(form_frame, text="Şifre:", width=15).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(form_frame, textvariable=self.password_var, width=30, show="*")
        self.password_entry.grid(row=1, column=1, pady=5)

        # Butonlar
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=20)

        ttk.Button(btn_frame, text="Tamam", command=self.on_ok, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="İptal", command=self.on_cancel, width=15).pack(side=tk.LEFT, padx=10)

        # Enter tuşu ile onay
        self.bind('<Return>', lambda e: self.on_ok())

    def on_ok(self):
        self.username = self.username_var.get().strip()
        self.password = self.password_var.get()

        if not self.username or not self.password:
            messagebox.showwarning("Uyarı", "Kullanıcı adı ve şifre boş olamaz!", parent=self)
            return

        self.result = True
        self.destroy()

    def on_cancel(self):
        self.result = False
        self.destroy()


class ACMedulaApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ACS - Medula Oturum Koruma")
        self.root.geometry("750x700")
        self.root.resizable(True, True)

        # Giriş bilgileri (sadece hafızada - şifrelenmemiş)
        self.medula_username = None
        self.medula_password = None

        # Icon oluştur
        self.icon_image = self.create_icon_image()

        # Değişkenler
        self.is_running = False
        self.click_thread = None
        self.monitor_thread = None
        self.stop_event = threading.Event()
        self.config = self.load_config()
        self.tray_icon = None
        self.relogin_count = 0
        self.last_relogin_time = None

        # Login kilidi - click ve monitor çakışmasını önler
        self.login_lock = threading.Lock()
        self.is_logging_in = False

        # Başarısız giriş sayacı (backoff için)
        self.consecutive_login_failures = 0
        self.max_backoff_minutes = 10

        # GUI oluştur
        self.create_gui()

        # Kapatma olayı
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Başlangıçta önce Medula kontrolü yap
        self.root.after(100, self.startup_check)

    def startup_check(self):
        """Başlangıçta Medula kontrolü yap"""
        self.log("Başlangıç kontrolü yapılıyor...")

        # Medula açık mı kontrol et
        is_running, windows = self.is_medula_running()

        if is_running:
            window_count = len(windows)
            self.log(f"Medula açık bulundu: {window_count} pencere")
            self.medula_status_var.set(f"Medula Durumu: AÇIK ({window_count} pencere)")

            # Giriş bilgilerini iste (sistem düşerse lazım olacak)
            self.ask_credentials()

            # e-Reçete butonuna bas ve çalışmaya başla
            self.log("e-Reçete butonuna tıklanıyor...")
            self.click_erecete_button()

            # Otomatik korumayı başlat
            self.root.after(1000, self.start_clicking)
        else:
            self.log("Medula kapalı! Giriş yapılacak...")
            self.medula_status_var.set("Medula Durumu: KAPALI")

            # Önce giriş bilgilerini al
            self.ask_credentials()

            # Giriş bilgileri alındıysa otomatik giriş yap
            if self.medula_username and self.medula_password:
                self.log("Otomatik giriş başlatılıyor...")
                self.root.after(500, self.perform_login_and_start)
            else:
                self.log("Giriş bilgileri olmadan devam edilemiyor!")

    def click_erecete_button(self):
        """e-Reçete butonuna tıkla"""
        try:
            # Config'den e-Reçete buton koordinatlarını al
            for point in self.config["click_points"]:
                if "e-Reçete" in point.get("name", "") or "e-recete" in point.get("name", "").lower():
                    if point.get("enabled", True):
                        pyautogui.click(point["x"], point["y"])
                        self.log(f"e-Reçete butonuna tıklandı: ({point['x']}, {point['y']})")
                        return
            # Varsayılan e-Reçete koordinatları
            erecete_point = self.config["click_points"][0] if self.config["click_points"] else None
            if erecete_point:
                pyautogui.click(erecete_point["x"], erecete_point["y"])
                self.log(f"İlk noktaya tıklandı: {erecete_point['name']}")
        except Exception as e:
            self.log(f"e-Reçete tıklama hatası: {e}")

    def perform_login_and_start(self):
        """Giriş yap ve korumayı başlat"""
        def login_thread():
            success = self.perform_login()
            if success:
                self.root.after(0, lambda: self.log("Giriş başarılı! Koruma başlatılıyor..."))
                self.root.after(1000, self.click_erecete_button)
                self.root.after(2000, self.start_clicking)
            else:
                self.root.after(0, lambda: self.log("Giriş başarısız! Manuel giriş yapın."))

        threading.Thread(target=login_thread, daemon=True).start()

    def ask_credentials(self):
        """Kullanıcı adı ve şifre iste"""
        dialog = LoginDialog(self.root)
        self.root.wait_window(dialog)

        if dialog.result:
            self.medula_username = dialog.username
            self.medula_password = dialog.password
            self.log(f"Giriş bilgileri alındı: {self.medula_username}")
            self.credentials_label.config(text=f"Kullanıcı: {self.medula_username}", foreground="green")
        else:
            self.log("Giriş bilgileri girilmedi")
            self.credentials_label.config(text="Kullanıcı: Girilmedi", foreground="red")

    def create_icon_image(self):
        """Tray icon için resim oluştur"""
        img = Image.new('RGB', (64, 64), color=(0, 120, 215))
        draw = ImageDraw.Draw(img)
        draw.ellipse([8, 8, 56, 56], fill=(255, 255, 255))
        draw.text((22, 20), "M", fill=(0, 120, 215))
        return img

    def load_config(self):
        """Ayarları yükle"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # Eksik anahtarları varsayılan değerlerle doldur
                    for key, value in DEFAULT_CONFIG.items():
                        if key not in config:
                            config[key] = value
                    # Login settings için de kontrol
                    if "login_settings" not in config:
                        config["login_settings"] = DEFAULT_CONFIG["login_settings"].copy()
                    else:
                        for key, value in DEFAULT_CONFIG["login_settings"].items():
                            if key not in config["login_settings"]:
                                config["login_settings"][key] = value
                    return config
            except:
                pass
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        """Ayarları kaydet"""
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            messagebox.showerror("Hata", f"Ayarlar kaydedilemedi: {e}")
            return False

    def create_gui(self):
        """Ana GUI'yi oluştur"""
        # Ana notebook (tab kontrolü)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Ana Sayfa Tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="  Ana Sayfa  ")
        self.create_main_tab()

        # Ayarlar Tab
        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="  Ayarlar  ")
        self.create_settings_tab()

        # Tıklama Noktaları Tab
        self.points_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.points_frame, text="  Tiklama Noktalari  ")
        self.create_points_tab()

        # Giriş Ayarları Tab
        self.login_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.login_frame, text="  Giris Ayarlari  ")
        self.create_login_tab()

        # Log Tab
        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="  Log  ")
        self.create_log_tab()

    def create_main_tab(self):
        """Ana sayfa tab'ını oluştur"""
        # Kullanıcı bilgisi
        cred_frame = ttk.LabelFrame(self.main_frame, text="Giriş Bilgileri", padding=10)
        cred_frame.pack(fill=tk.X, padx=10, pady=5)

        self.credentials_label = ttk.Label(cred_frame, text="Kullanıcı: Bekleniyor...",
                                           font=("Segoe UI", 10))
        self.credentials_label.pack(side=tk.LEFT)

        ttk.Button(cred_frame, text="Bilgileri Değiştir",
                   command=self.ask_credentials).pack(side=tk.RIGHT)

        # Durum çerçevesi
        status_frame = ttk.LabelFrame(self.main_frame, text="Durum", padding=20)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_label = ttk.Label(status_frame, text="DURDURULDU",
                                       font=("Segoe UI", 24, "bold"), foreground="red")
        self.status_label.pack(pady=10)

        self.info_label = ttk.Label(status_frame, text="Başlatmak için butona tıklayın",
                                     font=("Segoe UI", 10))
        self.info_label.pack(pady=5)

        # Medula durumu
        self.medula_status_var = tk.StringVar(value="Medula Durumu: Kontrol edilmedi")
        ttk.Label(status_frame, textvariable=self.medula_status_var,
                  font=("Segoe UI", 10)).pack(pady=5)

        # Thread durumu
        self.thread_status_var = tk.StringVar(value="Thread Durumu: -")
        ttk.Label(status_frame, textvariable=self.thread_status_var,
                  font=("Segoe UI", 9), foreground="gray").pack(pady=2)

        # Kontrol butonları
        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_btn = ttk.Button(btn_frame, text="BAŞLAT",
                                     command=self.start_clicking, width=20)
        self.start_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.stop_btn = ttk.Button(btn_frame, text="DURDUR",
                                    command=self.stop_clicking, width=20, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.minimize_btn = ttk.Button(btn_frame, text="Simge Durumuna Küçült",
                                        command=self.minimize_to_tray, width=25)
        self.minimize_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        # Manuel kontrol butonları
        manual_frame = ttk.LabelFrame(self.main_frame, text="Manuel Kontrol", padding=10)
        manual_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(manual_frame, text="Medula Durumunu Kontrol Et",
                   command=self.manual_check_medula).pack(side=tk.LEFT, padx=5)
        ttk.Button(manual_frame, text="Manuel Giriş Yap",
                   command=self.manual_login).pack(side=tk.LEFT, padx=5)
        ttk.Button(manual_frame, text="Medula'yı Kapat (Taskkill)",
                   command=self.kill_medula).pack(side=tk.LEFT, padx=5)

        # İstatistikler
        stats_frame = ttk.LabelFrame(self.main_frame, text="İstatistikler", padding=10)
        stats_frame.pack(fill=tk.X, padx=10, pady=5)

        self.click_count_var = tk.StringVar(value="Toplam Tıklama: 0")
        ttk.Label(stats_frame, textvariable=self.click_count_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.last_click_var = tk.StringVar(value="Son Tıklama: -")
        ttk.Label(stats_frame, textvariable=self.last_click_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.next_click_var = tk.StringVar(value="Sonraki Tıklama: -")
        ttk.Label(stats_frame, textvariable=self.next_click_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.relogin_count_var = tk.StringVar(value="Otomatik Giriş Sayısı: 0")
        ttk.Label(stats_frame, textvariable=self.relogin_count_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.click_count = 0

        # Mevcut ayarlar özeti
        summary_frame = ttk.LabelFrame(self.main_frame, text="Aktif Ayarlar", padding=10)
        summary_frame.pack(fill=tk.X, padx=10, pady=5)

        self.summary_label = ttk.Label(summary_frame, text="", font=("Segoe UI", 9))
        self.summary_label.pack(anchor=tk.W)
        self.update_summary()

    def create_settings_tab(self):
        """Ayarlar tab'ını oluştur"""
        # Zamanlama ayarları
        time_frame = ttk.LabelFrame(self.settings_frame, text="Zamanlama Ayarları", padding=10)
        time_frame.pack(fill=tk.X, padx=10, pady=10)

        # Dakika
        row1 = ttk.Frame(time_frame)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="Tıklama Aralığı (Dakika):", width=30).pack(side=tk.LEFT)
        self.minutes_var = tk.StringVar(value=str(self.config["interval_minutes"]))
        self.minutes_spin = ttk.Spinbox(row1, from_=0, to=60, width=10,
                                         textvariable=self.minutes_var)
        self.minutes_spin.pack(side=tk.LEFT, padx=5)

        # Saniye
        row2 = ttk.Frame(time_frame)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="Tıklama Aralığı (Saniye):", width=30).pack(side=tk.LEFT)
        self.seconds_var = tk.StringVar(value=str(self.config["interval_seconds"]))
        self.seconds_spin = ttk.Spinbox(row2, from_=0, to=59, width=10,
                                         textvariable=self.seconds_var)
        self.seconds_spin.pack(side=tk.LEFT, padx=5)

        # Tıklamalar arası gecikme
        row3 = ttk.Frame(time_frame)
        row3.pack(fill=tk.X, pady=5)
        ttk.Label(row3, text="Tıklamalar Arası Gecikme (ms):", width=30).pack(side=tk.LEFT)
        self.delay_var = tk.StringVar(value=str(self.config["click_delay_ms"]))
        self.delay_spin = ttk.Spinbox(row3, from_=100, to=5000, width=10,
                                       textvariable=self.delay_var, increment=100)
        self.delay_spin.pack(side=tk.LEFT, padx=5)

        # Kontrol aralığı
        row4 = ttk.Frame(time_frame)
        row4.pack(fill=tk.X, pady=5)
        ttk.Label(row4, text="Sistem Kontrol Aralığı (Saniye):", width=30).pack(side=tk.LEFT)
        self.check_interval_var = tk.StringVar(value=str(self.config.get("check_interval_seconds", 30)))
        ttk.Spinbox(row4, from_=10, to=300, width=10,
                    textvariable=self.check_interval_var).pack(side=tk.LEFT, padx=5)

        # Otomatik giriş ayarları
        auto_frame = ttk.LabelFrame(self.settings_frame, text="Otomatik Giriş Ayarları", padding=10)
        auto_frame.pack(fill=tk.X, padx=10, pady=10)

        self.auto_relogin_var = tk.BooleanVar(value=self.config.get("auto_relogin", True))
        ttk.Checkbutton(auto_frame, text="Sistem düşerse otomatik giriş yap",
                        variable=self.auto_relogin_var).pack(anchor=tk.W)

        # Başlangıç ayarları
        startup_frame = ttk.LabelFrame(self.settings_frame, text="Başlangıç Ayarları", padding=10)
        startup_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_minimized_var = tk.BooleanVar(value=self.config.get("start_minimized", False))
        ttk.Checkbutton(startup_frame, text="Başlangıçta simge durumuna küçült",
                        variable=self.start_minimized_var).pack(anchor=tk.W)

        self.auto_start_var = tk.BooleanVar(value=self.config.get("auto_start", False))
        ttk.Checkbutton(startup_frame, text="Başlangıçta otomatik başlat",
                        variable=self.auto_start_var).pack(anchor=tk.W)

        # Kaydet butonu
        btn_frame = ttk.Frame(self.settings_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=20)

        ttk.Button(btn_frame, text="Ayarları Kaydet",
                   command=self.save_settings, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Varsayılana Döndür",
                   command=self.reset_settings, width=20).pack(side=tk.LEFT, padx=5)

    def create_points_tab(self):
        """Tıklama noktaları tab'ını oluştur"""
        # Açıklama
        desc_label = ttk.Label(self.points_frame,
                               text="Tıklanacak noktaları buradan yönetebilirsiniz. "
                                    "Sıra ile belirtilen noktalara tıklama yapılır.",
                               wraplength=700)
        desc_label.pack(padx=10, pady=10)

        # Liste çerçevesi
        list_frame = ttk.Frame(self.points_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview
        columns = ("enabled", "name", "x", "y")
        self.points_tree = ttk.Treeview(list_frame, columns=columns, show="headings", height=10)

        self.points_tree.heading("enabled", text="Aktif")
        self.points_tree.heading("name", text="İsim")
        self.points_tree.heading("x", text="X Koordinatı")
        self.points_tree.heading("y", text="Y Koordinatı")

        self.points_tree.column("enabled", width=60, anchor=tk.CENTER)
        self.points_tree.column("name", width=200)
        self.points_tree.column("x", width=100, anchor=tk.CENTER)
        self.points_tree.column("y", width=100, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.points_tree.yview)
        self.points_tree.configure(yscrollcommand=scrollbar.set)

        self.points_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.refresh_points_list()

        # Düzenleme çerçevesi
        edit_frame = ttk.LabelFrame(self.points_frame, text="Nokta Düzenle / Ekle", padding=10)
        edit_frame.pack(fill=tk.X, padx=10, pady=10)

        # Düzenleme alanları
        row1 = ttk.Frame(edit_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="İsim:", width=15).pack(side=tk.LEFT)
        self.point_name_var = tk.StringVar()
        ttk.Entry(row1, textvariable=self.point_name_var, width=30).pack(side=tk.LEFT, padx=5)

        row2 = ttk.Frame(edit_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="X Koordinatı:", width=15).pack(side=tk.LEFT)
        self.point_x_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.point_x_var, width=10).pack(side=tk.LEFT, padx=5)

        ttk.Label(row2, text="Y Koordinatı:", width=15).pack(side=tk.LEFT)
        self.point_y_var = tk.StringVar()
        ttk.Entry(row2, textvariable=self.point_y_var, width=10).pack(side=tk.LEFT, padx=5)

        self.point_enabled_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(row2, text="Aktif", variable=self.point_enabled_var).pack(side=tk.LEFT, padx=10)

        # Butonlar
        btn_frame = ttk.Frame(edit_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        ttk.Button(btn_frame, text="Yeni Ekle", command=self.add_point, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Güncelle", command=self.update_point, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Sil", command=self.delete_point, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Yukarı", command=self.move_point_up, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Aşağı", command=self.move_point_down, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(btn_frame, text="Mouse Pozisyonu Al", command=self.capture_mouse_position,
                   width=18).pack(side=tk.LEFT, padx=3)

        # Seçim olayı
        self.points_tree.bind('<<TreeviewSelect>>', self.on_point_select)

    def create_login_tab(self):
        """Giriş ayarları tab'ını oluştur"""
        desc_label = ttk.Label(self.login_frame,
                               text="Medula giriş ekranı koordinatları. "
                                    "Sistem düştüğünde bu koordinatlar kullanılarak otomatik giriş yapılır.",
                               wraplength=700)
        desc_label.pack(padx=10, pady=10)

        login_settings = self.config.get("login_settings", DEFAULT_CONFIG["login_settings"])

        # Desktop exe
        exe_frame = ttk.LabelFrame(self.login_frame, text="Masaüstü Kısayolu", padding=10)
        exe_frame.pack(fill=tk.X, padx=10, pady=5)

        row1 = ttk.Frame(exe_frame)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="X:", width=5).pack(side=tk.LEFT)
        self.exe_x_var = tk.StringVar(value=str(login_settings["desktop_exe_x"]))
        ttk.Entry(row1, textvariable=self.exe_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="Y:", width=5).pack(side=tk.LEFT)
        self.exe_y_var = tk.StringVar(value=str(login_settings["desktop_exe_y"]))
        ttk.Entry(row1, textvariable=self.exe_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row1, text="Pozisyon Al",
                   command=lambda: self.capture_login_position("exe")).pack(side=tk.LEFT, padx=10)

        # Kullanıcı adı
        user_frame = ttk.LabelFrame(self.login_frame, text="Kullanıcı Adı Alanı", padding=10)
        user_frame.pack(fill=tk.X, padx=10, pady=5)

        row2 = ttk.Frame(user_frame)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="X:", width=5).pack(side=tk.LEFT)
        self.user_x_var = tk.StringVar(value=str(login_settings["username_x"]))
        ttk.Entry(row2, textvariable=self.user_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="Y:", width=5).pack(side=tk.LEFT)
        self.user_y_var = tk.StringVar(value=str(login_settings["username_y"]))
        ttk.Entry(row2, textvariable=self.user_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row2, text="Pozisyon Al",
                   command=lambda: self.capture_login_position("user")).pack(side=tk.LEFT, padx=10)

        # Şifre
        pass_frame = ttk.LabelFrame(self.login_frame, text="Şifre Alanı", padding=10)
        pass_frame.pack(fill=tk.X, padx=10, pady=5)

        row3 = ttk.Frame(pass_frame)
        row3.pack(fill=tk.X, pady=3)
        ttk.Label(row3, text="X:", width=5).pack(side=tk.LEFT)
        self.pass_x_var = tk.StringVar(value=str(login_settings["password_x"]))
        ttk.Entry(row3, textvariable=self.pass_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row3, text="Y:", width=5).pack(side=tk.LEFT)
        self.pass_y_var = tk.StringVar(value=str(login_settings["password_y"]))
        ttk.Entry(row3, textvariable=self.pass_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row3, text="Pozisyon Al",
                   command=lambda: self.capture_login_position("pass")).pack(side=tk.LEFT, padx=10)

        # Giriş butonu
        btn_login_frame = ttk.LabelFrame(self.login_frame, text="Giriş Butonu", padding=10)
        btn_login_frame.pack(fill=tk.X, padx=10, pady=5)

        row4 = ttk.Frame(btn_login_frame)
        row4.pack(fill=tk.X, pady=3)
        ttk.Label(row4, text="X:", width=5).pack(side=tk.LEFT)
        self.login_btn_x_var = tk.StringVar(value=str(login_settings["login_button_x"]))
        ttk.Entry(row4, textvariable=self.login_btn_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row4, text="Y:", width=5).pack(side=tk.LEFT)
        self.login_btn_y_var = tk.StringVar(value=str(login_settings["login_button_y"]))
        ttk.Entry(row4, textvariable=self.login_btn_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row4, text="Pozisyon Al",
                   command=lambda: self.capture_login_position("login")).pack(side=tk.LEFT, padx=10)

        # Pencere başlıkları
        title_frame = ttk.LabelFrame(self.login_frame, text="Pencere Başlıkları", padding=10)
        title_frame.pack(fill=tk.X, padx=10, pady=5)

        row5 = ttk.Frame(title_frame)
        row5.pack(fill=tk.X, pady=3)
        ttk.Label(row5, text="Ana Pencere İçerir:", width=20).pack(side=tk.LEFT)
        self.window_title_var = tk.StringVar(value=login_settings["window_title_contains"])
        ttk.Entry(row5, textvariable=self.window_title_var, width=30).pack(side=tk.LEFT, padx=5)

        row6 = ttk.Frame(title_frame)
        row6.pack(fill=tk.X, pady=3)
        ttk.Label(row6, text="Giriş Penceresi İçerir:", width=20).pack(side=tk.LEFT)
        self.login_window_var = tk.StringVar(value=login_settings["login_window_title"])
        ttk.Entry(row6, textvariable=self.login_window_var, width=30).pack(side=tk.LEFT, padx=5)

        # Bekleme süreleri
        wait_frame = ttk.LabelFrame(self.login_frame, text="Bekleme Süreleri (Saniye)", padding=10)
        wait_frame.pack(fill=tk.X, padx=10, pady=5)

        row7 = ttk.Frame(wait_frame)
        row7.pack(fill=tk.X, pady=3)
        ttk.Label(row7, text="EXE tıklamasından sonra:", width=25).pack(side=tk.LEFT)
        self.wait_exe_var = tk.StringVar(value=str(login_settings["wait_after_exe_click"]))
        ttk.Spinbox(row7, from_=1, to=30, width=10, textvariable=self.wait_exe_var).pack(side=tk.LEFT, padx=5)

        row8 = ttk.Frame(wait_frame)
        row8.pack(fill=tk.X, pady=3)
        ttk.Label(row8, text="Giriş sonrası bekleme:", width=25).pack(side=tk.LEFT)
        self.wait_login_var = tk.StringVar(value=str(login_settings["wait_after_login"]))
        ttk.Spinbox(row8, from_=1, to=60, width=10, textvariable=self.wait_login_var).pack(side=tk.LEFT, padx=5)

        # Kaydet butonu
        btn_frame = ttk.Frame(self.login_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=20)

        ttk.Button(btn_frame, text="Giriş Ayarlarını Kaydet",
                   command=self.save_login_settings, width=25).pack(side=tk.LEFT, padx=5)

    def create_log_tab(self):
        """Log tab'ını oluştur"""
        self.log_text = tk.Text(self.log_frame, wrap=tk.WORD, state=tk.DISABLED,
                                font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        scrollbar = ttk.Scrollbar(self.log_text, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

        # Temizle butonu
        ttk.Button(self.log_frame, text="Log'u Temizle",
                   command=self.clear_log).pack(pady=5)

    def log(self, message):
        """Log mesajı ekle"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"

        try:
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, full_message)
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)
        except tk.TclError:
            pass

    def clear_log(self):
        """Log'u temizle"""
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def refresh_points_list(self):
        """Tıklama noktaları listesini yenile"""
        for item in self.points_tree.get_children():
            self.points_tree.delete(item)

        for i, point in enumerate(self.config["click_points"]):
            enabled = "Evet" if point.get("enabled", True) else "Hayır"
            self.points_tree.insert("", tk.END, iid=str(i),
                                    values=(enabled, point["name"], point["x"], point["y"]))

    def on_point_select(self, event):
        """Nokta seçildiğinde"""
        selection = self.points_tree.selection()
        if selection:
            idx = int(selection[0])
            point = self.config["click_points"][idx]
            self.point_name_var.set(point["name"])
            self.point_x_var.set(str(point["x"]))
            self.point_y_var.set(str(point["y"]))
            self.point_enabled_var.set(point.get("enabled", True))

    def add_point(self):
        """Yeni nokta ekle"""
        try:
            name = self.point_name_var.get().strip()
            x = int(self.point_x_var.get())
            y = int(self.point_y_var.get())
            enabled = self.point_enabled_var.get()

            if not name:
                messagebox.showwarning("Uyarı", "Lütfen bir isim girin!")
                return

            self.config["click_points"].append({
                "name": name, "x": x, "y": y, "enabled": enabled
            })
            self.save_config()
            self.refresh_points_list()
            self.update_summary()
            self.log(f"Yeni nokta eklendi: {name} ({x}, {y})")
        except ValueError:
            messagebox.showerror("Hata", "Geçerli koordinat değerleri girin!")

    def update_point(self):
        """Seçili noktayı güncelle"""
        selection = self.points_tree.selection()
        if not selection:
            messagebox.showwarning("Uyarı", "Lütfen bir nokta seçin!")
            return

        try:
            idx = int(selection[0])
            name = self.point_name_var.get().strip()
            x = int(self.point_x_var.get())
            y = int(self.point_y_var.get())
            enabled = self.point_enabled_var.get()

            if not name:
                messagebox.showwarning("Uyarı", "Lütfen bir isim girin!")
                return

            self.config["click_points"][idx] = {
                "name": name, "x": x, "y": y, "enabled": enabled
            }
            self.save_config()
            self.refresh_points_list()
            self.update_summary()
            self.log(f"Nokta güncellendi: {name} ({x}, {y})")
        except ValueError:
            messagebox.showerror("Hata", "Geçerli koordinat değerleri girin!")

    def delete_point(self):
        """Seçili noktayı sil"""
        selection = self.points_tree.selection()
        if not selection:
            messagebox.showwarning("Uyarı", "Lütfen bir nokta seçin!")
            return

        idx = int(selection[0])
        name = self.config["click_points"][idx]["name"]

        if messagebox.askyesno("Onay", f"'{name}' noktasını silmek istediğinize emin misiniz?"):
            del self.config["click_points"][idx]
            self.save_config()
            self.refresh_points_list()
            self.update_summary()
            self.log(f"Nokta silindi: {name}")

    def move_point_up(self):
        """Seçili noktayı yukarı taşı"""
        selection = self.points_tree.selection()
        if not selection:
            return

        idx = int(selection[0])
        if idx > 0:
            self.config["click_points"][idx], self.config["click_points"][idx-1] = \
                self.config["click_points"][idx-1], self.config["click_points"][idx]
            self.save_config()
            self.refresh_points_list()
            self.points_tree.selection_set(str(idx-1))

    def move_point_down(self):
        """Seçili noktayı aşağı taşı"""
        selection = self.points_tree.selection()
        if not selection:
            return

        idx = int(selection[0])
        if idx < len(self.config["click_points"]) - 1:
            self.config["click_points"][idx], self.config["click_points"][idx+1] = \
                self.config["click_points"][idx+1], self.config["click_points"][idx]
            self.save_config()
            self.refresh_points_list()
            self.points_tree.selection_set(str(idx+1))

    def capture_mouse_position(self):
        """3 saniye sonra mouse pozisyonunu al"""
        self.log("3 saniye içinde mouse pozisyonu alınacak...")
        messagebox.showinfo("Bilgi", "3 saniye sonra mouse pozisyonu alınacak.\n"
                                     "Mouse'u istediğiniz yere götürün!")

        def capture():
            time.sleep(3)
            x, y = pyautogui.position()
            self.root.after(0, lambda: self.set_captured_position(x, y))

        threading.Thread(target=capture, daemon=True).start()

    def set_captured_position(self, x, y):
        """Yakalanan pozisyonu ayarla"""
        self.point_x_var.set(str(x))
        self.point_y_var.set(str(y))
        self.log(f"Mouse pozisyonu yakalandı: ({x}, {y})")
        messagebox.showinfo("Başarılı", f"Mouse pozisyonu: X={x}, Y={y}")

    def capture_login_position(self, field):
        """Giriş ekranı pozisyonunu al"""
        self.log(f"3 saniye içinde {field} pozisyonu alınacak...")
        messagebox.showinfo("Bilgi", "3 saniye sonra mouse pozisyonu alınacak.\n"
                                     "Mouse'u istediğiniz yere götürün!")

        def capture():
            time.sleep(3)
            x, y = pyautogui.position()
            self.root.after(0, lambda: self.set_login_position(field, x, y))

        threading.Thread(target=capture, daemon=True).start()

    def set_login_position(self, field, x, y):
        """Giriş pozisyonunu ayarla"""
        if field == "exe":
            self.exe_x_var.set(str(x))
            self.exe_y_var.set(str(y))
        elif field == "user":
            self.user_x_var.set(str(x))
            self.user_y_var.set(str(y))
        elif field == "pass":
            self.pass_x_var.set(str(x))
            self.pass_y_var.set(str(y))
        elif field == "login":
            self.login_btn_x_var.set(str(x))
            self.login_btn_y_var.set(str(y))

        self.log(f"{field} pozisyonu yakalandı: ({x}, {y})")
        messagebox.showinfo("Başarılı", f"Pozisyon: X={x}, Y={y}")

    def save_login_settings(self):
        """Giriş ayarlarını kaydet"""
        try:
            self.config["login_settings"] = {
                "desktop_exe_x": int(self.exe_x_var.get()),
                "desktop_exe_y": int(self.exe_y_var.get()),
                "username_x": int(self.user_x_var.get()),
                "username_y": int(self.user_y_var.get()),
                "password_x": int(self.pass_x_var.get()),
                "password_y": int(self.pass_y_var.get()),
                "login_button_x": int(self.login_btn_x_var.get()),
                "login_button_y": int(self.login_btn_y_var.get()),
                "window_title_contains": self.window_title_var.get(),
                "login_window_title": self.login_window_var.get(),
                "wait_after_exe_click": int(self.wait_exe_var.get()),
                "wait_after_login": int(self.wait_login_var.get())
            }
            self.save_config()
            self.log("Giriş ayarları kaydedildi")
            messagebox.showinfo("Başarılı", "Giriş ayarları kaydedildi!")
        except ValueError:
            messagebox.showerror("Hata", "Geçerli değerler girin!")

    def save_settings(self):
        """Ayarları kaydet"""
        try:
            self.config["interval_minutes"] = int(self.minutes_var.get())
            self.config["interval_seconds"] = int(self.seconds_var.get())
            self.config["click_delay_ms"] = int(self.delay_var.get())
            self.config["check_interval_seconds"] = int(self.check_interval_var.get())
            self.config["start_minimized"] = self.start_minimized_var.get()
            self.config["auto_start"] = self.auto_start_var.get()
            self.config["auto_relogin"] = self.auto_relogin_var.get()

            if self.config["interval_minutes"] == 0 and self.config["interval_seconds"] == 0:
                messagebox.showwarning("Uyarı", "En az 1 saniye aralık olmalıdır!")
                self.config["interval_seconds"] = 1
                self.seconds_var.set("1")

            self.save_config()
            self.update_summary()
            self.log("Ayarlar kaydedildi")
            messagebox.showinfo("Başarılı", "Ayarlar kaydedildi!")
        except ValueError:
            messagebox.showerror("Hata", "Geçerli sayısal değerler girin!")

    def reset_settings(self):
        """Varsayılan ayarlara döndür"""
        if messagebox.askyesno("Onay", "Tüm ayarlar varsayılana döndürülecek. Emin misiniz?"):
            self.config = DEFAULT_CONFIG.copy()
            self.save_config()

            self.minutes_var.set(str(self.config["interval_minutes"]))
            self.seconds_var.set(str(self.config["interval_seconds"]))
            self.delay_var.set(str(self.config["click_delay_ms"]))
            self.check_interval_var.set(str(self.config["check_interval_seconds"]))
            self.start_minimized_var.set(self.config["start_minimized"])
            self.auto_start_var.set(self.config["auto_start"])
            self.auto_relogin_var.set(self.config["auto_relogin"])

            self.refresh_points_list()
            self.update_summary()
            self.log("Ayarlar varsayılana döndürüldü")

    def update_summary(self):
        """Ayar özetini güncelle"""
        minutes = self.config["interval_minutes"]
        seconds = self.config["interval_seconds"]
        total_seconds = minutes * 60 + seconds

        enabled_points = [p for p in self.config["click_points"] if p.get("enabled", True)]
        auto_relogin = "Açık" if self.config.get("auto_relogin", True) else "Kapalı"

        summary = f"Tıklama Aralığı: {minutes} dk {seconds} sn ({total_seconds} saniye)\n"
        summary += f"Aktif Tıklama Noktası: {len(enabled_points)} adet\n"
        summary += f"Tıklamalar Arası Gecikme: {self.config['click_delay_ms']} ms\n"
        summary += f"Otomatik Giriş: {auto_relogin}"

        self.summary_label.config(text=summary)

    def find_window_by_title(self, title_contains):
        """Belirli başlık içeren pencereyi bul"""
        EnumWindows = ctypes.windll.user32.EnumWindows
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
        GetWindowText = ctypes.windll.user32.GetWindowTextW
        GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
        IsWindowVisible = ctypes.windll.user32.IsWindowVisible

        found_windows = []

        def foreach_window(hwnd, lParam):
            if IsWindowVisible(hwnd):
                length = GetWindowTextLength(hwnd)
                buff = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buff, length + 1)
                title = buff.value
                if title_contains.lower() in title.lower():
                    found_windows.append((hwnd, title))
            return True

        EnumWindows(EnumWindowsProc(foreach_window), 0)
        return found_windows

    def is_medula_running(self):
        """Medula penceresinin açık olup olmadığını kontrol et"""
        # Hem MEDULA hem BotanikEOS ara
        title_contains = self.config["login_settings"]["window_title_contains"]
        login_title = self.config["login_settings"]["login_window_title"]

        windows = self.find_window_by_title(title_contains)
        if not windows:
            windows = self.find_window_by_title(login_title)
        return len(windows) > 0, windows

    def is_login_window_open(self):
        """Giriş penceresinin açık olup olmadığını kontrol et"""
        title_contains = self.config["login_settings"]["login_window_title"]
        windows = self.find_window_by_title(title_contains)
        return len(windows) > 0, windows

    def manual_check_medula(self):
        """Manuel olarak Medula durumunu kontrol et"""
        is_running, windows = self.is_medula_running()
        if is_running:
            window_count = len(windows)
            window_names = [w[1] for w in windows]
            self.medula_status_var.set(f"Medula Durumu: AÇIK ({window_count} pencere)")
            self.log(f"Medula açık: {window_names}")
        else:
            self.medula_status_var.set("Medula Durumu: KAPALI")
            self.log("Medula kapalı!")

    def kill_medula(self):
        """Medula'yı taskkill ile kapat"""
        try:
            self.log("Medula kapatılıyor (taskkill)...")
            # BotanikEOS.exe veya benzeri process'i kapat
            subprocess.run(["taskkill", "/F", "/IM", "BotanikEOS.exe"],
                         capture_output=True, shell=True)
            subprocess.run(["taskkill", "/F", "/IM", "Medula.exe"],
                         capture_output=True, shell=True)
            time.sleep(1)
            self.log("Medula kapatıldı")
            self.medula_status_var.set("Medula Durumu: KAPATILDI")
        except Exception as e:
            self.log(f"Taskkill hatası: {e}")

    def perform_login(self):
        """Otomatik giriş yap"""
        if not self.medula_username or not self.medula_password:
            self.log("HATA: Kullanıcı adı veya şifre girilmemiş!")
            return False

        # Login kilidi al - aynı anda birden fazla login denemesini engelle
        if not self.login_lock.acquire(blocking=False):
            self.log("Giriş zaten devam ediyor, bekleniyor...")
            return False

        self.is_logging_in = True

        login_settings = self.config["login_settings"]

        try:
            self.log("Otomatik giriş başlatılıyor...")

            # 1. Önce mevcut Medula'yı kapat
            self.log("Mevcut Medula kapatılıyor...")
            subprocess.run(["taskkill", "/F", "/IM", "BotanikEOS.exe"],
                         capture_output=True, shell=True)
            time.sleep(2)

            # 2. Masaüstündeki exe'ye çift tıkla
            self.log(f"Masaüstü kısayoluna tıklanıyor ({login_settings['desktop_exe_x']}, {login_settings['desktop_exe_y']})...")
            pyautogui.click(login_settings["desktop_exe_x"], login_settings["desktop_exe_y"])
            time.sleep(0.3)
            pyautogui.doubleClick(login_settings["desktop_exe_x"], login_settings["desktop_exe_y"])

            # 3. Pencere açılmasını bekle
            wait_time = login_settings["wait_after_exe_click"]
            self.log(f"{wait_time} saniye bekleniyor...")
            time.sleep(wait_time)

            # 4. Login penceresi kontrolü
            for attempt in range(5):
                is_login, windows = self.is_login_window_open()
                if is_login:
                    self.log("Giriş penceresi açıldı")
                    break
                time.sleep(1)

            # 5. Kullanıcı adı alanına tıkla ve yaz
            self.log("Kullanıcı adı giriliyor...")
            pyautogui.click(login_settings["username_x"], login_settings["username_y"])
            time.sleep(0.3)
            pyautogui.tripleClick(login_settings["username_x"], login_settings["username_y"])  # Seç
            time.sleep(0.2)
            safe_typewrite(self.medula_username)
            time.sleep(0.3)

            # 6. Şifre alanına tıkla ve yaz
            self.log("Şifre giriliyor...")
            pyautogui.click(login_settings["password_x"], login_settings["password_y"])
            time.sleep(0.3)
            pyautogui.tripleClick(login_settings["password_x"], login_settings["password_y"])  # Seç
            time.sleep(0.2)
            safe_typewrite(self.medula_password)
            time.sleep(0.3)

            # 7. Giriş butonuna tıkla
            self.log("Giriş butonuna tıklanıyor...")
            pyautogui.click(login_settings["login_button_x"], login_settings["login_button_y"])

            # 8. Giriş sonrası bekle
            wait_login = login_settings["wait_after_login"]
            self.log(f"Giriş sonrası {wait_login} saniye bekleniyor...")
            time.sleep(wait_login)

            # 9. Kontrol et
            is_running, _ = self.is_medula_running()
            if is_running:
                self.log("Otomatik giriş BAŞARILI!")
                self.relogin_count += 1
                self.last_relogin_time = datetime.now()
                self.consecutive_login_failures = 0  # Başarılı, sayacı sıfırla
                count = self.relogin_count
                self.root.after(0, lambda: self.relogin_count_var.set(f"Otomatik Giriş Sayısı: {count}"))
                return True
            else:
                self.consecutive_login_failures += 1
                fail_count = self.consecutive_login_failures
                self.log(f"Otomatik giriş başarısız olabilir, kontrol edin (ardışık hata: {fail_count})")
                return False

        except Exception as e:
            self.consecutive_login_failures += 1
            self.log(f"Otomatik giriş hatası: {e}")
            return False
        finally:
            self.is_logging_in = False
            self.login_lock.release()

    def manual_login(self):
        """Manuel giriş yap"""
        if not self.medula_username or not self.medula_password:
            messagebox.showwarning("Uyarı", "Önce giriş bilgilerini girin!")
            self.ask_credentials()
            return

        if messagebox.askyesno("Onay", "Otomatik giriş yapılacak. Devam edilsin mi?"):
            self.consecutive_login_failures = 0  # Manuel giriş, sayacı sıfırla
            threading.Thread(target=self.perform_login, daemon=True).start()

    def start_clicking(self):
        """Tıklamayı başlat"""
        if self.is_running:
            return

        enabled_points = [p for p in self.config["click_points"] if p.get("enabled", True)]
        if not enabled_points:
            messagebox.showwarning("Uyarı", "En az bir aktif tıklama noktası gerekli!")
            return

        self.is_running = True
        self.stop_event.clear()

        self.status_label.config(text="ÇALIŞIYOR", foreground="green")
        self.info_label.config(text="Otomatik koruma aktif")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        self.log("Otomatik koruma başlatıldı")

        # Tıklama thread'i (otomatik yeniden başlatma ile)
        self.click_thread = threading.Thread(target=self._resilient_click_loop, daemon=True)
        self.click_thread.start()

        # Monitor thread'i (otomatik yeniden başlatma ile)
        self.monitor_thread = threading.Thread(target=self._resilient_monitor_loop, daemon=True)
        self.monitor_thread.start()

        # Thread sağlık kontrolcüsü
        self._start_thread_watchdog()

    def stop_clicking(self):
        """Tıklamayı durdur"""
        if not self.is_running:
            return

        self.is_running = False
        self.stop_event.set()

        self.status_label.config(text="DURDURULDU", foreground="red")
        self.info_label.config(text="Başlatmak için butona tıklayın")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.next_click_var.set("Sonraki Tıklama: -")
        self.thread_status_var.set("Thread Durumu: Durduruldu")

        self.log("Otomatik koruma durduruldu")

    def _resilient_click_loop(self):
        """Hata durumunda otomatik yeniden başlayan tıklama döngüsü"""
        while not self.stop_event.is_set():
            try:
                self.click_loop()
            except Exception as e:
                if not self.stop_event.is_set():
                    error_msg = str(e)
                    self.root.after(0, lambda: self.log(f"UYARI: Tıklama thread'i hata aldı: {error_msg}, yeniden başlatılıyor..."))
                    time.sleep(2)

    def _resilient_monitor_loop(self):
        """Hata durumunda otomatik yeniden başlayan monitor döngüsü"""
        while not self.stop_event.is_set():
            try:
                self.monitor_loop()
            except Exception as e:
                if not self.stop_event.is_set():
                    error_msg = str(e)
                    self.root.after(0, lambda: self.log(f"UYARI: Monitor thread'i hata aldı: {error_msg}, yeniden başlatılıyor..."))
                    time.sleep(2)

    def _start_thread_watchdog(self):
        """Thread'lerin hayatta olduğunu periyodik kontrol et"""
        if not self.is_running:
            return

        click_alive = self.click_thread is not None and self.click_thread.is_alive()
        monitor_alive = self.monitor_thread is not None and self.monitor_thread.is_alive()

        status_parts = []
        if click_alive:
            status_parts.append("Tıklama: OK")
        else:
            status_parts.append("Tıklama: OLDU")
        if monitor_alive:
            status_parts.append("Monitor: OK")
        else:
            status_parts.append("Monitor: OLDU")

        status_text = "Thread Durumu: " + " | ".join(status_parts)
        self.thread_status_var.set(status_text)

        # Ölen thread'leri yeniden başlat
        if not click_alive and self.is_running and not self.stop_event.is_set():
            self.log("UYARI: Tıklama thread'i öldü, yeniden başlatılıyor!")
            self.click_thread = threading.Thread(target=self._resilient_click_loop, daemon=True)
            self.click_thread.start()

        if not monitor_alive and self.is_running and not self.stop_event.is_set():
            self.log("UYARI: Monitor thread'i öldü, yeniden başlatılıyor!")
            self.monitor_thread = threading.Thread(target=self._resilient_monitor_loop, daemon=True)
            self.monitor_thread.start()

        # 10 saniyede bir kontrol et
        if self.is_running:
            self.root.after(10000, self._start_thread_watchdog)

    def click_loop(self):
        """Tıklama döngüsü"""
        while not self.stop_event.is_set():
            # Her döngüde config'den güncel ayarları oku
            interval = self.config["interval_minutes"] * 60 + self.config["interval_seconds"]
            delay_ms = self.config["click_delay_ms"]

            # Login sırasında tıklama yapma
            if self.is_logging_in:
                self.root.after(0, lambda: self.log("Login devam ediyor, tıklama atlanıyor"))
                time.sleep(5)
                continue

            # Önce Medula'nın açık olduğunu kontrol et
            is_running, _ = self.is_medula_running()

            if is_running:
                # Aktif noktaları al
                enabled_points = [p for p in self.config["click_points"] if p.get("enabled", True)]

                # Tıklamaları yap
                for point in enabled_points:
                    if self.stop_event.is_set() or self.is_logging_in:
                        break

                    try:
                        pyautogui.click(point["x"], point["y"])
                        self.click_count += 1

                        # GUI güncelle
                        point_name = point["name"]
                        count = self.click_count
                        self.root.after(0, lambda pn=point_name, c=count: self.update_stats(pn))

                        time.sleep(delay_ms / 1000.0)
                    except Exception as e:
                        error_msg = str(e)
                        self.root.after(0, lambda em=error_msg: self.log(f"Tıklama hatası: {em}"))
            else:
                self.root.after(0, lambda: self.log("Medula kapalı, tıklama atlanıyor"))

            # Sonraki tıklamaya kadar bekle
            for i in range(interval):
                if self.stop_event.is_set():
                    break
                remaining = interval - i
                self.root.after(0, lambda r=remaining:
                               self.next_click_var.set(f"Sonraki Tıklama: {r} saniye"))
                time.sleep(1)

    def monitor_loop(self):
        """Sistem durumunu izleme döngüsü"""
        while not self.stop_event.is_set():
            # Her döngüde güncel kontrol aralığını oku
            check_interval = self.config.get("check_interval_seconds", 30)

            # Medula durumunu kontrol et
            is_running, windows = self.is_medula_running()
            window_count = len(windows)

            if is_running:
                self.root.after(0, lambda wc=window_count: self.medula_status_var.set(
                    f"Medula Durumu: AÇIK ({wc} pencere)"))
            else:
                self.root.after(0, lambda: self.medula_status_var.set("Medula Durumu: KAPALI"))
                self.root.after(0, lambda: self.log("UYARI: Medula kapalı tespit edildi!"))

                # Otomatik giriş aktifse giriş yap
                if self.config.get("auto_relogin", True):
                    if self.medula_username and self.medula_password:
                        # Backoff hesapla: ardışık hatada artan bekleme
                        if self.consecutive_login_failures > 0:
                            backoff_seconds = min(
                                self.consecutive_login_failures * 30,
                                self.max_backoff_minutes * 60
                            )
                            fail_count = self.consecutive_login_failures
                            backoff_display = backoff_seconds
                            self.root.after(0, lambda fc=fail_count, bs=backoff_display:
                                self.log(f"Ardışık {fc} başarısız giriş. {bs} saniye bekleniyor..."))
                            for _ in range(backoff_seconds):
                                if self.stop_event.is_set():
                                    break
                                time.sleep(1)
                            if self.stop_event.is_set():
                                break

                        self.root.after(0, lambda: self.log("Otomatik giriş başlatılıyor..."))
                        success = self.perform_login()
                        if success:
                            self.root.after(0, lambda: self.log("Sistem yeniden aktif!"))
                            # Başarılı girişten sonra e-Reçete butonuna tıkla
                            time.sleep(2)
                            try:
                                self.click_erecete_button()
                            except Exception:
                                pass
                        else:
                            self.root.after(0, lambda: self.log("Otomatik giriş başarısız!"))
                    else:
                        self.root.after(0, lambda: self.log("Giriş bilgileri olmadan otomatik giriş yapılamaz!"))

            # Bekle
            for _ in range(check_interval):
                if self.stop_event.is_set():
                    break
                time.sleep(1)

    def update_stats(self, point_name):
        """İstatistikleri güncelle"""
        self.click_count_var.set(f"Toplam Tıklama: {self.click_count}")
        now = datetime.now().strftime("%H:%M:%S")
        self.last_click_var.set(f"Son Tıklama: {now} - {point_name}")
        self.log(f"Tıklama yapıldı: {point_name}")

    def minimize_to_tray(self):
        """System tray'e küçült"""
        self.root.withdraw()

        if self.tray_icon is None:
            menu = pystray.Menu(
                pystray.MenuItem("Göster", self.show_from_tray),
                pystray.MenuItem("Başlat", self.tray_start),
                pystray.MenuItem("Durdur", self.tray_stop),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Çıkış", self.quit_app)
            )

            self.tray_icon = pystray.Icon("ACS", self.icon_image, "ACS", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_from_tray(self, icon=None, item=None):
        """Tray'den geri getir"""
        self.root.after(0, self.root.deiconify)

    def tray_start(self, icon=None, item=None):
        """Tray'den başlat"""
        self.root.after(0, self.start_clicking)

    def tray_stop(self, icon=None, item=None):
        """Tray'den durdur"""
        self.root.after(0, self.stop_clicking)

    def quit_app(self, icon=None, item=None):
        """Uygulamadan çık"""
        self.stop_clicking()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.after(0, self.root.destroy)

    def on_close(self):
        """Pencere kapatıldığında"""
        if messagebox.askyesno("Çıkış", "Programı kapatmak mı yoksa simge durumuna "
                                        "küçültmek mi istiyorsunuz?\n\n"
                                        "Evet = Kapat\nHayır = Simge Durumuna Küçült"):
            self.quit_app()
        else:
            self.minimize_to_tray()

    def run(self):
        """Uygulamayı çalıştır"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ACMedulaApp()
    app.run()
