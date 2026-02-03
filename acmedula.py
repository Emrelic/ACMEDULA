"""
ACMEDULA - Medula Oturum Koruma Programi v3.1
Medula sisteminden düşmemeyi sağlamak için akıllı otomatik tıklama yapar.
Sistem düşerse otomatik olarak tekrar giriş yapar.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import threading
import time
import random
import pyautogui
from datetime import datetime
import pystray
from PIL import Image, ImageDraw
import subprocess
import ctypes
from ctypes import wintypes
from pynput import mouse, keyboard

# PyAutoGUI güvenlik ayarları
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")

DEFAULT_CONFIG = {
    "click_buttons": [
        {"name": "Ana Sayfa (Giriş)", "x": 29, "y": 32, "enabled": True},
        {"name": "e-Reçete Sorgu", "x": 59, "y": 197, "enabled": True}
    ],
    "min_interval_seconds": 60,
    "max_interval_seconds": 120,
    "idle_wait_seconds": 30,
    "start_minimized": False,
    "auto_relogin": True,
    "check_interval_seconds": 30,
    "login_settings": {
        "desktop_exe_x": 1850,
        "desktop_exe_y": 869,
        "username_combobox_x": 969,
        "username_combobox_y": 485,
        "user_selection_method": "index",
        "user_index": 1,
        "password_x": 952,
        "password_y": 534,
        "login_button_x": 955,
        "login_button_y": 577,
        "main_page_button_x": 29,
        "main_page_button_y": 32,
        "window_title_contains": "MEDULA",
        "login_window_title": "BotanikEOS",
        "general_announcements_check_x": 1041,
        "general_announcements_check_y": 217,
        "wait_after_exe_click": 5,
        "wait_after_login": 10
    }
}


class LoginDialog(tk.Toplevel):
    """Kullanıcı adı ve şifre giriş penceresi"""

    def __init__(self, parent, config):
        super().__init__(parent)
        self.title("ACMEDULA - Giriş Bilgileri")
        self.geometry("480x380")
        self.resizable(False, False)
        self.config = config

        # Modal yap
        self.transient(parent)
        self.grab_set()

        # Ortala
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 480) // 2
        y = (self.winfo_screenheight() - 380) // 2
        self.geometry(f"+{x}+{y}")

        self.username = None
        self.password = None
        self.user_selection_method = None
        self.user_index = None
        self.result = False
        self.auto_start = False

        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_cancel)

    def create_widgets(self):
        # Başlık
        title_label = ttk.Label(self, text="Medula Giriş Bilgileri",
                                font=("Segoe UI", 14, "bold"))
        title_label.pack(pady=10)

        info_label = ttk.Label(self, text="Bu bilgiler sadece bu oturum için hafızada tutulur.",
                               font=("Segoe UI", 9), foreground="gray")
        info_label.pack(pady=3)

        # Form
        form_frame = ttk.Frame(self)
        form_frame.pack(pady=10, padx=30, fill=tk.X)

        # Kullanıcı Seçim Yöntemi
        ttk.Label(form_frame, text="Kullanıcı Seçim:", width=18).grid(row=0, column=0, sticky=tk.W, pady=5)

        method_frame = ttk.Frame(form_frame)
        method_frame.grid(row=0, column=1, sticky=tk.W, pady=5)

        ls = self.config.get("login_settings", {})
        current_method = ls.get("user_selection_method", "index")

        self.method_var = tk.StringVar(value=current_method)
        ttk.Radiobutton(method_frame, text="Sıra No", variable=self.method_var,
                        value="index", command=self.on_method_change).pack(side=tk.LEFT)
        ttk.Radiobutton(method_frame, text="İsim Ara", variable=self.method_var,
                        value="name", command=self.on_method_change).pack(side=tk.LEFT, padx=10)

        # Kullanıcı Sıra No
        ttk.Label(form_frame, text="Kullanıcı Sıra No:", width=18).grid(row=1, column=0, sticky=tk.W, pady=5)

        index_frame = ttk.Frame(form_frame)
        index_frame.grid(row=1, column=1, sticky=tk.W, pady=5)

        current_index = ls.get("user_index", 1)
        self.index_var = tk.StringVar(value=str(current_index))
        self.index_spin = ttk.Spinbox(index_frame, from_=1, to=20, width=5, textvariable=self.index_var)
        self.index_spin.pack(side=tk.LEFT)
        ttk.Label(index_frame, text=" (1 = ilk kullanıcı)", foreground="gray").pack(side=tk.LEFT)

        # Kullanıcı Adı (isim araması için)
        ttk.Label(form_frame, text="Kullanıcı Adı:", width=18).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(form_frame, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=2, column=1, sticky=tk.W, pady=5)

        # Şifre
        ttk.Label(form_frame, text="Şifre:", width=18).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(form_frame, textvariable=self.password_var, width=30, show="*")
        self.password_entry.grid(row=3, column=1, sticky=tk.W, pady=5)

        # Açıklama
        desc_frame = ttk.LabelFrame(self, text="Açıklama", padding=10)
        desc_frame.pack(fill=tk.X, padx=30, pady=10)

        desc_text = ("• Sıra No: Combobox'ta kaçıncı kullanıcıyı seçeceğini belirler\n"
                     "  (1 = ilk kullanıcı, 2 = ikinci kullanıcı...)\n"
                     "• İsim Ara: Kullanıcı adını combobox'a yazar ve eşleşeni seçer")
        ttk.Label(desc_frame, text=desc_text, font=("Segoe UI", 9), foreground="gray").pack(anchor=tk.W)

        # Butonlar
        btn_frame = ttk.Frame(self)
        btn_frame.pack(pady=15)

        self.start_btn = ttk.Button(btn_frame, text="Kaydet ve Başlat",
                                     command=self.on_save_and_start, width=16)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="Sadece Kaydet",
                   command=self.on_save_only, width=14).pack(side=tk.LEFT, padx=5)

        ttk.Button(btn_frame, text="İptal",
                   command=self.on_cancel, width=10).pack(side=tk.LEFT, padx=5)

        # Enter ile başlat
        self.bind('<Return>', lambda e: self.on_save_and_start())

        # İlk duruma göre ayarla
        self.on_method_change()
        self.password_entry.focus()

    def on_method_change(self):
        """Seçim yöntemi değiştiğinde"""
        method = self.method_var.get()
        if method == "index":
            self.index_spin.config(state=tk.NORMAL)
            self.username_entry.config(state=tk.DISABLED)
        else:
            self.index_spin.config(state=tk.DISABLED)
            self.username_entry.config(state=tk.NORMAL)
            self.username_entry.focus()

    def validate_input(self):
        self.password = self.password_var.get()
        self.user_selection_method = self.method_var.get()

        if not self.password:
            messagebox.showwarning("Uyarı", "Şifre boş olamaz!", parent=self)
            return False

        if self.user_selection_method == "index":
            try:
                self.user_index = int(self.index_var.get())
                if self.user_index < 1:
                    raise ValueError()
                self.username = None
            except:
                messagebox.showwarning("Uyarı", "Geçerli bir sıra numarası girin!", parent=self)
                return False
        else:
            self.username = self.username_var.get().strip()
            if not self.username:
                messagebox.showwarning("Uyarı", "Kullanıcı adı boş olamaz!", parent=self)
                return False
            self.user_index = None

        return True

    def on_save_and_start(self):
        if self.validate_input():
            self.result = True
            self.auto_start = True
            self.destroy()

    def on_save_only(self):
        if self.validate_input():
            self.result = True
            self.auto_start = False
            self.destroy()

    def on_cancel(self):
        self.result = False
        self.destroy()


class ACMedulaApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("ACMEDULA - Medula Oturum Koruma v3.1")
        self.root.geometry("800x750")
        self.root.resizable(True, True)

        # Giriş bilgileri (sadece hafızada)
        self.medula_username = None
        self.medula_password = None
        self.user_selection_method = "index"
        self.user_index = 1

        # Icon oluştur
        self.icon_image = self.create_icon_image()

        # Değişkenler
        self.is_running = False
        self.main_thread = None
        self.stop_event = threading.Event()
        self.config = self.load_config()
        self.tray_icon = None
        self.relogin_count = 0
        self.click_count = 0

        # Aktivite takibi
        self.last_activity_time = time.time()
        self.mouse_listener = None
        self.keyboard_listener = None
        self.activity_lock = threading.Lock()

        # GUI oluştur
        self.create_gui()

        # Kapatma olayı
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        # Başlangıçta giriş bilgilerini iste
        self.root.after(100, self.ask_credentials)

    def ask_credentials(self):
        """Kullanıcı adı ve şifre iste"""
        dialog = LoginDialog(self.root, self.config)
        self.root.wait_window(dialog)

        if dialog.result:
            self.medula_password = dialog.password
            self.user_selection_method = dialog.user_selection_method
            self.user_index = dialog.user_index
            self.medula_username = dialog.username

            # Config'e kaydet
            self.config["login_settings"]["user_selection_method"] = self.user_selection_method
            if self.user_index:
                self.config["login_settings"]["user_index"] = self.user_index
            self.save_config()

            if self.user_selection_method == "index":
                self.log(f"Giriş bilgileri kaydedildi: {self.user_index}. kullanıcı")
                self.credentials_label.config(
                    text=f"Kullanıcı: {self.user_index}. sıradaki", foreground="green")
            else:
                self.log(f"Giriş bilgileri kaydedildi: {self.medula_username}")
                self.credentials_label.config(
                    text=f"Kullanıcı: {self.medula_username}", foreground="green")

            if dialog.auto_start:
                self.root.after(500, self.start_protection)
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
                    for key, value in DEFAULT_CONFIG.items():
                        if key not in config:
                            config[key] = value
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
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text="  Ana Sayfa  ")
        self.create_main_tab()

        self.settings_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.settings_frame, text="  Ayarlar  ")
        self.create_settings_tab()

        self.login_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.login_frame, text="  Giris Ayarlari  ")
        self.create_login_tab()

        self.log_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.log_frame, text="  Log  ")
        self.create_log_tab()

    def create_main_tab(self):
        """Ana sayfa tab'ını oluştur"""
        cred_frame = ttk.LabelFrame(self.main_frame, text="Giriş Bilgileri", padding=10)
        cred_frame.pack(fill=tk.X, padx=10, pady=5)

        self.credentials_label = ttk.Label(cred_frame, text="Kullanıcı: Bekleniyor...",
                                           font=("Segoe UI", 10))
        self.credentials_label.pack(side=tk.LEFT)

        ttk.Button(cred_frame, text="Bilgileri Değiştir",
                   command=self.ask_credentials).pack(side=tk.RIGHT)

        status_frame = ttk.LabelFrame(self.main_frame, text="Durum", padding=15)
        status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.status_label = ttk.Label(status_frame, text="DURDURULDU",
                                       font=("Segoe UI", 24, "bold"), foreground="red")
        self.status_label.pack(pady=5)

        self.info_label = ttk.Label(status_frame, text="Başlatmak için butona tıklayın",
                                     font=("Segoe UI", 10))
        self.info_label.pack(pady=3)

        self.medula_status_var = tk.StringVar(value="Medula: Kontrol edilmedi")
        ttk.Label(status_frame, textvariable=self.medula_status_var,
                  font=("Segoe UI", 10)).pack(pady=2)

        self.activity_status_var = tk.StringVar(value="Aktivite: -")
        ttk.Label(status_frame, textvariable=self.activity_status_var,
                  font=("Segoe UI", 10)).pack(pady=2)

        self.countdown_var = tk.StringVar(value="Sonraki tıklama: -")
        ttk.Label(status_frame, textvariable=self.countdown_var,
                  font=("Segoe UI", 11, "bold")).pack(pady=5)

        btn_frame = ttk.Frame(self.main_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_btn = ttk.Button(btn_frame, text="BAŞLAT",
                                     command=self.start_protection, width=20)
        self.start_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.stop_btn = ttk.Button(btn_frame, text="DURDUR",
                                    command=self.stop_protection, width=20, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5, pady=5)

        self.minimize_btn = ttk.Button(btn_frame, text="Simge Durumuna Küçült",
                                        command=self.minimize_to_tray, width=22)
        self.minimize_btn.pack(side=tk.RIGHT, padx=5, pady=5)

        manual_frame = ttk.LabelFrame(self.main_frame, text="Manuel Kontrol", padding=10)
        manual_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(manual_frame, text="Sistem Kontrolü",
                   command=self.manual_system_check).pack(side=tk.LEFT, padx=5)
        ttk.Button(manual_frame, text="Manuel Giriş",
                   command=self.manual_login).pack(side=tk.LEFT, padx=5)
        ttk.Button(manual_frame, text="Medula Kapat",
                   command=self.kill_medula).pack(side=tk.LEFT, padx=5)

        stats_frame = ttk.LabelFrame(self.main_frame, text="İstatistikler", padding=10)
        stats_frame.pack(fill=tk.X, padx=10, pady=5)

        self.click_count_var = tk.StringVar(value="Toplam Tıklama: 0")
        ttk.Label(stats_frame, textvariable=self.click_count_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.last_click_var = tk.StringVar(value="Son Tıklama: -")
        ttk.Label(stats_frame, textvariable=self.last_click_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        self.relogin_count_var = tk.StringVar(value="Otomatik Giriş: 0")
        ttk.Label(stats_frame, textvariable=self.relogin_count_var,
                  font=("Segoe UI", 10)).pack(anchor=tk.W)

        summary_frame = ttk.LabelFrame(self.main_frame, text="Aktif Ayarlar", padding=10)
        summary_frame.pack(fill=tk.X, padx=10, pady=5)

        self.summary_label = ttk.Label(summary_frame, text="", font=("Segoe UI", 9))
        self.summary_label.pack(anchor=tk.W)
        self.update_summary()

    def create_settings_tab(self):
        """Ayarlar tab'ını oluştur"""
        time_frame = ttk.LabelFrame(self.settings_frame, text="Tıklama Zamanlaması", padding=10)
        time_frame.pack(fill=tk.X, padx=10, pady=10)

        row1 = ttk.Frame(time_frame)
        row1.pack(fill=tk.X, pady=5)
        ttk.Label(row1, text="Minimum Aralık (Saniye):", width=30).pack(side=tk.LEFT)
        self.min_interval_var = tk.StringVar(value=str(self.config["min_interval_seconds"]))
        ttk.Spinbox(row1, from_=30, to=300, width=10,
                    textvariable=self.min_interval_var).pack(side=tk.LEFT, padx=5)

        row2 = ttk.Frame(time_frame)
        row2.pack(fill=tk.X, pady=5)
        ttk.Label(row2, text="Maksimum Aralık (Saniye):", width=30).pack(side=tk.LEFT)
        self.max_interval_var = tk.StringVar(value=str(self.config["max_interval_seconds"]))
        ttk.Spinbox(row2, from_=60, to=600, width=10,
                    textvariable=self.max_interval_var).pack(side=tk.LEFT, padx=5)

        row3 = ttk.Frame(time_frame)
        row3.pack(fill=tk.X, pady=5)
        ttk.Label(row3, text="Aktivite Sonrası Bekleme (Saniye):", width=30).pack(side=tk.LEFT)
        self.idle_wait_var = tk.StringVar(value=str(self.config["idle_wait_seconds"]))
        ttk.Spinbox(row3, from_=10, to=120, width=10,
                    textvariable=self.idle_wait_var).pack(side=tk.LEFT, padx=5)

        buttons_frame = ttk.LabelFrame(self.settings_frame, text="Tıklanacak Butonlar", padding=10)
        buttons_frame.pack(fill=tk.X, padx=10, pady=10)

        self.btn1_enabled_var = tk.BooleanVar(value=self.config["click_buttons"][0]["enabled"])
        ttk.Checkbutton(buttons_frame, text="Ana Sayfa (Giriş) - X:29, Y:32",
                        variable=self.btn1_enabled_var).pack(anchor=tk.W)

        self.btn2_enabled_var = tk.BooleanVar(value=self.config["click_buttons"][1]["enabled"])
        ttk.Checkbutton(buttons_frame, text="e-Reçete Sorgu - X:59, Y:197",
                        variable=self.btn2_enabled_var).pack(anchor=tk.W)

        auto_frame = ttk.LabelFrame(self.settings_frame, text="Otomatik Giriş", padding=10)
        auto_frame.pack(fill=tk.X, padx=10, pady=10)

        self.auto_relogin_var = tk.BooleanVar(value=self.config.get("auto_relogin", True))
        ttk.Checkbutton(auto_frame, text="Sistem düşerse otomatik giriş yap",
                        variable=self.auto_relogin_var).pack(anchor=tk.W)

        startup_frame = ttk.LabelFrame(self.settings_frame, text="Başlangıç", padding=10)
        startup_frame.pack(fill=tk.X, padx=10, pady=10)

        self.start_minimized_var = tk.BooleanVar(value=self.config.get("start_minimized", False))
        ttk.Checkbutton(startup_frame, text="Başlangıçta simge durumuna küçült",
                        variable=self.start_minimized_var).pack(anchor=tk.W)

        btn_frame = ttk.Frame(self.settings_frame)
        btn_frame.pack(fill=tk.X, padx=10, pady=20)

        ttk.Button(btn_frame, text="Ayarları Kaydet",
                   command=self.save_settings, width=20).pack(side=tk.LEFT, padx=5)

    def create_login_tab(self):
        """Giriş ayarları tab'ını oluştur"""
        canvas = tk.Canvas(self.login_frame)
        scrollbar = ttk.Scrollbar(self.login_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        ls = self.config.get("login_settings", DEFAULT_CONFIG["login_settings"])

        # Kullanıcı Seçimi
        user_sel_frame = ttk.LabelFrame(scrollable_frame, text="Kullanıcı Seçim Ayarları", padding=10)
        user_sel_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(user_sel_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Seçim Yöntemi:", width=18).pack(side=tk.LEFT)
        self.sel_method_var = tk.StringVar(value=ls.get("user_selection_method", "index"))
        ttk.Combobox(row, textvariable=self.sel_method_var, values=["index", "name"],
                     width=15, state="readonly").pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="(index=sıra no, name=isim)", foreground="gray").pack(side=tk.LEFT)

        row = ttk.Frame(user_sel_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Kullanıcı Sıra No:", width=18).pack(side=tk.LEFT)
        self.user_index_var = tk.StringVar(value=str(ls.get("user_index", 1)))
        ttk.Spinbox(row, from_=1, to=20, width=8, textvariable=self.user_index_var).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="(1 = ilk kullanıcı)", foreground="gray").pack(side=tk.LEFT)

        # Masaüstü exe
        exe_frame = ttk.LabelFrame(scrollable_frame, text="Masaüstü Kısayolu", padding=10)
        exe_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(exe_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="X:", width=5).pack(side=tk.LEFT)
        self.exe_x_var = tk.StringVar(value=str(ls["desktop_exe_x"]))
        ttk.Entry(row, textvariable=self.exe_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Y:", width=5).pack(side=tk.LEFT)
        self.exe_y_var = tk.StringVar(value=str(ls["desktop_exe_y"]))
        ttk.Entry(row, textvariable=self.exe_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Pozisyon Al",
                   command=lambda: self.capture_position("exe")).pack(side=tk.LEFT, padx=10)

        # Kullanıcı combobox
        user_frame = ttk.LabelFrame(scrollable_frame, text="Kullanıcı Combobox", padding=10)
        user_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(user_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="X:", width=5).pack(side=tk.LEFT)
        self.user_x_var = tk.StringVar(value=str(ls["username_combobox_x"]))
        ttk.Entry(row, textvariable=self.user_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Y:", width=5).pack(side=tk.LEFT)
        self.user_y_var = tk.StringVar(value=str(ls["username_combobox_y"]))
        ttk.Entry(row, textvariable=self.user_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Pozisyon Al",
                   command=lambda: self.capture_position("user")).pack(side=tk.LEFT, padx=10)

        # Şifre
        pass_frame = ttk.LabelFrame(scrollable_frame, text="Şifre Alanı", padding=10)
        pass_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(pass_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="X:", width=5).pack(side=tk.LEFT)
        self.pass_x_var = tk.StringVar(value=str(ls["password_x"]))
        ttk.Entry(row, textvariable=self.pass_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Y:", width=5).pack(side=tk.LEFT)
        self.pass_y_var = tk.StringVar(value=str(ls["password_y"]))
        ttk.Entry(row, textvariable=self.pass_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Pozisyon Al",
                   command=lambda: self.capture_position("pass")).pack(side=tk.LEFT, padx=10)

        # Giriş butonu
        login_btn_frame = ttk.LabelFrame(scrollable_frame, text="Giriş Butonu (Login)", padding=10)
        login_btn_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(login_btn_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="X:", width=5).pack(side=tk.LEFT)
        self.login_btn_x_var = tk.StringVar(value=str(ls["login_button_x"]))
        ttk.Entry(row, textvariable=self.login_btn_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Y:", width=5).pack(side=tk.LEFT)
        self.login_btn_y_var = tk.StringVar(value=str(ls["login_button_y"]))
        ttk.Entry(row, textvariable=self.login_btn_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Pozisyon Al",
                   command=lambda: self.capture_position("login")).pack(side=tk.LEFT, padx=10)

        # Ana Sayfa butonu
        main_btn_frame = ttk.LabelFrame(scrollable_frame, text="Ana Sayfa Butonu (Medula içi)", padding=10)
        main_btn_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(main_btn_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="X:", width=5).pack(side=tk.LEFT)
        self.main_btn_x_var = tk.StringVar(value=str(ls["main_page_button_x"]))
        ttk.Entry(row, textvariable=self.main_btn_x_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(row, text="Y:", width=5).pack(side=tk.LEFT)
        self.main_btn_y_var = tk.StringVar(value=str(ls["main_page_button_y"]))
        ttk.Entry(row, textvariable=self.main_btn_y_var, width=10).pack(side=tk.LEFT, padx=5)
        ttk.Button(row, text="Pozisyon Al",
                   command=lambda: self.capture_position("main")).pack(side=tk.LEFT, padx=10)

        # Pencere başlıkları
        title_frame = ttk.LabelFrame(scrollable_frame, text="Pencere Başlıkları", padding=10)
        title_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(title_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Ana Pencere İçerir:", width=20).pack(side=tk.LEFT)
        self.window_title_var = tk.StringVar(value=ls["window_title_contains"])
        ttk.Entry(row, textvariable=self.window_title_var, width=25).pack(side=tk.LEFT, padx=5)

        row = ttk.Frame(title_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="Giriş Penceresi İçerir:", width=20).pack(side=tk.LEFT)
        self.login_window_var = tk.StringVar(value=ls["login_window_title"])
        ttk.Entry(row, textvariable=self.login_window_var, width=25).pack(side=tk.LEFT, padx=5)

        # Bekleme süreleri
        wait_frame = ttk.LabelFrame(scrollable_frame, text="Bekleme Süreleri (Saniye)", padding=10)
        wait_frame.pack(fill=tk.X, padx=10, pady=5)

        row = ttk.Frame(wait_frame)
        row.pack(fill=tk.X, pady=3)
        ttk.Label(row, text="EXE sonrası:", width=15).pack(side=tk.LEFT)
        self.wait_exe_var = tk.StringVar(value=str(ls["wait_after_exe_click"]))
        ttk.Spinbox(row, from_=1, to=30, width=8, textvariable=self.wait_exe_var).pack(side=tk.LEFT, padx=5)

        ttk.Label(row, text="Giriş sonrası:", width=15).pack(side=tk.LEFT)
        self.wait_login_var = tk.StringVar(value=str(ls["wait_after_login"]))
        ttk.Spinbox(row, from_=1, to=60, width=8, textvariable=self.wait_login_var).pack(side=tk.LEFT, padx=5)

        btn_frame = ttk.Frame(scrollable_frame)
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

        ttk.Button(self.log_frame, text="Log'u Temizle",
                   command=self.clear_log).pack(pady=5)

    def log(self, message):
        """Log mesajı ekle"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        full_message = f"[{timestamp}] {message}\n"

        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, full_message)
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def clear_log(self):
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def capture_position(self, field):
        self.log(f"3 saniye içinde {field} pozisyonu alınacak...")
        messagebox.showinfo("Bilgi", "3 saniye sonra mouse pozisyonu alınacak.\nMouse'u istediğiniz yere götürün!")

        def capture():
            time.sleep(3)
            x, y = pyautogui.position()
            self.root.after(0, lambda: self.set_position(field, x, y))

        threading.Thread(target=capture, daemon=True).start()

    def set_position(self, field, x, y):
        field_map = {
            "exe": (self.exe_x_var, self.exe_y_var),
            "user": (self.user_x_var, self.user_y_var),
            "pass": (self.pass_x_var, self.pass_y_var),
            "login": (self.login_btn_x_var, self.login_btn_y_var),
            "main": (self.main_btn_x_var, self.main_btn_y_var)
        }
        if field in field_map:
            field_map[field][0].set(str(x))
            field_map[field][1].set(str(y))
        self.log(f"{field} pozisyonu: ({x}, {y})")
        messagebox.showinfo("Başarılı", f"Pozisyon: X={x}, Y={y}")

    def save_settings(self):
        try:
            min_val = int(self.min_interval_var.get())
            max_val = int(self.max_interval_var.get())

            if min_val >= max_val:
                messagebox.showwarning("Uyarı", "Minimum değer maksimumdan küçük olmalı!")
                return

            self.config["min_interval_seconds"] = min_val
            self.config["max_interval_seconds"] = max_val
            self.config["idle_wait_seconds"] = int(self.idle_wait_var.get())
            self.config["click_buttons"][0]["enabled"] = self.btn1_enabled_var.get()
            self.config["click_buttons"][1]["enabled"] = self.btn2_enabled_var.get()
            self.config["auto_relogin"] = self.auto_relogin_var.get()
            self.config["start_minimized"] = self.start_minimized_var.get()

            self.save_config()
            self.update_summary()
            self.log("Ayarlar kaydedildi")
            messagebox.showinfo("Başarılı", "Ayarlar kaydedildi!")
        except ValueError:
            messagebox.showerror("Hata", "Geçerli değerler girin!")

    def save_login_settings(self):
        try:
            self.config["login_settings"] = {
                "desktop_exe_x": int(self.exe_x_var.get()),
                "desktop_exe_y": int(self.exe_y_var.get()),
                "username_combobox_x": int(self.user_x_var.get()),
                "username_combobox_y": int(self.user_y_var.get()),
                "user_selection_method": self.sel_method_var.get(),
                "user_index": int(self.user_index_var.get()),
                "password_x": int(self.pass_x_var.get()),
                "password_y": int(self.pass_y_var.get()),
                "login_button_x": int(self.login_btn_x_var.get()),
                "login_button_y": int(self.login_btn_y_var.get()),
                "main_page_button_x": int(self.main_btn_x_var.get()),
                "main_page_button_y": int(self.main_btn_y_var.get()),
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

    def update_summary(self):
        min_sec = self.config["min_interval_seconds"]
        max_sec = self.config["max_interval_seconds"]
        idle = self.config["idle_wait_seconds"]

        enabled_btns = sum(1 for b in self.config["click_buttons"] if b["enabled"])
        auto = "Açık" if self.config.get("auto_relogin", True) else "Kapalı"

        method = self.config["login_settings"].get("user_selection_method", "index")
        user_info = f"{self.config['login_settings'].get('user_index', 1)}. sıra" if method == "index" else "İsimle"

        summary = f"Tıklama Aralığı: {min_sec}-{max_sec} saniye (random)\n"
        summary += f"Aktivite Bekleme: {idle} saniye | Kullanıcı: {user_info}\n"
        summary += f"Aktif Buton: {enabled_btns} adet | Otomatik Giriş: {auto}"

        self.summary_label.config(text=summary)

    # === Pencere ve Sistem Fonksiyonları ===

    def find_window_by_title(self, title_contains):
        EnumWindows = ctypes.windll.user32.EnumWindows
        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.POINTER(ctypes.c_int), ctypes.POINTER(ctypes.c_int))
        GetWindowText = ctypes.windll.user32.GetWindowTextW
        GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
        IsWindowVisible = ctypes.windll.user32.IsWindowVisible

        found = []

        def callback(hwnd, lParam):
            if IsWindowVisible(hwnd):
                length = GetWindowTextLength(hwnd)
                buff = ctypes.create_unicode_buffer(length + 1)
                GetWindowText(hwnd, buff, length + 1)
                if title_contains.lower() in buff.value.lower():
                    found.append((hwnd, buff.value))
            return True

        EnumWindows(EnumWindowsProc(callback), 0)
        return found

    def get_foreground_window_title(self):
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        buff = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buff, length + 1)
        return buff.value

    def is_medula_running(self):
        windows = self.find_window_by_title(self.config["login_settings"]["window_title_contains"])
        return len(windows) > 0, windows

    def is_medula_foreground(self):
        title = self.get_foreground_window_title()
        return self.config["login_settings"]["window_title_contains"].lower() in title.lower()

    def check_general_announcements(self):
        try:
            is_running, _ = self.is_medula_running()
            return is_running
        except Exception as e:
            self.log(f"Kontrol hatası: {e}")
            return False

    def kill_medula(self):
        try:
            self.log("Medula kapatılıyor...")
            subprocess.run(["taskkill", "/F", "/IM", "BotanikEOS.exe"],
                         capture_output=True, shell=True)
            time.sleep(2)
            self.log("Medula kapatıldı")
            self.medula_status_var.set("Medula: KAPATILDI")
        except Exception as e:
            self.log(f"Taskkill hatası: {e}")

    def perform_login(self):
        """Otomatik giriş yap"""
        if not self.medula_password:
            self.log("HATA: Şifre girilmemiş!")
            return False

        ls = self.config["login_settings"]

        try:
            self.log("=== OTOMATİK GİRİŞ BAŞLIYOR ===")

            # 1. Medula'yı kapat
            subprocess.run(["taskkill", "/F", "/IM", "BotanikEOS.exe"],
                         capture_output=True, shell=True)
            time.sleep(2)

            # 2. Masaüstü exe'ye çift tıkla
            self.log(f"Masaüstü kısayoluna tıklanıyor ({ls['desktop_exe_x']}, {ls['desktop_exe_y']})...")
            pyautogui.click(ls["desktop_exe_x"], ls["desktop_exe_y"])
            time.sleep(0.5)
            pyautogui.doubleClick(ls["desktop_exe_x"], ls["desktop_exe_y"])

            # 3. Bekle
            self.log(f"{ls['wait_after_exe_click']} saniye bekleniyor...")
            time.sleep(ls["wait_after_exe_click"])

            # 4. Login penceresi bekle
            for _ in range(10):
                windows = self.find_window_by_title(ls["login_window_title"])
                if windows:
                    self.log("Giriş penceresi açıldı")
                    break
                time.sleep(1)

            # 5. Kullanıcı seçimi
            self.log("Kullanıcı seçiliyor...")
            pyautogui.click(ls["username_combobox_x"], ls["username_combobox_y"])
            time.sleep(0.3)

            if self.user_selection_method == "index":
                # Sıra numarasına göre seç
                user_idx = self.user_index if self.user_index else ls.get("user_index", 1)
                self.log(f"Combobox'tan {user_idx}. kullanıcı seçiliyor...")

                # Combobox'u aç
                pyautogui.click(ls["username_combobox_x"], ls["username_combobox_y"])
                time.sleep(0.3)

                # Önce en başa git (Home tuşu)
                pyautogui.press('home')
                time.sleep(0.1)

                # İstenen sıraya kadar aşağı git
                for i in range(user_idx - 1):
                    pyautogui.press('down')
                    time.sleep(0.1)

                # Seç
                pyautogui.press('enter')
                time.sleep(0.3)
            else:
                # İsme göre ara
                self.log(f"Kullanıcı adı yazılıyor: {self.medula_username}...")
                pyautogui.tripleClick(ls["username_combobox_x"], ls["username_combobox_y"])
                time.sleep(0.2)
                pyautogui.typewrite(self.medula_username, interval=0.05)
                time.sleep(0.3)

            # 6. Şifre
            self.log("Şifre giriliyor...")
            pyautogui.click(ls["password_x"], ls["password_y"])
            time.sleep(0.3)
            pyautogui.tripleClick(ls["password_x"], ls["password_y"])
            time.sleep(0.2)
            pyautogui.typewrite(self.medula_password, interval=0.05)
            time.sleep(0.3)

            # 7. Giriş butonu
            self.log("Giriş butonuna tıklanıyor...")
            pyautogui.click(ls["login_button_x"], ls["login_button_y"])

            # 8. Bekle
            self.log(f"Giriş sonrası {ls['wait_after_login']} saniye bekleniyor...")
            time.sleep(ls["wait_after_login"])

            # 9. Kontrol
            is_running, _ = self.is_medula_running()
            if is_running:
                self.log("=== OTOMATİK GİRİŞ BAŞARILI ===")
                self.relogin_count += 1
                self.root.after(0, lambda: self.relogin_count_var.set(f"Otomatik Giriş: {self.relogin_count}"))
                return True
            else:
                self.log("Giriş başarısız olabilir")
                return False

        except Exception as e:
            self.log(f"Giriş hatası: {e}")
            return False

    def manual_system_check(self):
        self.log("Manuel sistem kontrolü...")

        is_running, windows = self.is_medula_running()
        if is_running:
            self.medula_status_var.set(f"Medula: AÇIK ({len(windows)} pencere)")
            self.log(f"Medula açık: {[w[1] for w in windows]}")
        else:
            self.medula_status_var.set("Medula: KAPALI")
            self.log("Medula kapalı!")

    def manual_login(self):
        if not self.medula_password:
            messagebox.showwarning("Uyarı", "Önce giriş bilgilerini girin!")
            return

        if messagebox.askyesno("Onay", "Otomatik giriş yapılacak. Devam?"):
            threading.Thread(target=self.perform_login, daemon=True).start()

    # === Aktivite Takibi ===

    def on_mouse_activity(self, x, y, button=None, pressed=None):
        if self.is_medula_foreground():
            with self.activity_lock:
                self.last_activity_time = time.time()

    def on_keyboard_activity(self, key):
        if self.is_medula_foreground():
            with self.activity_lock:
                self.last_activity_time = time.time()

    def start_activity_listeners(self):
        self.mouse_listener = mouse.Listener(
            on_move=self.on_mouse_activity,
            on_click=self.on_mouse_activity,
            on_scroll=self.on_mouse_activity
        )
        self.keyboard_listener = keyboard.Listener(
            on_press=self.on_keyboard_activity
        )
        self.mouse_listener.start()
        self.keyboard_listener.start()
        self.log("Aktivite dinleyicileri başlatıldı")

    def stop_activity_listeners(self):
        if self.mouse_listener:
            self.mouse_listener.stop()
        if self.keyboard_listener:
            self.keyboard_listener.stop()
        self.log("Aktivite dinleyicileri durduruldu")

    def get_idle_time(self):
        with self.activity_lock:
            return time.time() - self.last_activity_time

    # === Ana Koruma Döngüsü ===

    def start_protection(self):
        if self.is_running:
            return

        if not self.medula_password:
            messagebox.showwarning("Uyarı", "Önce giriş bilgilerini girin!")
            return

        enabled = [b for b in self.config["click_buttons"] if b["enabled"]]
        if not enabled:
            messagebox.showwarning("Uyarı", "En az bir buton aktif olmalı!")
            return

        self.is_running = True
        self.stop_event.clear()

        self.status_label.config(text="ÇALIŞIYOR", foreground="green")
        self.info_label.config(text="Koruma aktif")
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        self.start_activity_listeners()
        self.log("=== KORUMA BAŞLATILDI ===")

        self.main_thread = threading.Thread(target=self.protection_loop, daemon=True)
        self.main_thread.start()

    def stop_protection(self):
        if not self.is_running:
            return

        self.is_running = False
        self.stop_event.set()

        self.stop_activity_listeners()

        self.status_label.config(text="DURDURULDU", foreground="red")
        self.info_label.config(text="Koruma durduruldu")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.countdown_var.set("Sonraki tıklama: -")

        self.log("=== KORUMA DURDURULDU ===")

    def protection_loop(self):
        ls = self.config["login_settings"]

        self.log("İlk sistem kontrolü yapılıyor...")

        is_running, _ = self.is_medula_running()

        if not is_running:
            self.log("Medula kapalı - otomatik giriş yapılıyor...")
            self.root.after(0, lambda: self.medula_status_var.set("Medula: KAPALI - Giriş yapılıyor"))
            self.perform_login()
            time.sleep(3)

        # Ana sayfaya git
        self.log("Ana sayfaya gidiliyor...")
        pyautogui.click(ls["main_page_button_x"], ls["main_page_button_y"])
        time.sleep(2)

        if not self.check_general_announcements():
            self.log("Sistem aktif değil - yeniden giriş yapılıyor...")
            self.perform_login()
            time.sleep(3)
            pyautogui.click(ls["main_page_button_x"], ls["main_page_button_y"])
            time.sleep(2)

        self.root.after(0, lambda: self.medula_status_var.set("Medula: AKTİF"))
        self.log("Sistem aktif - tıklama döngüsüne başlanıyor...")

        while not self.stop_event.is_set():
            try:
                min_sec = self.config["min_interval_seconds"]
                max_sec = self.config["max_interval_seconds"]
                wait_seconds = random.randint(min_sec, max_sec)

                self.log(f"Sonraki tıklama: {wait_seconds} saniye")

                for remaining in range(wait_seconds, 0, -1):
                    if self.stop_event.is_set():
                        return

                    idle_time = self.get_idle_time()
                    idle_wait = self.config["idle_wait_seconds"]

                    if self.is_medula_foreground() and idle_time < idle_wait:
                        self.root.after(0, lambda: self.activity_status_var.set(
                            f"Aktivite: Kullanıcı aktif (bekliyor)"))
                        self.root.after(0, lambda: self.countdown_var.set(
                            f"Kullanıcı aktif - bekleniyor..."))

                        while self.get_idle_time() < idle_wait and not self.stop_event.is_set():
                            time.sleep(1)

                        if self.stop_event.is_set():
                            return

                        wait_seconds = random.randint(min_sec, max_sec)
                        self.log(f"Aktivite bitti - yeni bekleme: {wait_seconds} saniye")
                        remaining = wait_seconds
                        continue

                    self.root.after(0, lambda r=remaining: self.countdown_var.set(
                        f"Sonraki tıklama: {r} saniye"))
                    self.root.after(0, lambda: self.activity_status_var.set(
                        f"Aktivite: Boşta ({int(idle_time)}s)"))

                    time.sleep(1)

                is_running, _ = self.is_medula_running()

                if not is_running:
                    self.log("UYARI: Medula kapalı tespit edildi!")
                    self.root.after(0, lambda: self.medula_status_var.set("Medula: KAPALI"))

                    if self.config.get("auto_relogin", True):
                        self.log("Otomatik giriş yapılıyor...")
                        self.perform_login()
                        time.sleep(3)
                        pyautogui.click(ls["main_page_button_x"], ls["main_page_button_y"])
                        time.sleep(2)
                    continue

                enabled_buttons = [b for b in self.config["click_buttons"] if b["enabled"]]
                if enabled_buttons:
                    button = random.choice(enabled_buttons)

                    self.log(f"Tıklanıyor: {button['name']} ({button['x']}, {button['y']})")
                    pyautogui.click(button["x"], button["y"])

                    self.click_count += 1
                    self.root.after(0, lambda: self.click_count_var.set(f"Toplam Tıklama: {self.click_count}"))

                    now = datetime.now().strftime("%H:%M:%S")
                    self.root.after(0, lambda b=button, t=now: self.last_click_var.set(
                        f"Son Tıklama: {t} - {b['name']}"))

                self.root.after(0, lambda: self.medula_status_var.set("Medula: AKTİF"))

            except Exception as e:
                self.log(f"Döngü hatası: {e}")
                time.sleep(5)

    # === Tray ve Pencere ===

    def minimize_to_tray(self):
        self.root.withdraw()

        if self.tray_icon is None:
            menu = pystray.Menu(
                pystray.MenuItem("Göster", self.show_from_tray),
                pystray.MenuItem("Başlat", self.tray_start),
                pystray.MenuItem("Durdur", self.tray_stop),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("Çıkış", self.quit_app)
            )
            self.tray_icon = pystray.Icon("ACMEDULA", self.icon_image, "ACMEDULA", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_from_tray(self, icon=None, item=None):
        self.root.after(0, self.root.deiconify)

    def tray_start(self, icon=None, item=None):
        self.root.after(0, self.start_protection)

    def tray_stop(self, icon=None, item=None):
        self.root.after(0, self.stop_protection)

    def quit_app(self, icon=None, item=None):
        self.stop_protection()
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.after(0, self.root.destroy)

    def on_close(self):
        if messagebox.askyesno("Çıkış", "Kapatmak mı, küçültmek mi?\n\nEvet = Kapat\nHayır = Küçült"):
            self.quit_app()
        else:
            self.minimize_to_tray()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = ACMedulaApp()
    app.run()
