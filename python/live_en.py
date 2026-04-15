/*
 * Copyright 2026 Sakuramori Lab (Sakku)
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 *
 * == Special Thanks ==
 * Development made easy thanks to a certain cat's magic. 🐈✨
 */
import os
import re
import threading
import json
import tkinter as tk
from tkinter import messagebox, filedialog, scrolledtext, ttk
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import random
import sys
import urllib.request
import queue
import subprocess

# --- 🔐 Encryption ---
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False

# --- 📺 pytchat ---
try:
    import pytchat
    PYTCHAT_AVAILABLE = True
except ImportError:
    PYTCHAT_AVAILABLE = False

# --- 📡 OBS WebSocket v5 ---
try:
    import websocket
    OBS_WS_AVAILABLE = True
except ImportError:
    OBS_WS_AVAILABLE = False

# --- 🐈 YouTube API ---
try:
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from googleapiclient.discovery import build
    import pickle
    YOUTUBE_API_AVAILABLE = True
except ImportError:
    YOUTUBE_API_AVAILABLE = False

APP_VERSION = "v5.2.0-PYTCHAT-FIRST-EN"

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

JSON_KEY_PATH = get_resource_path("credentials.json")
SETTINGS_PATH = get_resource_path("settings.json")
OBS_TEXT_FILE = get_resource_path("sakuneko_display.txt")
TOKEN_PATH = get_resource_path("token.pickle")
OBS_EVENTS_FILE = get_resource_path("obs_events.json")


# ==========================================
# 🐾 Trial Management (5 launches / 14 days limit)
# ==========================================
TRIAL_FILE      = get_resource_path("live_trial.json")
TRIAL_MAX_DAYS  = 14
TRIAL_MAX_BOOTS = 5

def check_trial():
    """Check trial status, return False if limit exceeded"""
    today = datetime.now().date()
    data  = {}

    if os.path.exists(TRIAL_FILE):
        try:
            with open(TRIAL_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = {}

    if "first_launch" not in data:
        data["first_launch"] = today.isoformat()
        data["boot_count"]   = 0

    first_day   = datetime.fromisoformat(data["first_launch"]).date()
    days_passed = (today - first_day).days
    boot_count  = data.get("boot_count", 0)

    if days_passed >= TRIAL_MAX_DAYS:
        _show_trial_expired(f"Trial period ({TRIAL_MAX_DAYS} days) has expired.\nPlease purchase the full version.")
        return False
    if boot_count >= TRIAL_MAX_BOOTS:
        _show_trial_expired(f"Trial launch limit ({TRIAL_MAX_BOOTS} launches) reached.\nPlease purchase the full version.")
        return False

    data["boot_count"] = boot_count + 1
    try:
        with open(TRIAL_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return True

def _show_trial_expired(msg):
    """Show trial expiration dialog"""
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showwarning(
        "⏰ Trial Limit Reached",
        f"【PoiBox Live - Trial】\n\n{msg}\n\nThank you for trying! 🐾"
    )
    _root.destroy()

# ==========================================
# 🔐 Encrypted Config Manager
# ==========================================
class ConfigManager:
    def __init__(self, config_dir: str = "poibox_config"):
        self.config_dir = get_resource_path(config_dir)
        self.key_file = os.path.join(self.config_dir, ".key")
        self._cipher = None
        os.makedirs(self.config_dir, exist_ok=True)
        self._init_crypto()
    
    def _init_crypto(self):
        if not CRYPTO_AVAILABLE:
            return
        if os.path.exists(self.key_file):
            with open(self.key_file, 'rb') as f:
                key = f.read()
        else:
            key = Fernet.generate_key()
            with open(self.key_file, 'wb') as f:
                f.write(key)
        self._cipher = Fernet(key)
    
    def encrypt(self, data: str) -> str:
        if self._cipher:
            return self._cipher.encrypt(data.encode()).decode()
        return data
    
    def decrypt(self, data: str) -> str:
        if self._cipher:
            try:
                return self._cipher.decrypt(data.encode()).decode()
            except:
                return data
        return data
    
    def save_api_key(self, key: str):
        enc = self.encrypt(key)
        with open(os.path.join(self.config_dir, "youtube_api_key.enc"), 'w') as f:
            f.write(enc)
    
    def load_api_key(self) -> str:
        path = os.path.join(self.config_dir, "youtube_api_key.enc")
        if os.path.exists(path):
            with open(path, 'r') as f:
                return self.decrypt(f.read().strip())
        return ""


# ==========================================
# 📡 OBS WebSocket v5
# ==========================================
class OBSWebSocketV5:
    def __init__(self, logger=None):
        self.log = logger or print
        self.ws = None
        self.is_connected = False
        self.req_id = 0
    
    def _next_id(self):
        self.req_id += 1
        return self.req_id
    
    def connect(self, url: str = "ws://localhost:4455", password: str = ""):
        if not OBS_WS_AVAILABLE:
            return False
        try:
            self.ws = websocket.WebSocketApp(
                url,
                on_open=lambda ws: self._on_open(ws, password),
                on_message=self._on_message,
                on_error=self._on_error,
                on_close=self._on_close
            )
            threading.Thread(target=self.ws.run_forever, daemon=True).start()
            for _ in range(20):
                if self.is_connected:
                    return True
                time.sleep(0.5)
            return False
        except:
            return False
    
    def _on_open(self, ws, password):
        if password:
            ws.send(json.dumps({"op": 1, "d": {"rpcVersion": 1, "authentication": password}}))
        else:
            ws.send(json.dumps({"op": 0, "d": {"rpcVersion": 1}}))
    
    def _on_message(self, ws, msg):
        try:
            data = json.loads(msg)
            if data.get("op") == 2:
                self.is_connected = True
        except:
            pass
    
    def _on_error(self, ws, error):
        self.log(f"⚠️ OBS: {error}")
    
    def _on_close(self, ws, *args):
        self.is_connected = False
    
    def set_text(self, source: str, text: str):
        if not self.is_connected:
            return False
        req = {
            "op": 6,
            "d": {
                "requestId": str(self._next_id()),
                "requestType": "SetInputSettings",
                "requestData": {"inputName": source, "inputSettings": {"text": text}}
            }
        }
        try:
            self.ws.send(json.dumps(req))
            return True
        except:
            return False
    
    def disconnect(self):
        if self.ws:
            self.ws.close()
        self.is_connected = False


# ==========================================
# 🐾 Main Application
# ==========================================
class PoiBoxLiveApp:
    def __init__(self, root):
        self.root = root
        self.root.title(f"🐾 PoiBox Live {APP_VERSION}")
        self.root.geometry("750x950")
        self.root.configure(bg="#0f0f0f")
        
        # Components
        self.config = ConfigManager()
        self.obs = OBSWebSocketV5(self.log)
        
        # State
        self.is_monitoring = False
        self.target_sheet = None
        self.current_admin = ""
        self.current_stream_title = ""
        self.concurrent_viewers = 0
        self.user_cache = {}
        self.next_row = 2
        self.total_comments = 0
        self.user_comment_counts = {}
        self.user_greeting_counts = {}
        self.comment_timestamps = {}
        self.api_queue = queue.Queue()
        self.api_call_history = []
        self.monitoring_start_time = None
        self.next_milestone = 1000
        self.point_multiplier = 1.0
        self.youtube_service = None
        self.live_chat_id = None
        
        # Mode management: pytchat first
        self.use_api_mode = False  # False=pytchat first, True=API active
        self.api_fallback_attempted = False
        
        # Load settings
        self.load_all_settings()
        
        # Build UI
        self.setup_ui()
        
        # Periodic updates
        self.update_status_loop()
    
    def load_all_settings(self):
        """Load all settings"""
        if os.path.exists(SETTINGS_PATH):
            try:
                with open(SETTINGS_PATH, 'r') as f:
                    d = json.load(f)
                    self.sheet_url = d.get("sheet_url", "")
                    self.admin_name = d.get("admin_name", "")
                    self.secret_path = d.get("secret_path", "")
            except:
                self.sheet_url = self.admin_name = self.secret_path = ""
        else:
            self.sheet_url = self.admin_name = self.secret_path = ""
        
        obs_path = get_resource_path("obs_v5_settings.json")
        if os.path.exists(obs_path):
            try:
                with open(obs_path, 'r') as f:
                    d = json.load(f)
                    self.obs_url = d.get("obs_url", "ws://localhost:4455")
                    self.obs_password = d.get("obs_password", "")
                    self.obs_source = d.get("obs_source", "LiveComments")
                    self.obs_enabled = d.get("obs_enabled", False)
            except:
                self.obs_url = "ws://localhost:4455"
                self.obs_password = ""
                self.obs_source = "LiveComments"
                self.obs_enabled = False
        else:
            self.obs_url = "ws://localhost:4455"
            self.obs_password = ""
            self.obs_source = "LiveComments"
            self.obs_enabled = False
        
        self.api_key = self.config.load_api_key()
    
    def save_obs_settings(self):
        obs_path = get_resource_path("obs_v5_settings.json")
        with open(obs_path, 'w') as f:
            json.dump({
                "obs_url": self.obs_url,
                "obs_password": self.obs_password,
                "obs_source": self.obs_source,
                "obs_enabled": self.obs_enabled
            }, f)
    
    def setup_ui(self):
        """Build UI"""
        # Header
        header = tk.Frame(self.root, bg="#0f0f0f")
        header.pack(fill="x", pady=(10, 5))
        tk.Label(header, text="🐾 PoiBox Live", font=("Meiryo", 20, "bold"), 
                 fg="#ffcc00", bg="#0f0f0f").pack()
        tk.Label(header, text=f"{APP_VERSION} | pytchat First + API Fallback", 
                 font=("Meiryo", 9), fg="#888888", bg="#0f0f0f").pack()
        
        # Notebook (tabs)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tab_main = tk.Frame(self.notebook, bg="#1a1a1a")
        self.notebook.add(self.tab_main, text="🎮 Control")
        
        self.tab_settings = tk.Frame(self.notebook, bg="#1a1a1a")
        self.notebook.add(self.tab_settings, text="⚙️ Settings")
        
        self.tab_diagnostic = tk.Frame(self.notebook, bg="#1a1a1a")
        self.notebook.add(self.tab_diagnostic, text="🔍 Diagnostics")
        
        self.tab_log = tk.Frame(self.notebook, bg="#1a1a1a")
        self.notebook.add(self.tab_log, text="📝 Log")
        
        self.setup_main_tab()
        self.setup_settings_tab()
        self.setup_diagnostic_tab()
        self.setup_log_tab()
        
        # Status bar
        self.status_var = tk.StringVar(value="🟢 Stopped")
        status_bar = tk.Label(self.root, textvariable=self.status_var, bg="#2a2a2a", 
                              fg="#00ff00", font=("Consolas", 10), anchor="w", padx=10)
        status_bar.pack(side="bottom", fill="x")
    
    def setup_main_tab(self):
        """Main tab"""
        main = self.tab_main
        
        # Status area
        status_frame = tk.Frame(main, bg="#1a1a1a", relief="ridge", bd=1)
        status_frame.pack(fill="x", padx=15, pady=10)
        
        self.stream_title_var = tk.StringVar(value="📺 Stream Title: Not fetched")
        tk.Label(status_frame, textvariable=self.stream_title_var, font=("Meiryo", 10),
                 fg="#00ffcc", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        
        self.mode_var = tk.StringVar(value="🔄 Mode: pytchat (Priority)")
        tk.Label(status_frame, textvariable=self.mode_var, font=("Consolas", 10),
                 fg="#ffcc00", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        
        self.stats_var = tk.StringVar(value="📊 Comments: 0 | CCV: -- | Multiplier: x1.0")
        tk.Label(status_frame, textvariable=self.stats_var, font=("Consolas", 10),
                 fg="#00ff00", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        
        self.api_meter_var = tk.StringVar(value="🚥 API Load: --")
        tk.Label(status_frame, textvariable=self.api_meter_var, font=("Consolas", 9),
                 fg="#888888", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        
        # Buttons
        btn_frame = tk.Frame(main, bg="#1a1a1a")
        btn_frame.pack(pady=15)
        
        self.btn_start = tk.Button(btn_frame, text="▶ Start Monitoring", command=self.start_monitoring,
                                   bg="#28a745", fg="white", font=("Meiryo", 14, "bold"),
                                   width=15, height=1, relief="flat")
        self.btn_start.pack(side="left", padx=10)
        
        self.btn_stop = tk.Button(btn_frame, text="⏹ Stop", command=self.stop_monitoring,
                                  bg="#dc3545", fg="white", font=("Meiryo", 14, "bold"),
                                  width=10, height=1, relief="flat", state="disabled")
        self.btn_stop.pack(side="left", padx=10)
        
        # Info
        info_frame = tk.Frame(main, bg="#1a1a1a", relief="sunken", bd=1)
        info_frame.pack(fill="x", padx=15, pady=10)
        
        tk.Label(info_frame, text="💡 Mode: pytchat First (no quota usage)", 
                 font=("Meiryo", 9), fg="#00ffcc", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        tk.Label(info_frame, text="   API works as fallback when pytchat fails", 
                 font=("Meiryo", 8), fg="#888888", bg="#1a1a1a").pack(anchor="w", padx=10, pady=2)
    
    def setup_settings_tab(self):
        """Settings tab"""
        settings = self.tab_settings
        
        canvas = tk.Canvas(settings, bg="#1a1a1a", highlightthickness=0)
        scrollbar = ttk.Scrollbar(settings, orient="vertical", command=canvas.yview)
        scrollable = tk.Frame(canvas, bg="#1a1a1a")
        scrollable.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        row = 0
        
        # Spreadsheet
        self._add_section(scrollable, "📊 Spreadsheet Settings", row)
        row += 1
        
        tk.Label(scrollable, text="Sheet URL:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.sheet_url_entry = tk.Entry(scrollable, width=55, bg="#2a2a2a", fg="#00ffcc")
        self.sheet_url_entry.insert(0, getattr(self, 'sheet_url', ''))
        self.sheet_url_entry.grid(row=row, column=1, padx=10, pady=2)
        row += 1
        
        tk.Label(scrollable, text="Liver Name:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.admin_entry = tk.Entry(scrollable, width=25, bg="#2a2a2a", fg="#ffcc00")
        self.admin_entry.insert(0, getattr(self, 'admin_name', ''))
        self.admin_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        row += 1
        
        # YouTube Fallback Settings
        self._add_section(scrollable, "🎬 YouTube Settings (Fallback)", row)
        row += 1
        
        tk.Label(scrollable, text="Stream URL:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.url_entry = tk.Entry(scrollable, width=55, bg="#2a2a2a", fg="#00ffcc")
        self.url_entry.grid(row=row, column=1, padx=10, pady=2)
        row += 1
        
        tk.Label(scrollable, text="Client Secret JSON:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.secret_entry = tk.Entry(scrollable, width=45, bg="#2a2a2a", fg="#ff99cc")
        self.secret_entry.insert(0, getattr(self, 'secret_path', ''))
        self.secret_entry.grid(row=row, column=1, padx=10, pady=2)
        tk.Button(scrollable, text="📁", command=self.browse_secret, bg="#444", fg="white").grid(row=row, column=2, padx=2)
        row += 1
        
        tk.Label(scrollable, text="API Key (for CCV):", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.api_entry = tk.Entry(scrollable, width=35, bg="#2a2a2a", fg="#00ffcc", show="*")
        self.api_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        tk.Button(scrollable, text="🔐 Save Encrypted", command=self.save_api_key, bg="#444", fg="white").grid(row=row, column=2, padx=2)
        row += 1
        
        # Point Settings
        self._add_section(scrollable, "💎 Point Settings", row)
        row += 1
        
        tk.Label(scrollable, text="Cooldown (sec):", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.cooldown_entry = tk.Entry(scrollable, width=10, bg="#2a2a2a", fg="#00ffcc")
        self.cooldown_entry.insert(0, "300")
        self.cooldown_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        row += 1
        
        tk.Label(scrollable, text="Campaign Multiplier:", fg="#ffcc00", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.multiplier_combo = ttk.Combobox(scrollable, values=["1.0", "1.5", "2.0", "3.0", "5.0"], width=8, state="readonly")
        self.multiplier_combo.set("1.0")
        self.multiplier_combo.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        self.multiplier_combo.bind("<<ComboboxSelected>>", self.on_multiplier_change)
        row += 1
        
        # OBS v5 Settings
        self._add_section(scrollable, "📡 OBS WebSocket v5", row)
        row += 1
        
        tk.Label(scrollable, text="URL:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.obs_url_entry = tk.Entry(scrollable, width=30, bg="#2a2a2a", fg="#00ffcc")
        self.obs_url_entry.insert(0, getattr(self, 'obs_url', 'ws://localhost:4455'))
        self.obs_url_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        row += 1
        
        tk.Label(scrollable, text="Password:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.obs_pass_entry = tk.Entry(scrollable, width=20, bg="#2a2a2a", fg="#ffcc00", show="*")
        self.obs_pass_entry.insert(0, getattr(self, 'obs_password', ''))
        self.obs_pass_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        row += 1
        
        tk.Label(scrollable, text="Text Source:", fg="white", bg="#1a1a1a").grid(row=row, column=0, sticky="w", padx=20, pady=2)
        self.obs_source_entry = tk.Entry(scrollable, width=25, bg="#2a2a2a", fg="#00ffcc")
        self.obs_source_entry.insert(0, getattr(self, 'obs_source', 'LiveComments'))
        self.obs_source_entry.grid(row=row, column=1, padx=10, pady=2, sticky="w")
        row += 1
        
        self.obs_enable_var = tk.BooleanVar(value=getattr(self, 'obs_enabled', False))
        tk.Checkbutton(scrollable, text="Enable OBS v5 Integration", variable=self.obs_enable_var,
                       bg="#1a1a1a", fg="white", selectcolor="#1a1a1a").grid(row=row, column=0, columnspan=2, sticky="w", padx=20, pady=5)
        row += 1
        
        tk.Button(scrollable, text="🔌 Test OBS Connection", command=self.test_obs, bg="#17a2b8", fg="white").grid(row=row, column=0, columnspan=2, pady=10)
        row += 1
        
        tk.Button(scrollable, text="💾 Save All Settings", command=self.save_all_settings,
                  bg="#28a745", fg="white", font=("Meiryo", 11, "bold"), padx=20, pady=5).grid(row=row, column=0, columnspan=2, pady=20)
    
    def setup_diagnostic_tab(self):
        """Diagnostic tab - strict checks"""
        diag = self.tab_diagnostic
        
        info_frame = tk.Frame(diag, bg="#1a1a1a", relief="ridge", bd=1)
        info_frame.pack(fill="x", padx=15, pady=10)
        tk.Label(info_frame, text="🔍 Quick Diagnostics (Strict)", font=("Meiryo", 12, "bold"),
                 fg="#ffcc00", bg="#1a1a1a").pack(anchor="w", padx=10, pady=5)
        tk.Label(info_frame, text="Checks system status strictly. Red items need immediate attention.",
                 font=("Meiryo", 9), fg="#888888", bg="#1a1a1a").pack(anchor="w", padx=10, pady=2)
        
        btn_frame = tk.Frame(diag, bg="#1a1a1a")
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="🔍 Run Strict Diagnostics", command=self.run_harsh_diagnostic,
                  bg="#ff6600", fg="white", font=("Meiryo", 12, "bold"), padx=20, pady=5).pack()
        
        result_frame = tk.LabelFrame(diag, text="📋 Diagnostic Results", bg="#1a1a1a", fg="#00ffcc")
        result_frame.pack(fill="both", expand=True, padx=15, pady=10)
        
        self.diagnostic_text = scrolledtext.ScrolledText(result_frame, bg="#0a0a0a", fg="#00ff88",
                                                          font=("Consolas", 10), height=20)
        self.diagnostic_text.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.diagnostic_text.tag_config("pass", foreground="#44ff44")
        self.diagnostic_text.tag_config("warn", foreground="#ffaa44")
        self.diagnostic_text.tag_config("fail", foreground="#ff4444")
        self.diagnostic_text.tag_config("info", foreground="#00ffcc")
    
    def setup_log_tab(self):
        """Log tab"""
        log_frame = tk.Frame(self.tab_log, bg="#0a0a0a")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, bg="#0a0a0a", fg="#00ff88",
                                                   font=("Consolas", 10), insertbackground="white")
        self.log_area.pack(fill="both", expand=True)
        
        self.log_area.tag_config("error", foreground="#ff4444")
        self.log_area.tag_config("success", foreground="#44ff44")
        self.log_area.tag_config("info", foreground="#00ffcc")
        self.log_area.tag_config("warning", foreground="#ffaa44")
        self.log_area.tag_config("event", foreground="#ff66cc")
        
        btn_frame = tk.Frame(self.tab_log, bg="#1a1a1a")
        btn_frame.pack(fill="x", pady=5)
        tk.Button(btn_frame, text="🗑 Clear", command=self.clear_log, bg="#444", fg="white").pack(side="left", padx=5)
        tk.Button(btn_frame, text="📋 Copy", command=self.copy_log, bg="#444", fg="white").pack(side="left", padx=5)
    
    def _add_section(self, parent, title, row):
        tk.Label(parent, text=title, font=("Meiryo", 11, "bold"), fg="#ffcc00", 
                 bg="#1a1a1a", anchor="w").grid(row=row, column=0, columnspan=3, sticky="w", padx=15, pady=(15, 5))
    
    # ==========================================
    # Diagnostic Methods (Strict)
    # ==========================================
    def run_harsh_diagnostic(self):
        """Run strict diagnostic"""
        self.diagnostic_text.delete("1.0", tk.END)
        self.diagnostic_text.insert(tk.END, "=" * 50 + "\n", "info")
        self.diagnostic_text.insert(tk.END, f"🔍 PoiBox Live Strict Diagnostics\n", "info")
        self.diagnostic_text.insert(tk.END, f"Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n", "info")
        self.diagnostic_text.insert(tk.END, "=" * 50 + "\n\n", "info")
        
        total_score = 100
        fail_count = 0
        warn_count = 0
        
        # 1. Required files
        self.diagnostic_text.insert(tk.END, "【1. Required Files】\n", "info")
        
        creds_ok = os.path.exists(JSON_KEY_PATH)
        if creds_ok:
            self.diagnostic_text.insert(tk.END, f"   ✅ credentials.json: Exists\n", "pass")
        else:
            self.diagnostic_text.insert(tk.END, f"   ❌ credentials.json: Missing (Sheet integration will fail)\n", "fail")
            total_score -= 20
            fail_count += 1
        
        secret_path = self.secret_entry.get().strip() if hasattr(self, 'secret_entry') else ""
        secret_ok = secret_path and os.path.exists(secret_path)
        if secret_ok:
            self.diagnostic_text.insert(tk.END, f"   ✅ client_secret.json: {os.path.basename(secret_path)}\n", "pass")
        else:
            self.diagnostic_text.insert(tk.END, f"   ⚠️ client_secret.json: Not set (API fallback unavailable)\n", "warn")
            total_score -= 10
            warn_count += 1
        
        # 2. Libraries
        self.diagnostic_text.insert(tk.END, "\n【2. Libraries】\n", "info")
        
        libs = [
            ("pytchat", PYTCHAT_AVAILABLE, 15),
            ("gspread", True, 10),
            ("cryptography", CRYPTO_AVAILABLE, 10),
            ("websocket-client", OBS_WS_AVAILABLE, 10),
            ("google-api-client", YOUTUBE_API_AVAILABLE, 10)
        ]
        
        for name, available, points in libs:
            if available:
                self.diagnostic_text.insert(tk.END, f"   ✅ {name}: Installed\n", "pass")
            else:
                self.diagnostic_text.insert(tk.END, f"   ❌ {name}: Not installed (pip install {name})\n", "fail")
                total_score -= points
                fail_count += 1
        
        # 3. Network
        self.diagnostic_text.insert(tk.END, "\n【3. Network】\n", "info")
        
        try:
            urllib.request.urlopen("https://www.google.com", timeout=5)
            self.diagnostic_text.insert(tk.END, f"   ✅ Internet: OK\n", "pass")
        except:
            self.diagnostic_text.insert(tk.END, f"   ❌ Internet: FAILED\n", "fail")
            total_score -= 15
            fail_count += 1
        
        try:
            urllib.request.urlopen("https://www.youtube.com", timeout=5)
            self.diagnostic_text.insert(tk.END, f"   ✅ YouTube: OK\n", "pass")
        except:
            self.diagnostic_text.insert(tk.END, f"   ⚠️ YouTube: Unstable\n", "warn")
            total_score -= 5
            warn_count += 1
        
        # 4. OBS
        self.diagnostic_text.insert(tk.END, "\n【4. OBS Integration】\n", "info")
        
        obs_enabled = hasattr(self, 'obs_enable_var') and self.obs_enable_var.get()
        if obs_enabled:
            url = self.obs_url_entry.get() if hasattr(self, 'obs_url_entry') else "ws://localhost:4455"
            self.diagnostic_text.insert(tk.END, f"   ℹ️ OBS v5: Enabled ({url})\n", "info")
            self.diagnostic_text.insert(tk.END, f"   → Run connection test from Settings tab\n", "info")
        else:
            self.diagnostic_text.insert(tk.END, f"   ℹ️ OBS v5: Disabled\n", "info")
        
        # 5. Config state
        self.diagnostic_text.insert(tk.END, "\n【5. Configuration】\n", "info")
        
        sheet_url = self.sheet_url_entry.get().strip() if hasattr(self, 'sheet_url_entry') else ""
        if sheet_url:
            self.diagnostic_text.insert(tk.END, f"   ✅ Spreadsheet: Configured\n", "pass")
        else:
            self.diagnostic_text.insert(tk.END, f"   ❌ Spreadsheet: Not set\n", "fail")
            total_score -= 10
            fail_count += 1
        
        admin = self.admin_entry.get().strip() if hasattr(self, 'admin_entry') else ""
        if admin:
            self.diagnostic_text.insert(tk.END, f"   ✅ Liver Name: {admin}\n", "pass")
        else:
            self.diagnostic_text.insert(tk.END, f"   ❌ Liver Name: Not set\n", "fail")
            total_score -= 5
            fail_count += 1
        
        # Final score
        self.diagnostic_text.insert(tk.END, "\n" + "=" * 50 + "\n", "info")
        self.diagnostic_text.insert(tk.END, f"【Diagnostic Result】\n", "info")
        self.diagnostic_text.insert(tk.END, f"   Total Score: {total_score}/100\n", "info")
        self.diagnostic_text.insert(tk.END, f"   Errors: {fail_count}\n", "info")
        self.diagnostic_text.insert(tk.END, f"   Warnings: {warn_count}\n", "info")
        
        if total_score >= 90:
            self.diagnostic_text.insert(tk.END, f"\n   🎉 Rating: Excellent - No issues\n", "pass")
        elif total_score >= 70:
            self.diagnostic_text.insert(tk.END, f"\n   ⚠️ Rating: Caution - Some features may be limited\n", "warn")
        else:
            self.diagnostic_text.insert(tk.END, f"\n   ❌ Rating: Action Required - Check results\n", "fail")
        
        self.diagnostic_text.insert(tk.END, "=" * 50 + "\n", "info")
    
    # ==========================================
    # Utilities
    # ==========================================
    def clear_log(self):
        self.log_area.delete("1.0", tk.END)
    
    def copy_log(self):
        text = self.log_area.get("1.0", tk.END)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.log("📋 Log copied to clipboard", "info")
    
    def log(self, msg, tag="success"):
        now = datetime.now().strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{now}] {msg}\n", tag)
        self.log_area.see(tk.END)
    
    def save_all_settings(self):
        with open(SETTINGS_PATH, 'w') as f:
            json.dump({
                "sheet_url": self.sheet_url_entry.get(),
                "admin_name": self.admin_entry.get(),
                "secret_path": self.secret_entry.get()
            }, f)
        
        self.obs_url = self.obs_url_entry.get()
        self.obs_password = self.obs_pass_entry.get()
        self.obs_source = self.obs_source_entry.get()
        self.obs_enabled = self.obs_enable_var.get()
        self.save_obs_settings()
        
        self.log("💾 All settings saved", "info")
        messagebox.showinfo("Saved", "Settings saved!")
    
    def save_api_key(self):
        key = self.api_entry.get().strip()
        if key:
            self.config.save_api_key(key)
            self.api_key = key
            self.log("🔐 API key saved (encrypted)", "success")
            messagebox.showinfo("Success", "API key saved encrypted!")
    
    def on_multiplier_change(self, event=None):
        try:
            self.point_multiplier = float(self.multiplier_combo.get())
            self.log(f"🔥 Campaign multiplier: x{self.point_multiplier}", "event")
        except:
            pass
    
    def test_obs(self):
        url = self.obs_url_entry.get()
        pwd = self.obs_pass_entry.get()
        self.log(f"🔌 Testing OBS v5 connection...", "info")
        
        def test():
            client = OBSWebSocketV5(self.log)
            if client.connect(url, pwd):
                self.log("✅ OBS v5 connection successful!", "success")
                client.disconnect()
            else:
                self.log("❌ OBS connection failed", "error")
        
        threading.Thread(target=test, daemon=True).start()
    
    def browse_secret(self):
        path = filedialog.askopenfilename(title="Select client_secret.json")
        if path:
            self.secret_entry.delete(0, tk.END)
            self.secret_entry.insert(0, path)
    
    def update_status_loop(self):
        if self.is_monitoring:
            mode_text = "🔄 pytchat (Priority)" if not self.use_api_mode else "🔄 API (Fallback active)"
            self.mode_var.set(f"🔄 Mode: {mode_text}")
            
            viewers = self.concurrent_viewers if self.concurrent_viewers > 0 else "--"
            self.stats_var.set(f"📊 Comments: {self.total_comments} | CCV: {viewers} | Multiplier: x{self.point_multiplier}")
            
            now = time.time()
            self.api_call_history = [t for t in self.api_call_history if now - t <= 60]
            cnt = len(self.api_call_history)
            if cnt < 30:
                meter = f"🟢 Stable ({cnt}/60)"
            elif cnt < 50:
                meter = f"🟡 Busy ({cnt}/60)"
            else:
                meter = f"🔴 Near limit ({cnt}/60)"
            self.api_meter_var.set(f"🚥 API Load: {meter}")
        
        self.root.after(1000, self.update_status_loop)
    
    # ==========================================
    # Monitoring Methods (pytchat First)
    # ==========================================
    def start_monitoring(self):
        sheet_url = self.sheet_url_entry.get().strip()
        admin_name = self.admin_entry.get().strip()
        yt_url = self.url_entry.get().strip()
        
        v_match = re.search(r"(v=|/)([0-9A-Za-z_-]{11})", yt_url)
        if not v_match or not sheet_url or not admin_name:
            messagebox.showwarning("Error", "Please fill in all required fields!")
            return
        
        # Connect to spreadsheet
        try:
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_PATH, scope)
            self.target_sheet = gspread.authorize(creds).open_by_url(sheet_url).get_worksheet(0)
            records = self.target_sheet.get_all_records()
            self.user_cache = {str(r['Handle']): {'Points': r['Points'], 'Total': r['Total'], 'LastTouch': r['LastTouch'], 'Row': i+2} 
                              for i, r in enumerate(records) if r.get('Handle')}
            self.next_row = len(records) + 2
        except Exception as e:
            self.log(f"❌ Spreadsheet connection failed: {e}", "error")
            messagebox.showerror("Error", "Cannot connect to spreadsheet")
            return
        
        video_id = v_match.group(2)
        self.video_id = video_id
        
        self.is_monitoring = True
        self.current_admin = admin_name
        self.total_comments = 0
        self.user_comment_counts = {}
        self.user_greeting_counts = {}
        self.comment_timestamps = {}
        self.monitoring_start_time = datetime.now()
        self.concurrent_viewers = 0
        self.use_api_mode = False
        self.api_fallback_attempted = False
        
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.status_var.set(f"🟢 Monitoring - {admin_name}")
        self.log(f"🚀 Started monitoring for {admin_name}!", "success")
        self.log(f"📡 Mode: pytchat First (no quota usage)", "info")
        
        # OBS v5 connection
        if self.obs_enable_var.get():
            self.obs_url = self.obs_url_entry.get()
            self.obs_password = self.obs_pass_entry.get()
            self.save_obs_settings()
            threading.Thread(target=lambda: self.obs.connect(self.obs_url, self.obs_password), daemon=True).start()
        
        # Start threads
        threading.Thread(target=self._update_viewers_loop, args=(video_id,), daemon=True).start()
        threading.Thread(target=self.api_worker, daemon=True).start()
        
        # pytchat first loop
        threading.Thread(target=self._pytchat_priority_loop, args=(video_id,), daemon=True).start()
    
    def _pytchat_priority_loop(self, video_id):
        """pytchat priority loop - fallback to API if it dies"""
        self.log("🚀 Starting pytchat mode (Priority)", "success")
        
        while self.is_monitoring and not self.use_api_mode:
            try:
                chat = pytchat.create(video_id=video_id, interruptable=False)
                self.log("✅ pytchat connected", "success")
                
                while chat.is_alive() and self.is_monitoring and not self.use_api_mode:
                    for c in chat.get().items:
                        if not self.is_monitoring:
                            break
                        self._record_comment(c.author.name, c.message, c.amountValue)
                    time.sleep(1)
                
                if self.is_monitoring and not self.use_api_mode:
                    self.log("⚠️ pytchat disconnected, retrying...", "warning")
                    time.sleep(3)
                    
            except Exception as e:
                self.log(f"❌ pytchat error: {e}", "error")
                if self.is_monitoring and not self.use_api_mode:
                    self.log("⏳ Retrying in 3 seconds...", "warning")
                    time.sleep(3)
        
        # Fallback to API if pytchat died
        if self.is_monitoring and not self.use_api_mode:
            self._fallback_to_api(video_id)
    
    def _fallback_to_api(self, video_id):
        """API fallback"""
        self.log("=" * 40, "info")
        self.log("🔄 pytchat stopped → Executing API fallback", "event")
        
        secret_path = self.secret_entry.get().strip()
        if not secret_path or not os.path.exists(secret_path):
            self.log("❌ No API credentials file → Fallback failed", "error")
            self.log("⚠️ Stopping monitoring", "error")
            self.stop_monitoring()
            return
        
        try:
            scopes = ["https://www.googleapis.com/auth/youtube.force-ssl"]
            creds = None
            if os.path.exists(TOKEN_PATH):
                with open(TOKEN_PATH, "rb") as t:
                    creds = pickle.load(t)
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    creds.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(secret_path, scopes)
                    creds = flow.run_local_server(port=0)
                with open(TOKEN_PATH, "wb") as t:
                    pickle.dump(creds, t)
            
            self.youtube_service = build("youtube", "v3", credentials=creds)
            res = self.youtube_service.videos().list(part="liveStreamingDetails,snippet", id=video_id).execute()
            if res.get("items"):
                self.live_chat_id = res["items"][0]["liveStreamingDetails"].get("activeLiveChatId")
                self.current_stream_title = res["items"][0]["snippet"]["title"]
                self.stream_title_var.set(f"📺 {self.current_stream_title[:60]}")
                
                if self.live_chat_id:
                    self.use_api_mode = True
                    self.log("✅ API mode switch successful!", "success")
                    self.log(f"📺 Title: {self.current_stream_title[:50]}", "info")
                    self._api_chat_loop()
                    return
            
            self.log("❌ API mode switch failed", "error")
            self.stop_monitoring()
            
        except Exception as e:
            self.log(f"❌ API authentication error: {e}", "error")
            self.stop_monitoring()
    
    def _api_chat_loop(self):
        """API chat loop"""
        self.log("🚀 Starting API mode (Fallback)", "success")
        next_token = None
        
        while self.is_monitoring and self.use_api_mode:
            try:
                res = self.youtube_service.liveChatMessages().list(
                    liveChatId=self.live_chat_id,
                    part="snippet,authorDetails",
                    pageToken=next_token
                ).execute()
                
                for item in res.get("items", []):
                    handle = item["authorDetails"]["displayName"]
                    msg = item["snippet"]["displayMessage"]
                    sc = 0
                    if item["snippet"]["type"] == "superChatEvent":
                        sc = item["snippet"]["superChatDetails"]["amountMicros"] / 1000000
                    self._record_comment(handle, msg, sc)
                
                next_token = res.get("nextPageToken")
                time.sleep(max(res.get("pollingIntervalMillis", 3000) / 1000, 3))
                
            except Exception as e:
                if "quotaExceeded" in str(e):
                    self.log("❌ API quota exceeded!", "error")
                    self.log("⚠️ Stopping monitoring", "error")
                    self.stop_monitoring()
                    break
                self.log(f"⚠️ API error: {e}", "warning")
                time.sleep(5)
    
    def _update_viewers_loop(self, video_id):
        while self.is_monitoring:
            api_key = self.config.load_api_key()
            if not api_key:
                api_key = self.api_entry.get().strip()
            if api_key:
                try:
                    url = f"https://www.googleapis.com/youtube/v3/videos?part=liveStreamingDetails&id={video_id}&key={api_key}"
                    with urllib.request.urlopen(url, timeout=5) as resp:
                        data = json.loads(resp.read())
                        if data.get("items"):
                            viewers = data["items"][0].get("liveStreamingDetails", {}).get("concurrentViewers")
                            if viewers:
                                self.concurrent_viewers = int(viewers)
                except:
                    pass
            time.sleep(30)
    
    def _record_comment(self, handle, message, sc_amount):
        self.total_comments += 1
        self.user_comment_counts[handle] = self.user_comment_counts.get(handle, 0) + 1
        
        mode_mark = "🔵" if self.use_api_mode else "🟢"
        self.log(f"{mode_mark} 💬 [{self.total_comments}] {handle}: {message[:50]}", "info")
        
        # Simple point processing
        is_overall = self.total_comments % 100 == 0
        self._process_points_simple(handle, message, is_overall, self.user_comment_counts[handle], sc_amount)
    
    def _process_points_simple(self, handle, message, is_overall, personal, sc_amount):
        """Simple point processing"""
        try:
            wait_sec = int(self.cooldown_entry.get().strip()) if hasattr(self, 'cooldown_entry') else 300
            
            now = datetime.now()
            now_str = now.strftime('%Y-%m-%d %H:%M:%S')
            clean_handle = str(handle).strip().replace("@", "")
            matched = next((k for k in self.user_cache if str(k).strip().replace("@", "") == clean_handle), None)
            
            base_pts = 1
            sc_pts = int((sc_amount / 10) * self.point_multiplier) if sc_amount > 0 else 0
            
            if sc_pts > 0:
                self.log(f"💰 {handle} sent {sc_amount} JPY → {sc_pts} pts!", "event")
            
            if not matched:
                init_pts = 100 + sc_pts
                self.api_queue.put({'type': 'append', 'data': [handle, init_pts, init_pts, now_str]})
                self.user_cache[handle] = {'Points': init_pts, 'Total': init_pts, 'LastTouch': now_str, 'Row': self.next_row}
                self.next_row += 1
                self.log(f"✨ 【New】 {handle} registered! (+{init_pts} pts)", "event")
            else:
                user = self.user_cache[matched]
                try:
                    last_t = datetime.strptime(str(user['LastTouch']), '%Y-%m-%d %H:%M:%S')
                except:
                    last_t = datetime(2000, 1, 1)
                
                if (now - last_t).total_seconds() >= wait_sec:
                    add_val = base_pts + sc_pts
                    if add_val > 0:
                        user['Points'] = int(user['Points']) + add_val
                        user['Total'] = int(user['Total']) + add_val
                        user['LastTouch'] = now_str
                        self.api_queue.put({'type': 'update', 'range': f'B{user["Row"]}:D{user["Row"]}', 'data': [[user['Points'], user['Total'], now_str]]})
                        if sc_pts == 0:
                            self.log(f"💎 {matched}: +{add_val} pts", "success")
        except Exception as e:
            self.log(f"⚠️ Point processing error: {e}", "warning")
    
    def api_worker(self):
        while self.is_monitoring:
            try:
                task = self.api_queue.get(timeout=1)
                self.api_call_history.append(time.time())
                success = False
                while not success and self.is_monitoring:
                    try:
                        if task['type'] == 'append':
                            self.target_sheet.append_row(task['data'])
                        else:
                            self.target_sheet.update(task['range'], task['data'])
                        success = True
                    except:
                        time.sleep(15)
                time.sleep(1.2)
                self.api_queue.task_done()
            except queue.Empty:
                continue
    
    def stop_monitoring(self):
        self.is_monitoring = False
        self.obs.disconnect()
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.status_var.set("🟡 Stopped")
        self.log("🛑 Monitoring stopped", "warning")
        
        if self.comment_timestamps:
            peak = max(self.comment_timestamps, key=self.comment_timestamps.get)
            self.log(f"📊 Peak activity: {peak} ({self.comment_timestamps[peak]} comments/min)", "event")
            self.log(f"📝 Total comments: {self.total_comments}", "event")


if __name__ == "__main__":
    # 🐾 Trial check (5 launches / 14 days limit)
    if not check_trial():
        sys.exit(0)
    root = tk.Tk()
    app = PoiBoxLiveApp(root)
    root.mainloop()
