import os
import json
import time
import subprocess
import threading
import uuid
import hashlib
import platform
from pathlib import Path
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox
from google import genai
from google.genai import types
import yt_dlp

# ==========================================
# ⚙️ AIdirector V14.0 PRO - 30 Days Trial (HWID Bound)
# ==========================================
PROMPT_FILE = "aidirector_v14_pro_prompt.json"
API_KEY_FILE = "gemini_api_key.txt"
TRIAL_CONFIG_FILE = "aidirector_trial.dat"
CHUNK_DELAY = 10
DEFAULT_MODEL = "gemini-2.0-flash"
AVAILABLE_MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]

DEFAULT_PROMPT = """You are a world-class video editing director who creates viral content for social media.
Analyze the provided video and propose a structure for "Shorts/Highlight video" that grabs viewers instantly and keeps them watching.

【Golden Rules of Structure】
1. HOOK (0-5s): Shocking image or statement in the first 3 seconds that makes viewers go "What!?"
2. BUILD (5-20s): Story development that builds anticipation.
3. PEAK (20-45s): Climax - laughter, emotion, or excitement at its peak.
4. OUTRO (45-60s): Ending that leaves a lasting impression and encourages channel subscription.

【Output Format】
Output ONLY in JSON format. No Markdown decoration.
{
  "title": "Viral title suggestion",
  "strategy": "Explanation of why this structure will go viral",
  "scenes": {
    "hook": {"start": "00:00:10", "end": "00:00:15", "description": "..."},
    "build": {"start": "00:01:20", "end": "00:01:35", "description": "..."},
    "peak": {"start": "00:05:00", "end": "00:05:25", "description": "..."},
    "outro": {"start": "00:06:10", "end": "00:06:15", "description": "..."}
  }
}"""

class LicenseManager:
    def __init__(self):
        self.config_file = Path(TRIAL_CONFIG_FILE)
        self.trial_days = 30
        
    def get_hardware_id(self):
        """Generate a unique hardware ID for this machine"""
        system = platform.system()
        
        if system == "Windows":
            try:
                result = subprocess.run(['wmic', 'csproduct', 'get', 'uuid'], capture_output=True, text=True, timeout=5)
                uuid_line = [line for line in result.stdout.split('\n') if 'UUID' not in line and line.strip()]
                if uuid_line:
                    return hashlib.sha256(uuid_line[0].strip().encode()).hexdigest()[:32]
            except:
                pass
        elif system == "Darwin":  # macOS
            try:
                result = subprocess.run(['system_profiler', 'SPHardwareDataType'], capture_output=True, text=True, timeout=5)
                for line in result.stdout.split('\n'):
                    if 'Hardware UUID' in line:
                        uuid_val = line.split(':')[-1].strip()
                        return hashlib.sha256(uuid_val.encode()).hexdigest()[:32]
            except:
                pass
        elif system == "Linux":
            try:
                # Try machine-id first
                machine_id_paths = ['/etc/machine-id', '/var/lib/dbus/machine-id']
                for path in machine_id_paths:
                    if os.path.exists(path):
                        with open(path, 'r') as f:
                            machine_id = f.read().strip()
                            if machine_id:
                                return hashlib.sha256(machine_id.encode()).hexdigest()[:32]
            except:
                pass
        
        # Fallback: hash hostname + username + cpu count
        try:
            hostname = platform.node()
            username = os.getlogin() if hasattr(os, 'getlogin') else 'unknown'
            cpu_count = str(os.cpu_count() or 0)
            combined = f"{hostname}_{username}_{cpu_count}"
            return hashlib.sha256(combined.encode()).hexdigest()[:32]
        except:
            return hashlib.sha256(str(uuid.getnode()).encode()).hexdigest()[:32]
    
    def is_licensed(self):
        if not self.config_file.exists():
            return False
        
        try:
            with open(self.config_file, 'r') as f:
                data = json.load(f)
            
            # Check HWID binding
            stored_hwid = data.get('hardware_id')
            current_hwid = self.get_hardware_id()
            
            if stored_hwid and stored_hwid != current_hwid:
                return False
            
            if data.get('license_key'):
                return True
            
            install_date = datetime.fromisoformat(data.get('install_date', ''))
            days_used = (datetime.now() - install_date).days
            return days_used < self.trial_days
        except:
            return False
    
    def get_days_left(self):
        if not self.config_file.exists():
            return self.trial_days
        
        try:
            with open(self.config_file, 'r') as f:
                data = json.load(f)
            
            if data.get('license_key'):
                return -1
            
            install_date = datetime.fromisoformat(data.get('install_date', ''))
            days_used = (datetime.now() - install_date).days
            return max(0, self.trial_days - days_used)
        except:
            return self.trial_days
    
    def start_trial(self):
        data = {
            'install_date': datetime.now().isoformat(),
            'license_key': None,
            'hardware_id': self.get_hardware_id()
        }
        with open(self.config_file, 'w') as f:
            json.dump(data, f, indent=2)
    
    def activate_license(self, license_key):
        # In production, verify with a server
        if len(license_key) >= 16 and license_key.startswith("DIRECTOR-"):
            data = {
                'install_date': datetime.now().isoformat(),
                'license_key': license_key,
                'hardware_id': self.get_hardware_id()
            }
            with open(self.config_file, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        return False

class TrialDialog(ctk.CTkToplevel):
    def __init__(self, parent, license_mgr):
        super().__init__(parent)
        self.license_mgr = license_mgr
        self.result = None
        self.title("AIdirector V14.0 PRO - License")
        self.geometry("500x420")
        self.resizable(False, False)
        self.configure(fg_color="#1e1e1e")
        
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
    def create_widgets(self):
        ctk.CTkLabel(self, text="🎬 AIdirector V14.0 PRO", font=("Impact", 28), text_color="#00FF41").pack(pady=30)
        
        info_frame = ctk.CTkFrame(self, fg_color="#2d2d2d")
        info_frame.pack(fill="x", padx=30, pady=10)
        
        days_left = self.license_mgr.get_days_left()
        if days_left > 0:
            ctk.CTkLabel(info_frame, text=f"Trial Period: {days_left} days remaining", 
                        font=("Consolas", 14), text_color="#00ff00").pack(pady=15)
            ctk.CTkLabel(info_frame, text="Start your 30-day free trial now!\n(HWID locked to this machine)", 
                        font=("Segoe UI", 11), text_color="#aaaaaa").pack()
        else:
            ctk.CTkLabel(info_frame, text="Trial Expired", 
                        font=("Consolas", 18), text_color="#ff5555").pack(pady=15)
            ctk.CTkLabel(info_frame, text="Please purchase a license to continue using AIdirector PRO", 
                        font=("Segoe UI", 12), text_color="#aaaaaa").pack()
        
        ctk.CTkLabel(self, text="License Key (optional)", font=("Segoe UI", 12)).pack(pady=(20,5))
        self.key_entry = ctk.CTkEntry(self, placeholder_text="DIRECTOR-XXXX-XXXX-XXXX", width=350, fg_color="#121212")
        self.key_entry.pack(pady=5)
        
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        if days_left > 0:
            ctk.CTkButton(btn_frame, text="Start Trial", command=self.start_trial,
                         fg_color="#00ff00", text_color="#000000", width=150, height=40,
                         font=("Impact", 14)).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Activate License", command=self.activate_license,
                     fg_color="#00FF41", text_color="#000000", width=150, height=40,
                     font=("Impact", 14)).pack(side="left", padx=10)
        
        if days_left == 0:
            ctk.CTkButton(btn_frame, text="Purchase License", command=self.purchase,
                         fg_color="#ffaa00", text_color="#000000", width=150, height=40,
                         font=("Impact", 14)).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Exit", command=self.exit_app,
                     fg_color="#333333", text_color="#ffffff", width=100, height=40,
                     font=("Segoe UI", 12)).pack(side="left", padx=10)
    
    def start_trial(self):
        self.license_mgr.start_trial()
        self.result = "trial"
        self.destroy()
    
    def activate_license(self):
        key = self.key_entry.get().strip()
        if self.license_mgr.activate_license(key):
            messagebox.showinfo("Success", "License activated successfully!\nThank you for purchasing AIdirector PRO!")
            self.result = "licensed"
            self.destroy()
        else:
            messagebox.showerror("Error", "Invalid license key.\nFormat: DIRECTOR-XXXX-XXXX-XXXX")
    
    def purchase(self):
        import webbrowser
        webbrowser.open("https://example.com/purchase")
        messagebox.showinfo("Purchase", "Purchase page opened in your browser.")
    
    def exit_app(self):
        self.result = "exit"
        self.destroy()

class DirectorV14Trial(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # License check first
        self.license_mgr = LicenseManager()
        if not self.check_license():
            return
        
        self.title("🎬 AIdirector V14.0 PRO - 30 Day Trial")
        self.geometry("1300x950")
        ctk.set_appearance_mode("dark")
        
        self.is_running = False
        self.stop_requested = False
        self.selected_files = []
        self.output_dir = Path.cwd() / "DIRECTOR_OUTPUT"
        
        # 3-column layout
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_columnconfigure(2, weight=1)
        self.grid_rowconfigure(0, weight=4)
        self.grid_rowconfigure(1, weight=1)
        
        self.setup_obs_ui()
        self.load_saved_api_key()
        self.update_title_with_days()
        
        # Handle window close event
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
    
    def check_license(self):
        if not self.license_mgr.is_licensed():
            dialog = TrialDialog(self, self.license_mgr)
            self.wait_window(dialog)
            
            if dialog.result == "exit" or dialog.result is None:
                self.destroy()
                return False
            elif dialog.result == "trial" or dialog.result == "licensed":
                if not self.license_mgr.is_licensed():
                    self.destroy()
                    return False
        
        days_left = self.license_mgr.get_days_left()
        if days_left == 0:
            messagebox.showerror("Trial Expired", 
                                "Your 30-day trial has expired.\nPlease purchase a license to continue using AIdirector PRO.")
            self.destroy()
            return False
        
        return True
    
    def update_title_with_days(self):
        days_left = self.license_mgr.get_days_left()
        if days_left == -1:
            title_suffix = "LICENSED"
        else:
            title_suffix = f"TRIAL: {days_left} days left"
        self.title(f"🎬 AIdirector V14.0 PRO - {title_suffix}")
        self.after(3600000, self.update_title_with_days)
    
    def on_closing(self):
        """Graceful shutdown"""
        if self.is_running:
            self.stop_requested = True
            self.log("⚠️ Shutting down gracefully... Please wait.")
            self.after(2000, self.destroy)
        else:
            self.destroy()
    
    def log(self, text):
        self.after(0, lambda: (self.console_log.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {text}\n"), self.console_log.see("end")))
    
    def load_saved_api_key(self):
        if os.path.exists(API_KEY_FILE):
            with open(API_KEY_FILE, "r") as f:
                self.api_entry.insert(0, f.read().strip())
    
    def save_api_key(self):
        api_key = self.api_entry.get().strip()
        if api_key:
            with open(API_KEY_FILE, "w") as f:
                f.write(api_key)
    
    def time_to_seconds(self, time_str):
        try:
            parts = str(time_str).strip().split(':')
            if len(parts) == 3:
                return int(parts[0]) * 3600 + int(parts[1]) * 60 + float(parts[2])
            elif len(parts) == 2:
                return int(parts[0]) * 60 + float(parts[1])
            else:
                return float(parts[0])
        except:
            return 0
    
    def get_video_duration(self, video_path):
        cmd = ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "default=noprint_wrappers=1:nokey=1", str(video_path)]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            return float(result.stdout.strip())
        except:
            return None
    
    def setup_obs_ui(self):
        # ========== LEFT: SOURCES ==========
        self.left_frame = ctk.CTkFrame(self, corner_radius=0, border_width=1, border_color="#333333")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        
        ctk.CTkLabel(self.left_frame, text="SOURCES", font=("Impact", 18), text_color="#00FF41").pack(pady=10)
        
        self.file_list_box = ctk.CTkTextbox(self.left_frame, fg_color="#121212", font=("Consolas", 11))
        self.file_list_box.pack(fill="both", expand=True, padx=5, pady=5)
        
        ctk.CTkButton(self.left_frame, text="+ LOCAL FILES", command=self.browse_files, fg_color="#333333").pack(fill="x", padx=10, pady=2)
        
        self.url_entry = ctk.CTkEntry(self.left_frame, placeholder_text="YouTube URL...", fg_color="#121212")
        self.url_entry.pack(fill="x", padx=10, pady=5)
        
        # Output folder selection
        ctk.CTkButton(self.left_frame, text="📁 OUTPUT FOLDER", command=self.select_output_folder, fg_color="#333333").pack(fill="x", padx=10, pady=2)
        self.output_label = ctk.CTkLabel(self.left_frame, text=str(self.output_dir), font=("Consolas", 9), text_color="#888888")
        self.output_label.pack(pady=2)
        
        # ========== CENTER: DIRECTOR'S EYE ==========
        self.center_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#000000", border_width=1, border_color="#00FF41")
        self.center_frame.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        
        self.status_label = ctk.CTkLabel(self.center_frame, text="[ DIRECTOR READY ]", font=("Impact", 42), text_color="#00FF41")
        self.status_label.place(relx=0.5, rely=0.4, anchor="center")
        
        self.sub_status = ctk.CTkLabel(self.center_frame, text="Waiting for Target...", font=("Consolas", 14), text_color="#AAAAAA")
        self.sub_status.place(relx=0.5, rely=0.5, anchor="center")
        
        self.progress_bar = ctk.CTkProgressBar(self.center_frame, progress_color="#00FF41", height=15, width=600)
        self.progress_bar.place(relx=0.5, rely=0.9, anchor="center")
        self.progress_bar.set(0)
        
        # ========== RIGHT: CONTROLS ==========
        self.right_frame = ctk.CTkFrame(self, corner_radius=0, border_width=1, border_color="#333333")
        self.right_frame.grid(row=0, column=2, sticky="nsew", padx=2, pady=2)
        
        ctk.CTkLabel(self.right_frame, text="CONTROLS", font=("Impact", 18)).pack(pady=10)
        
        # API Key input
        self.api_entry = ctk.CTkEntry(self.right_frame, placeholder_text="API Key (Gemini)", show="*", fg_color="#121212")
        self.api_entry.pack(fill="x", padx=10, pady=5)
        
        # Model selection
        self.model_menu = ctk.CTkOptionMenu(self.right_frame, values=AVAILABLE_MODELS, fg_color="#333333")
        self.model_menu.pack(fill="x", padx=10, pady=10)
        self.model_menu.set(DEFAULT_MODEL)
        
        # Aspect ratio selection
        self.format_var = ctk.StringVar(value="9:16")
        ctk.CTkRadioButton(self.right_frame, text="Shorts (9:16)", variable=self.format_var, value="9:16").pack(pady=5, padx=20, anchor="w")
        ctk.CTkRadioButton(self.right_frame, text="Wide (16:9)", variable=self.format_var, value="16:9").pack(pady=5, padx=20, anchor="w")
        
        # Editable prompt area
        ctk.CTkLabel(self.right_frame, text="📝 EDITABLE PROMPT", font=("Consolas", 12), text_color="#aaaaaa").pack(pady=(10,0))
        self.prompt_text = ctk.CTkTextbox(self.right_frame, height=250, fg_color="#121212", font=("Consolas", 10))
        self.prompt_text.pack(fill="x", padx=10, pady=5)
        self.prompt_text.insert("1.0", self.load_prompt())
        
        # Run button
        self.btn_run = ctk.CTkButton(self.right_frame, text="DIRECT\nCUT", font=("Impact", 32), fg_color="#00FF41", text_color="#000000", height=150, command=self.launch)
        self.btn_run.pack(side="bottom", fill="x", padx=10, pady=20)
        
        # ========== BOTTOM: LOGS ==========
        self.bottom_frame = ctk.CTkFrame(self, corner_radius=0, border_width=1, border_color="#333333")
        self.bottom_frame.grid(row=1, column=0, columnspan=3, sticky="nsew", padx=2, pady=2)
        
        self.console_log = ctk.CTkTextbox(self.bottom_frame, fg_color="#000000", font=("Consolas", 12), text_color="#00FF00")
        self.console_log.pack(fill="both", expand=True, padx=5, pady=5)
    
    def select_output_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir = Path(folder)
            self.output_label.configure(text=str(self.output_dir))
    
    def browse_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Video", "*.mp4 *.mov *.mkv *.avi")])
        if files:
            self.selected_files = [str(Path(f).absolute()) for f in files]
            self.file_list_box.delete("1.0", "end")
            for f in files:
                self.file_list_box.insert("end", f"● {Path(f).name}\n")
    
    def load_prompt(self):
        if os.path.exists(PROMPT_FILE):
            with open(PROMPT_FILE, 'r', encoding='utf-8') as f:
                return json.load(f).get("prompt", DEFAULT_PROMPT)
        return DEFAULT_PROMPT
    
    def save_prompt(self):
        prompt = self.prompt_text.get("1.0", "end").strip()
        with open(PROMPT_FILE, 'w', encoding='utf-8') as f:
            json.dump({"prompt": prompt}, f, ensure_ascii=False, indent=2)
    
    def launch(self):
        if self.is_running:
            return
        
        targets = self.selected_files if self.selected_files else []
        url = self.url_entry.get().strip()
        if url:
            targets.append(url)
        
        if not targets:
            messagebox.showwarning("Warning", "Please select a video or enter a URL")
            return
        
        api_key = self.api_entry.get().strip()
        if not api_key:
            messagebox.showwarning("Warning", "Please enter your Gemini API Key")
            return
        
        self.save_api_key()
        self.save_prompt()
        
        self.is_running = True
        self.stop_requested = False
        self.status_label.configure(text="[ DIRECTING NOW ]", text_color="#00FF41")
        self.btn_run.configure(state="disabled", text="RUNNING")
        threading.Thread(target=self.batch_process, daemon=True).start()
    
    def batch_process(self):
        try:
            targets = self.selected_files if self.selected_files else []
            url = self.url_entry.get().strip()
            if url:
                targets.append(url)
            
            if not targets:
                self.log("❌ No targets selected")
                return
            
            api_key = self.api_entry.get().strip()
            model_name = self.model_menu.get()
            format_type = self.format_var.get()
            prompt = self.prompt_text.get("1.0", "end").strip()
            
            total = len(targets)
            for idx, target in enumerate(targets):
                if self.stop_requested:
                    break
                
                # Update main progress
                main_progress = idx / total
                self.after(0, lambda p=main_progress: self.progress_bar.set(p))
                
                # Handle YouTube download
                if str(target).startswith("http"):
                    self.log(f"📥 Downloading from YouTube: {target[:50]}...")
                    self.after(0, lambda: self.sub_status.configure(text="Downloading from YouTube..."))
                    tmp_dl = Path.cwd() / f"dl_{uuid.uuid4().hex[:6]}.mp4"
                    ydl_opts = {
                        'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]',
                        'outtmpl': str(tmp_dl),
                        'quiet': True
                    }
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        ydl.download([target])
                    video_path = tmp_dl.absolute()
                else:
                    video_path = Path(target).absolute()
                
                if not video_path.exists():
                    self.log(f"❌ File not found: {video_path}")
                    continue
                
                # Process single video
                self.process_core(video_path, api_key, model_name, format_type, prompt)
                
                # Cleanup temp download
                if str(target).startswith("http") and video_path.exists():
                    video_path.unlink()
            
            self.status_label.configure(text="[ ALL CLEAR ]", text_color="#00FF41")
            self.after(0, lambda: self.progress_bar.set(1.0))
            if not self.stop_requested:
                messagebox.showinfo("AIdirector V14", "All missions complete! 🎬")
                
        except Exception as e:
            self.log(f"💥 ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.is_running = False
            self.stop_requested = False
            self.after(0, lambda: self.btn_run.configure(state="normal", text="DIRECT\nCUT"))
            self.after(0, lambda: self.sub_status.configure(text="Waiting for Target..."))
    
    def process_core(self, input_path, api_key, model_name, format_type, prompt):
        self.log(f"🎬 Processing: {input_path.name}")
        self.after(0, lambda: self.sub_status.configure(text=f"Processing: {input_path.name}"))
        
        # Create output filename
        output_video = self.output_dir / f"{input_path.stem}_Director_{datetime.now().strftime('%m%d_%H%M')}.mp4"
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        proxy_video = Path(f"temp_proxy_{uuid.uuid4().hex[:6]}.mp4")
        
        try:
            duration = self.get_video_duration(input_path)
            if duration:
                self.log(f"📊 Video length: {duration:.1f} seconds")
            
            # Phase 1: Create proxy
            self.log("📁 [Phase 1/4] Creating analysis proxy...")
            self.after(0, lambda: self.progress_bar.set(0.1))
            cmd_proxy = ["ffmpeg", "-y", "-i", str(input_path), "-vf", "scale=480:-2,fps=10", 
                        "-c:v", "libx264", "-crf", "30", "-preset", "ultrafast", "-an", str(proxy_video)]
            
            process = subprocess.Popen(cmd_proxy, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
            process.wait()
            if process.returncode != 0 or not proxy_video.exists():
                self.log("❌ Proxy creation failed")
                return
            
            # Phase 2: AI Analysis
            self.log(f"🧠 [Phase 2/4] Directing with {model_name}...")
            self.after(0, lambda: self.progress_bar.set(0.3))
            
            client = genai.Client(api_key=api_key)
            try:
                uploaded_file = client.files.upload(path=str(proxy_video))
                self.log("📤 Uploading file...")
                while uploaded_file.state == "PROCESSING":
                    if self.stop_requested:
                        client.files.delete(name=uploaded_file.name)
                        return
                    time.sleep(2)
                    uploaded_file = client.files.get(name=uploaded_file.name)
                
                self.log("🤔 AI is thinking about the structure...")
                response = client.models.generate_content(
                    model=model_name,
                    contents=[uploaded_file, prompt],
                    config=types.GenerateContentConfig(response_mime_type="application/json")
                )
                client.files.delete(name=uploaded_file.name)
                
                res_text = response.text.strip()
                if "```json" in res_text:
                    res_text = res_text.split("```json")[1].split("```")[0].strip()
                ai_data = json.loads(res_text)
                self.log(f"✨ Title suggestion: {ai_data.get('title', 'Untitled')}")
                strategy = ai_data.get('strategy', '')
                if strategy:
                    self.log(f"📝 Strategy: {strategy[:100]}...")
                
            except Exception as e:
                self.log(f"⚠️ API Error: {e}")
                return
            
            # Phase 3: Editing
            self.log(f"✂️ [Phase 3/4] Editing in progress (Format: {format_type})...")
            self.after(0, lambda: self.progress_bar.set(0.6))
            
            scenes = ai_data.get("scenes", {})
            temp_files = []
            
            if format_type == "9:16":
                vf = "scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920"
            else:
                vf = "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2"
            
            total_scenes = len(scenes)
            for idx, (key, scene) in enumerate(scenes.items()):
                if self.stop_requested:
                    break
                
                progress = 0.6 + (idx / total_scenes) * 0.3 if total_scenes > 0 else 0.8
                self.after(0, lambda p=progress: self.progress_bar.set(p))
                
                t_file = Path(f"temp_{key}_{uuid.uuid4().hex[:6]}.mp4")
                start = self.time_to_seconds(scene["start"])
                end = self.time_to_seconds(scene["end"])
                dur = max(0.5, end - start)
                
                cmd_cut = ["ffmpeg", "-y", "-ss", str(start), "-i", str(input_path), "-t", str(dur),
                          "-vf", vf, "-c:v", "libx264", "-preset", "fast", "-crf", "22",
                          "-c:a", "aac", "-b:a", "192k", str(t_file)]
                
                process = subprocess.Popen(cmd_cut, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                process.wait()
                if process.returncode == 0 and t_file.exists():
                    temp_files.append(str(t_file.absolute()))
                    self.log(f"  ✅ {key}: {scene['start']} → {scene['end']}")
            
            # Phase 4: Concatenation
            if temp_files and not self.stop_requested:
                self.log("🔗 [Phase 4/4] Final concatenation...")
                self.after(0, lambda: self.progress_bar.set(0.95))
                list_file = Path(f"list_{uuid.uuid4().hex[:6]}.txt")
                with open(list_file, "w") as f:
                    for tf in temp_files:
                        f.write(f"file '{tf}'\n")
                
                cmd_concat = ["ffmpeg", "-y", "-f", "concat", "-safe", "0", "-i", str(list_file), "-c", "copy", str(output_video)]
                process = subprocess.Popen(cmd_concat, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
                process.wait()
                
                if process.returncode == 0 and output_video.exists():
                    self.log(f"🎉 Complete! Saved as: {output_video}")
                    self.after(0, lambda: self.progress_bar.set(1.0))
                else:
                    self.log("❌ Concatenation failed")
                
                # Cleanup
                for tf in temp_files:
                    if os.path.exists(tf):
                        os.remove(tf)
                if list_file.exists():
                    list_file.unlink()
            
            # Cleanup proxy
            if proxy_video.exists():
                proxy_video.unlink()
                
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
        finally:
            self.after(0, lambda: self.sub_status.configure(text="Done"))

if __name__ == "__main__":
    app = DirectorV14Trial()
    app.mainloop()