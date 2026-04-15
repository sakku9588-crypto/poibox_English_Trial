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
import threading
import time
import json
import subprocess
import uuid
import hashlib
import platform
from pathlib import Path
from datetime import datetime, timedelta
import customtkinter as ctk
from tkinter import filedialog, messagebox, filedialog as tk_filedialog
from google import genai
from google.genai import types
import yt_dlp

# ==========================================
# ⚙️ Sniper V5.0 PRO - 30 Days Trial (Full Implementation)
# ==========================================
ANALYSIS_WIDTH = 426
ANALYSIS_HEIGHT = 240
ANALYSIS_FPS = 10
CHUNK_DURATION = 600      # 10 minutes per chunk
SPEED_FACTOR = 5.0        # 5x speed analysis
CHUNK_DELAY = 12
TRIAL_CONFIG_FILE = "sniper_trial.dat"  # Binary-like name but still JSON
API_KEY_FILE = "gemini_api_key.txt"
DEFAULT_MODEL = "gemini-2.0-flash"
AVAILABLE_MODELS = ["gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]

DEFAULT_PROMPT = """Analyze this video (5x speed, no audio). Find the most exciting/viral moments. Max 3 scenes.
Output format (JSON only): [{"proxy_start": seconds, "proxy_end": seconds, "score": 100, "reason": "why it's good"}]
Note: proxy_start/end are timestamps on this 5x video."""

class LicenseManager:
    def __init__(self):
        self.config_file = Path(TRIAL_CONFIG_FILE)
        self.trial_days = 30
        
    def get_hardware_id(self):
        """Generate a unique hardware ID"""
        system = platform.system()
        if system == "Windows":
            try:
                import subprocess
                result = subprocess.run(['wmic', 'csproduct', 'get', 'uuid'], capture_output=True, text=True)
                uuid_line = [line for line in result.stdout.split('\n') if 'UUID' not in line and line.strip()]
                if uuid_line:
                    return hashlib.sha256(uuid_line[0].strip().encode()).hexdigest()[:32]
            except:
                pass
        elif system == "Darwin":  # macOS
            try:
                result = subprocess.run(['system_profiler', 'SPHardwareDataType'], capture_output=True, text=True)
                for line in result.stdout.split('\n'):
                    if 'Hardware UUID' in line:
                        uuid_val = line.split(':')[-1].strip()
                        return hashlib.sha256(uuid_val.encode()).hexdigest()[:32]
            except:
                pass
        elif system == "Linux":
            try:
                with open('/etc/machine-id', 'r') as f:
                    machine_id = f.read().strip()
                    return hashlib.sha256(machine_id.encode()).hexdigest()[:32]
            except:
                pass
        
        # Fallback: hash the hostname + username
        hostname = platform.node()
        username = os.getlogin() if hasattr(os, 'getlogin') else 'unknown'
        combined = f"{hostname}_{username}"
        return hashlib.sha256(combined.encode()).hexdigest()[:32]
    
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
                # License copied to another machine!
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
        if len(license_key) >= 16 and license_key.startswith("SNIPER-"):
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
        self.title("Sniper V5.0 PRO - License")
        self.geometry("500x420")
        self.resizable(False, False)
        self.configure(fg_color="#1e1e1e")
        
        self.transient(parent)
        self.grab_set()
        
        self.create_widgets()
        
    def create_widgets(self):
        ctk.CTkLabel(self, text="🎯 SNIPER V5.0 PRO", font=("Impact", 28), text_color="#ff5555").pack(pady=30)
        
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
            ctk.CTkLabel(info_frame, text="Please purchase a license to continue using Sniper PRO", 
                        font=("Segoe UI", 12), text_color="#aaaaaa").pack()
        
        ctk.CTkLabel(self, text="License Key (optional)", font=("Segoe UI", 12)).pack(pady=(20,5))
        self.key_entry = ctk.CTkEntry(self, placeholder_text="SNIPER-XXXX-XXXX-XXXX", width=350, fg_color="#121212")
        self.key_entry.pack(pady=5)
        
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=30)
        
        if days_left > 0:
            ctk.CTkButton(btn_frame, text="Start Trial", command=self.start_trial,
                         fg_color="#00ff00", text_color="#000000", width=150, height=40,
                         font=("Impact", 14)).pack(side="left", padx=10)
        
        ctk.CTkButton(btn_frame, text="Activate License", command=self.activate_license,
                     fg_color="#ff5555", text_color="#ffffff", width=150, height=40,
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
            messagebox.showinfo("Success", "License activated successfully!\nThank you for purchasing Sniper PRO!")
            self.result = "licensed"
            self.destroy()
        else:
            messagebox.showerror("Error", "Invalid license key.\nPlease check and try again.")
    
    def purchase(self):
        import webbrowser
        webbrowser.open("https://example.com/purchase")
        messagebox.showinfo("Purchase", "Purchase page opened in your browser.")
    
    def exit_app(self):
        self.result = "exit"
        self.destroy()

class SniperV5Trial(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # License check first
        self.license_mgr = LicenseManager()
        if not self.check_license():
            return
        
        self.title("🎯 SNIPER V5.0 PRO - 30 Day Trial")
        self.geometry("1300x950")
        ctk.set_appearance_mode("dark")
        
        self.is_running = False
        self.stop_requested = False
        self.selected_files = []
        self.output_dir = Path.cwd() / "SNIPER_OUTPUT"
        
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
                                "Your 30-day trial has expired.\nPlease purchase a license to continue using Sniper PRO.")
            self.destroy()
            return False
        
        return True
    
    def update_title_with_days(self):
        days_left = self.license_mgr.get_days_left()
        if days_left == -1:
            title_suffix = "LICENSED"
        else:
            title_suffix = f"TRIAL: {days_left} days left"
        self.title(f"🎯 SNIPER V5.0 PRO - {title_suffix}")
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
    
    def get_video_duration(self, path):
        cmd = ['ffprobe', '-v', 'error', '-show_entries', 'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', str(path)]
        result = subprocess.run(cmd, capture_output=True, text=True)
        return float(result.stdout.strip())
    
    def setup_obs_ui(self):
        # ========== LEFT: SOURCES ==========
        self.left_frame = ctk.CTkFrame(self, corner_radius=0, border_width=1, border_color="#333333")
        self.left_frame.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)
        ctk.CTkLabel(self.left_frame, text="SOURCES", font=("Impact", 18), text_color="#ff5555").pack(pady=10)
        
        self.file_list_box = ctk.CTkTextbox(self.left_frame, fg_color="#121212", font=("Consolas", 11))
        self.file_list_box.pack(fill="both", expand=True, padx=5, pady=5)
        
        ctk.CTkButton(self.left_frame, text="+ LOCAL FILES", command=self.browse_files, fg_color="#333333").pack(fill="x", padx=10, pady=2)
        
        self.url_entry = ctk.CTkEntry(self.left_frame, placeholder_text="YouTube URL...", fg_color="#121212")
        self.url_entry.pack(fill="x", padx=10, pady=5)
        
        # Output folder selection
        ctk.CTkButton(self.left_frame, text="📁 OUTPUT FOLDER", command=self.select_output_folder, fg_color="#333333").pack(fill="x", padx=10, pady=2)
        self.output_label = ctk.CTkLabel(self.left_frame, text=str(self.output_dir), font=("Consolas", 9), text_color="#888888")
        self.output_label.pack(pady=2)
        
        # ========== CENTER: SNIPER'S EYE ==========
        self.center_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="#000000", border_width=1, border_color="#ff5555")
        self.center_frame.grid(row=0, column=1, sticky="nsew", padx=2, pady=2)
        
        self.status_label = ctk.CTkLabel(self.center_frame, text="[ SYSTEM READY ]", font=("Impact", 42), text_color="#ff5555")
        self.status_label.place(relx=0.5, rely=0.4, anchor="center")
        
        self.sub_status = ctk.CTkLabel(self.center_frame, text="Waiting for Target...", font=("Consolas", 14), text_color="#AAAAAA")
        self.sub_status.place(relx=0.5, rely=0.5, anchor="center")
        
        self.progress_bar = ctk.CTkProgressBar(self.center_frame, progress_color="#ff5555", height=15, width=600)
        self.progress_bar.place(relx=0.5, rely=0.9, anchor="center")
        self.progress_bar.set(0)
        
        # ========== RIGHT: CONTROLS ==========
        self.right_frame = ctk.CTkFrame(self, corner_radius=0, border_width=1, border_color="#333333")
        self.right_frame.grid(row=0, column=2, sticky="nsew", padx=2, pady=2)
        
        ctk.CTkLabel(self.right_frame, text="CONTROLS", font=("Impact", 18)).pack(pady=10)
        
        self.api_entry = ctk.CTkEntry(self.right_frame, placeholder_text="API Keys (comma separated)", show="*", fg_color="#121212")
        self.api_entry.pack(fill="x", padx=10, pady=5)
        
        self.model_menu = ctk.CTkOptionMenu(self.right_frame, values=AVAILABLE_MODELS, fg_color="#333333")
        self.model_menu.pack(fill="x", padx=10, pady=10)
        self.model_menu.set(DEFAULT_MODEL)
        
        self.format_var = ctk.StringVar(value="9:16")
        ctk.CTkRadioButton(self.right_frame, text="Shorts (9:16)", variable=self.format_var, value="9:16").pack(pady=5, padx=20, anchor="w")
        ctk.CTkRadioButton(self.right_frame, text="Wide (16:9)", variable=self.format_var, value="16:9").pack(pady=5, padx=20, anchor="w")
        
        ctk.CTkLabel(self.right_frame, text="📝 EDITABLE PROMPT", font=("Consolas", 12), text_color="#aaaaaa").pack(pady=(10,0))
        self.prompt_text = ctk.CTkTextbox(self.right_frame, height=180, fg_color="#121212", font=("Consolas", 10))
        self.prompt_text.pack(fill="x", padx=10, pady=5)
        self.prompt_text.insert("1.0", DEFAULT_PROMPT)
        
        self.btn_run = ctk.CTkButton(self.right_frame, text="MASTER\nSNIPE", font=("Impact", 32), fg_color="#ff5555", height=150, command=self.launch)
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
    
    def launch(self):
        if self.is_running:
            return
        self.is_running = True
        self.stop_requested = False
        self.status_label.configure(text="[ SNIPING NOW ]", text_color="#00FF00")
        self.btn_run.configure(state="disabled", text="RUNNING")
        threading.Thread(target=self.batch_process, daemon=True).start()
    
    def split_video(self, path, chunk_sec=600):
        """Split video into chunks for analysis"""
        self.log(f"📏 Splitting into {chunk_sec//60}min chunks...")
        duration = self.get_video_duration(path)
        chunks = []
        for start in range(0, int(duration), chunk_sec):
            if self.stop_requested:
                break
            c_path = Path(f"temp_chunk_{start}_{uuid.uuid4().hex[:4]}.mp4")
            subprocess.run(['ffmpeg', '-y', '-ss', str(start), '-t', str(chunk_sec), '-i', str(path), '-c', 'copy', str(c_path)], capture_output=True)
            if c_path.exists():
                chunks.append({"path": str(c_path.absolute()), "offset": start})
        return chunks
    
    def prepare_analysis_video(self, path):
        """Create 5x speed proxy for analysis"""
        px_path = Path(f"temp_proxy_{uuid.uuid4().hex[:6]}.mp4")
        vf = f"silenceremove=stop_periods=-1:stop_duration=1:stop_threshold=-30dB,setpts={1/SPEED_FACTOR}*PTS,scale={ANALYSIS_WIDTH}:{ANALYSIS_HEIGHT}"
        subprocess.run(['ffmpeg', '-y', '-i', str(path), '-vf', vf, '-an', '-r', str(ANALYSIS_FPS), '-crf', '32', '-preset', 'ultrafast', str(px_path)], capture_output=True)
        return px_path
    
    def analyze_chunk(self, api_key, px_path, offset, prompt, model_name):
        """Analyze a video chunk with Gemini"""
        self.log(f"🧠 {model_name} analyzing chunk (offset: {offset}s)...")
        try:
            client = genai.Client(api_key=api_key)
            uploaded_file = client.files.upload(path=str(px_path))
            while uploaded_file.state == "PROCESSING":
                time.sleep(2)
                uploaded_file = client.files.get(name=uploaded_file.name)
            
            response = client.models.generate_content(
                model=model_name,
                contents=[uploaded_file, prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json")
            )
            client.files.delete(name=uploaded_file.name)
            
            res_text = response.text.strip()
            if "```json" in res_text:
                res_text = res_text.split("```json")[1].split("```")[0].strip()
            scenes = json.loads(res_text)
            
            for s in scenes:
                s['real_start'] = float(s['proxy_start']) * SPEED_FACTOR + offset
                s['real_end'] = float(s['proxy_end']) * SPEED_FACTOR + offset
            return scenes
        except Exception as e:
            self.log(f"⚠️ API Error: {str(e)[:80]}")
            return []
    
    def render_output(self, path, scenes, format_type):
        """Render final snipped clips"""
        out_dir = self.output_dir / f"SNIPED_{datetime.now().strftime('%m%d_%H%M')}"
        out_dir.mkdir(parents=True, exist_ok=True)
        self.log(f"🎬 Rendering {len(scenes)} clips...")
        
        if format_type == "9:16":
            vf = "scale=1080:1920:force_original_aspect_ratio=increase,crop=1080:1920"
        else:
            vf = "scale=1920:1080:force_original_aspect_ratio=decrease,pad=1920:1080:(ow-iw)/2:(oh-ih)/2"
        
        for i, sc in enumerate(scenes):
            if self.stop_requested:
                break
            out_n = out_dir / f"Snipe_{i+1:02d}_{int(sc.get('score', 0))}pts.mp4"
            st = max(0, sc['real_start'] - 2)
            et = sc['real_end'] + 2
            duration = et - st
            
            cmd = ['ffmpeg', '-y', '-ss', f"{st:.3f}", '-i', str(path), '-t', f"{duration:.3f}",
                   '-vf', vf, '-c:v', 'libx264', '-preset', 'fast', '-crf', '20',
                   '-c:a', 'aac', '-b:a', '192k', str(out_n)]
            subprocess.run(cmd, capture_output=True)
            if out_n.exists():
                self.log(f"   ∟ ✅ {out_n.name}")
    
    def batch_process(self):
        try:
            targets = self.selected_files if self.selected_files else []
            url = self.url_entry.get().strip()
            if url:
                targets.append(url)
            
            if not targets:
                self.log("❌ Select video or enter URL")
                return
            
            api_keys = [k.strip() for k in self.api_entry.get().split(",") if k.strip()]
            if not api_keys:
                self.log("❌ No API keys provided")
                return
            
            model_name = self.model_menu.get()
            format_type = self.format_var.get()
            prompt = self.prompt_text.get("1.0", "end").strip()
            
            total_targets = len(targets)
            for target_idx, target in enumerate(targets):
                if self.stop_requested:
                    break
                
                # Update main progress
                main_progress = target_idx / total_targets
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
                self.log(f"🎯 Target acquired: {video_path.name}")
                self.after(0, lambda: self.sub_status.configure(text=f"Processing: {video_path.name}"))
                
                duration = self.get_video_duration(video_path)
                self.log(f"📊 Video length: {duration/60:.1f} minutes")
                
                # Split into chunks
                chunks = self.split_video(video_path, CHUNK_DURATION)
                all_scenes = []
                
                for chunk_idx, chunk in enumerate(chunks):
                    if self.stop_requested:
                        break
                    
                    # Update chunk progress (30% of progress for analysis)
                    chunk_progress = 0.3 + (chunk_idx / len(chunks)) * 0.5
                    self.after(0, lambda p=chunk_progress: self.progress_bar.set(p))
                    self.after(0, lambda idx=chunk_idx+1, total=len(chunks): 
                              self.sub_status.configure(text=f"Analyzing chunk {idx}/{total}..."))
                    
                    chunk_path = Path(chunk['path'])
                    px_path = self.prepare_analysis_video(chunk_path)
                    
                    # Try each API key
                    scenes = []
                    for api_key in api_keys:
                        scenes = self.analyze_chunk(api_key, px_path, chunk['offset'], prompt, model_name)
                        if scenes:
                            break
                        time.sleep(2)
                    
                    for s in scenes:
                        self.log(f"✅ Locked! ({s.get('score', 0)}pts): {s.get('reason', '')[:60]}")
                    all_scenes.extend(scenes)
                    
                    # Cleanup
                    if px_path.exists():
                        px_path.unlink()
                    if chunk_path.exists():
                        chunk_path.unlink()
                    
                    if not self.stop_requested:
                        self.log(f"😴 Waiting {CHUNK_DELAY}s to avoid quota...")
                        time.sleep(CHUNK_DELAY)
                
                # Render output
                if all_scenes:
                    all_scenes.sort(key=lambda x: x.get('score', 0), reverse=True)
                    self.after(0, lambda p=0.9: self.progress_bar.set(p))
                    self.render_output(video_path, all_scenes[:5], format_type)
                    self.log(f"✅ Completed: {video_path.name}")
                else:
                    self.log(f"⚠️ No interesting scenes found in {video_path.name}")
                
                # Cleanup temp files
                if str(target).startswith("http") and video_path.exists():
                    video_path.unlink()
            
            self.status_label.configure(text="[ ALL CLEAR ]", text_color="#ff5555")
            self.after(0, lambda: self.progress_bar.set(1.0))
            if not self.stop_requested:
                messagebox.showinfo("Sniper V5", "All missions complete! 🎯")
                
        except Exception as e:
            self.log(f"💥 ERROR: {e}")
            import traceback
            self.log(traceback.format_exc())
        finally:
            self.is_running = False
            self.stop_requested = False
            self.after(0, lambda: self.btn_run.configure(state="normal", text="MASTER\nSNIPE"))
            self.after(0, lambda: self.sub_status.configure(text="Waiting for Target..."))

if __name__ == "__main__":
    app = SniperV5Trial()
    app.mainloop()
