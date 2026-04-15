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
import sys
import tkinter as tk
import traceback
import time
import csv
from tkinter import ttk, scrolledtext, messagebox, filedialog
from datetime import datetime, timedelta, timezone
import pytchat
from googleapiclient.discovery import build
from flask import Flask
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import desc, func, text, inspect as sa_inspect, case

# --- Gemini API (official SDK) ---
from google import genai

# --- Local Text-to-Speech Library ---
try:
    import pyttsx3
except ImportError:
    pyttsx3 = None

# ==========================================
# Path Resolution Utilities
# ==========================================
def get_resource_path(relative_path):
    """Resolve path for PyInstaller packaged apps vs normal execution"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_data_path(relative_path):
    """Resolve path for data storage (same directory as executable)"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ==========================================
# 🐾 Trial Management (5 launches / 14 days limit)
# ==========================================
TRIAL_FILE      = get_data_path("ananeko_trial.json")
TRIAL_MAX_DAYS  = 14
TRIAL_MAX_BOOTS = 5

def check_trial():
    """Check trial status, return False if limits exceeded"""
    today = datetime.now().date()
    data  = {}

    if os.path.exists(TRIAL_FILE):
        try:
            with open(TRIAL_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except Exception:
            data = {}

    # Record first launch date
    if "first_launch" not in data:
        data["first_launch"] = today.isoformat()
        data["boot_count"]   = 0

    first_day   = datetime.fromisoformat(data["first_launch"]).date()
    days_passed = (today - first_day).days
    boot_count  = data.get("boot_count", 0)

    # Limit check
    if days_passed >= TRIAL_MAX_DAYS:
        _show_trial_expired(f"Trial period ({TRIAL_MAX_DAYS} days) has expired.\nPlease purchase the full version.")
        return False
    if boot_count >= TRIAL_MAX_BOOTS:
        _show_trial_expired(f"Trial launch count ({TRIAL_MAX_BOOTS} times) has been reached.\nPlease purchase the full version.")
        return False

    # Update and save boot count
    data["boot_count"] = boot_count + 1
    try:
        with open(TRIAL_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

    return True

def _show_trial_expired(msg):
    """Display trial expiration dialog"""
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _root = _tk.Tk()
    _root.withdraw()
    _mb.showwarning(
        "⏰ Trial Limit Reached",
        f"【PoiBox Sakuneko Analysis - Trial Version】\n\n{msg}\n\n"
        "Thank you for trying! 🐾"
    )
    _root.destroy()

# ==========================================
# 1. Basic Settings & Database
# ==========================================
SETTINGS_FILE = get_data_path("settings.json")
MEMORY_FILE   = get_data_path("gemineko_memory.json")
DB_NAME       = "sakuneko_v9_9_3.db"
DB_PATH       = get_data_path(DB_NAME)
LOG_DIR       = get_data_path("log")

def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception as e:
        print(f"Settings load error: {e}")
    return {
        "start_date": (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%d"), 
        "end_date": datetime.now().strftime("%Y-%m-%d"), 
        "gemini_api_key": "",
        "api_key": "",
        "handle": "JemiNeko"
    }

def save_settings(data):
    current = load_settings()
    current.update(data)
    try:
        with open(SETTINGS_FILE, 'w', encoding='utf-8') as f:
            json.dump(current, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Settings save error: {e}")

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = f'sqlite:///{DB_PATH}'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)

class PointLog(db.Model):
    id          = db.Column(db.Integer, primary_key=True)
    username    = db.Column(db.String(500), nullable=False, index=True)
    message     = db.Column(db.Text)
    reason      = db.Column(db.String(500))
    video_id    = db.Column(db.String(255))
    event_type  = db.Column(db.String(50), default="chat")
    point_value = db.Column(db.Integer, default=0)
    date        = db.Column(db.DateTime, default=datetime.now, index=True)

def migrate_db():
    with app.app_context():
        db.create_all()
        inspector = sa_inspect(db.engine)
        columns = [col['name'] for col in inspector.get_columns('point_log')]
        if 'event_type' not in columns:
            with db.engine.connect() as conn:
                conn.execute(text('ALTER TABLE point_log ADD COLUMN event_type VARCHAR(50) DEFAULT "chat"'))
                conn.commit()
        if 'point_value' not in columns:
            with db.engine.connect() as conn:
                conn.execute(text('ALTER TABLE point_log ADD COLUMN point_value INTEGER DEFAULT 0'))
                conn.commit()

# ==========================================
# 2. Log Analysis Utilities
# ==========================================

def clean_name(name):
    if not name:
        return None
    name = name.strip()
    name = re.split(r'[\s\r\n\t]|san[ni]?|[:：\(\)（）]', name)[0]
    if not name.startswith('@'):
        name = '@' + name
    return name if len(name) > 1 else None

def extract_video_title_from_log(filepath, logger=None):
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            header = "".join(f.readlines()[:20])
        t_match = re.search(r'📺 Title:\s*(.*)', header)
        if t_match:
            return t_match.group(1).strip()
    except Exception as e:
        if logger:
            logger(f"⚠️ Title extraction error: {filepath} / {e}")
    return None

def auto_scan_logs(logger=None, channel_handle=None):
    """Scan txt files in log folder and import into database.
    If channel_handle is specified, skip logs that don't contain the handle in the title.
    Old logs without title lines are skipped.
    """
    if not os.path.exists(LOG_DIR):
        if logger:
            logger("📁 log folder not found. Creating new one.")
        os.makedirs(LOG_DIR, exist_ok=True)
        return 0
    
    if logger:
        logger("🔍 Scanning local logs...")
    
    # Normalize channel handles (with/without @)
    normalized_handles = []
    if channel_handle:
        h = channel_handle.strip().lower()
        if h:
            normalized_handles.append(h)
            if h.startswith('@'):
                normalized_handles.append(h[1:])
            else:
                normalized_handles.append('@' + h)
    
    total_added = 0
    processed_files = 0
    skipped_files = 0
    
    with app.app_context():
        for fname in os.listdir(LOG_DIR):
            if not fname.endswith('.txt'):
                continue
            
            fpath = os.path.join(LOG_DIR, fname)
            video_title = extract_video_title_from_log(fpath, logger)
            
            if video_title and normalized_handles:
                title_lower = video_title.lower()
                if not any(h in title_lower for h in normalized_handles):
                    if logger:
                        logger(f"⏭️ Skip: {fname} (Not your channel's log: {video_title[:40]}...)")
                    skipped_files += 1
                    continue
            
            if not video_title:
                if logger:
                    logger(f"⏭️ Skip: {fname} (Video title not found)")
                skipped_files += 1
                continue
            
            final_reason = f"{video_title} (LOG)"
            
            try:
                with open(fpath, 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                
                existing = db.session.query(PointLog).filter(PointLog.reason == final_reason).first()
                
                bulk_points = []
                bulk_chats = []
                registered_users = set()
                
                for line in lines:
                    p_matches = re.finditer(r'(@[^\s\(\)（）:：]+).*?([\+\d,]+)\s*pt', line)
                    for pm in p_matches:
                        name = clean_name(pm.group(1))
                        if name and name not in ["@Sakuneko", "@System", "@SakunekoBot"]:
                            pts = int(re.sub(r'[^\d]', '', pm.group(2)))
                            
                            if '✨' in line and 'NEW' in line:
                                if name not in registered_users:
                                    registered_users.add(name)
                                    bulk_points.append(PointLog(
                                        username=name, 
                                        message=line.strip(),
                                        reason=final_reason,
                                        point_value=pts,
                                        event_type="system"
                                    ))
                            elif '🎉' in line:
                                pass
                            elif '💎' in line:
                                bulk_points.append(PointLog(
                                    username=name, 
                                    message=line.strip(),
                                    reason=final_reason,
                                    point_value=pts,
                                    event_type="system"
                                ))
                            else:
                                bulk_points.append(PointLog(
                                    username=name, 
                                    message=line.strip(),
                                    reason=final_reason,
                                    point_value=pts,
                                    event_type="system"
                                ))
                    
                    if not any(emoji in line for emoji in ["💎", "✨", "🎉", "💰", "👑", "🎲"]):
                        if line.strip().startswith('💬'):
                            c_matches = re.finditer(r'💬.*?(@[^\s:：]+)', line)
                            for cm in c_matches:
                                name = clean_name(cm.group(1))
                                if name and name not in ["@Sakuneko", "@System", "@SakunekoBot"]:
                                    bulk_chats.append(PointLog(
                                        username=name,
                                        message=line.strip(),
                                        reason=final_reason,
                                        point_value=0,
                                        event_type="chat"
                                    ))
                
                if bulk_points or bulk_chats:
                    if existing:
                        db.session.query(PointLog).filter(PointLog.reason == final_reason).delete()
                    if bulk_points:
                        db.session.bulk_save_objects(bulk_points)
                    if bulk_chats:
                        db.session.bulk_save_objects(bulk_chats)
                    db.session.commit()
                    total_added += len(bulk_points) + len(bulk_chats)
                    processed_files += 1
                    if logger:
                        logger(f"📄 {fname}: Title「{video_title[:40]}」")
                        logger(f"     Points: {len(bulk_points)}, Chats: {len(bulk_chats)}")
                        
            except Exception as e:
                if logger:
                    logger(f"❌ File read error: {fname} / {e}")
    
    if logger:
        logger(f"✅ Log scan complete: {processed_files} processed, {skipped_files} skipped, {total_added} total records")
    return total_added

# ==========================================
# 3. Gemini Strategist Brain
# ==========================================
class GeminekoBrain:
    def __init__(self, logger, liver_name="JemiNeko"):
        self.log = logger
        self.liver_name = liver_name
        self.engine = None
        if pyttsx3:
            try:
                self.engine = pyttsx3.init()
                self._setup_voice()
            except Exception as e:
                self.log(f"⚠️ TTS initialization failed: {e}")
                self.engine = None
        self.memory = self._load_memory()
        self.gemini_client = None
        self.init_gemini_client()

    def init_gemini_client(self):
        key = load_settings().get("gemini_api_key", "")
        if key:
            try:
                self.gemini_client = genai.Client(api_key=key)
                self.log("✅ Gemini API client initialized")
            except Exception as e:
                self.log(f"⚠️ Gemini API initialization failed: {e}")
                self.gemini_client = None
        else:
            self.log("⚠️ Gemini API key not set")

    def _setup_voice(self):
        if not self.engine:
            return
        try:
            for v in self.engine.getProperty('voices'):
                if "JP" in v.id or "Japanese" in v.name or "en" in v.id:
                    self.engine.setProperty('voice', v.id)
                    break
            self.engine.setProperty('rate', 165)
            self.engine.setProperty('volume', 0.9)
        except Exception as e:
            self.log(f"⚠️ Voice setup error: {e}")

    def _load_memory(self):
        if os.path.exists(MEMORY_FILE):
            try:
                with open(MEMORY_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                self.log(f"⚠️ Memory load error: {e}")
        return {"consultations": 0, "last_active": "", "history": [], "last_advice": ""}

    def _save_memory(self):
        self.memory["last_active"] = datetime.now().isoformat()
        try:
            with open(MEMORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.memory, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.log(f"⚠️ Memory save error: {e}")

    def speak(self, text):
        if self.engine:
            def _speak():
                try:
                    self.engine.say(text)
                    self.engine.runAndWait()
                except Exception as e:
                    self.log(f"⚠️ Speech error: {e}")
            threading.Thread(target=_speak, daemon=True).start()

    def _analyze_keywords(self, logs):
        all_text = " ".join([l.message for l in logs[:500] if l.message])
        words = re.findall(r'[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAFa-z]+', all_text)
        wc = {}
        for w in words:
            if len(w) > 1:
                wc[w] = wc.get(w, 0) + 1
        return sorted(wc.items(), key=lambda x: x[1], reverse=True)[:10]

    def _analyze_time_pattern(self, logs):
        hours = [l.date.hour for l in logs if l.date]
        if not hours:
            return None, {}
        hc = {}
        for h in hours:
            hc[h] = hc.get(h, 0) + 1
        return max(hc, key=hc.get), hc

    def _analyze_sentiment(self, logs):
        pos_words = ["happy","great","love","awesome","thank","good","nice","fun","cool","wow"]
        neg_words = ["sad","tired","hard","bad","boring","hate","sorry","miss","pain"]
        pos = neg = 0
        for l in logs[:200]:
            msg = (l.message or "").lower()
            if any(w in msg for w in pos_words):
                pos += 1
            if any(w in msg for w in neg_words):
                neg += 1
        return pos, neg

    def generate_gemini_analysis(self, total_chats, top_stream, top_fan, recent_logs):
        if not self.gemini_client:
            return self.generate_local_report(total_chats, top_stream, top_fan)
        
        keywords = self._analyze_keywords(recent_logs)
        peak_hour, hour_counts = self._analyze_time_pattern(recent_logs)
        pos, neg = self._analyze_sentiment(recent_logs)
        samples = "\n".join(
            f"[{l.date.strftime('%m/%d %H:%M')}] {l.username}: {(l.message or '')[:60]}"
            for l in recent_logs[:50] if l.date and l.username
        ) or "No data"
        
        display_name = self.liver_name if self.liver_name != "JemiNeko" else "JemiNeko"
        
        prompt = f"""You are "Strategist JemiNeko", a YouTube streamer's strategic advisor.
Please analyze the following data and provide strategic advice for {display_name} to create better streams.

## Basic Data
- Total chat count: {total_chats}
- Analysis date: {datetime.now().strftime('%Y/%m/%d')}

## Top Stream
{"「" + top_stream[0] + "」 (Chats: " + str(top_stream[1]) + ")" if top_stream else "No data"}

## Top Fan
{"@" + top_fan[0] + " (Points: " + str(top_fan[1]) + ")" if top_fan else "No data"}

## Keyword Analysis
{"".join(f"- 「{kw}」: {cnt} times\n" for kw,cnt in keywords) or "No data"}

## Peak Activity Time
{"Peak hour: " + str(peak_hour) + ":00" if peak_hour is not None else "No data"}

## Sentiment Analysis
- Positive: {pos} / Negative: {neg}
- Positive rate: {(pos/(pos+neg)*100) if pos+neg>0 else 0:.1f}%

## Recent Chat Samples (latest 50)
{samples}

Based on the above, create an analysis report with:
- Trend analysis / Viewer analysis / Improvement suggestions / Next stream advice / Closing message
- Finally, a "Summary" with a 10-point rating and closing comment
- Address {display_name} as "{display_name}"

【Output Format】
📊 **Data Summary**
🔍 **Detailed Analysis**
💡 **Strategic Advice**
📈 **Summary**"""
        
        try:
            res = self.gemini_client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt,
                config={"temperature": 0.7, "max_output_tokens": 8192}
            )
            report = res.text
            sp_res = self.gemini_client.models.generate_content(
                model="gemini-2.0-flash",
                contents=f"Summarize the following report in a way that can be read in 20 seconds:\n\n{report[:2000]}",
                config={"temperature": 0.5, "max_output_tokens": 300}
            )
            return report, sp_res.text
        except Exception as e:
            self.log(f"⚠️ Gemini API error: {e}, falling back to local analysis")
            return self.generate_local_report(total_chats, top_stream, top_fan)

    def generate_local_report(self, total_chats, top_stream, top_fan):
        self.memory["consultations"] += 1
        c = self.memory["consultations"]
        display_name = self.liver_name if self.liver_name != "JemiNeko" else "JemiNeko"
        
        lines = [
            f"🐾 Strategist JemiNeko's Report (Consultations: {c})",
            "=" * 65,
            f"📊 **Total Chats**: {total_chats}",
        ]
        if top_stream:
            lines.append(f"🔥 **Top Stream**: 「{top_stream[0][:60]}」 ({top_stream[1]} chats)")
        if top_fan:
            lines.append(f"👑 **Top Fan**: @{top_fan[0]} ({top_fan[1]} pts)")
        
        with app.app_context():
            recent = PointLog.query.filter_by(event_type="chat").order_by(desc(PointLog.date)).limit(400).all()
            if len(recent) > 5:
                kw = self._analyze_keywords(recent)
                if kw:
                    lines.append("\n🔑 **Frequent Keywords**")
                    for w, n in kw[:7]:
                        lines.append(f"   ・「{w}」: {n} times")
                ph, _ = self._analyze_time_pattern(recent)
                if ph is not None:
                    lines.append(f"\n⏰ **Peak Activity**: {ph}:00")
                pos, neg = self._analyze_sentiment(recent)
                if pos + neg > 0:
                    lines.append(f"\n😺 **Positive Rate**: {pos/(pos+neg)*100:.1f}% ({pos}/{pos+neg})")
        
        lines += [
            "\n" + "=" * 65,
            f"【Strategist JemiNeko's Advice for {display_name}】",
            f"・{display_name}, don't forget to welcome new viewers! Greeting them within the first 10 minutes increases retention.",
            "・Responding by name builds a stronger connection.",
            "・Take a little extra time at the end to express your gratitude.",
        ]
        if self.memory.get("last_advice"):
            lines.append(f"\n※Previous advice: {self.memory['last_advice'][:80]}…")
        advice = "\n".join(lines[-3:])
        self.memory["last_advice"] = advice[:300]
        self._save_memory()
        return "\n".join(lines), f"{display_name}, taking care of new viewers will bring you good luck!"

# ==========================================
# 4. Main GUI
# ==========================================
class ArchiveImporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🐾 PoiBox Sakuneko Analysis v15.6 (Gemini Strategist Edition)")
        self.root.geometry("1300x900")
        self.root.configure(bg="#f0f2f5")
        
        self.settings = load_settings()
        self.api_usage = 0
        self.sort_column_name = ""
        self.sort_reverse = False
        
        migrate_db()
        self.setup_ui()
        
        liver_name = self.ch_handle_ent.get().strip()
        if not liver_name:
            liver_name = "JemiNeko"
        self.brain = GeminekoBrain(self.log, liver_name)
        
        self.refresh_analytics()
        
        threading.Thread(target=self.initial_load, daemon=True).start()

    def setup_ui(self):
        main_paned = tk.PanedWindow(self.root, orient=tk.HORIZONTAL, bg="#f0f2f5", sashwidth=6)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # ========== Left Panel ==========
        left_frame = tk.Frame(main_paned, bg="#f0f2f5")
        main_paned.add(left_frame, width=480)
        
        settings_card = tk.Frame(left_frame, bg="white", relief="flat", bd=0)
        settings_card.pack(fill="x", pady=(0, 8))
        
        tk.Label(settings_card, text="⚙️ System Settings", font=("Segoe UI", 12, "bold"), bg="white", fg="#333").pack(anchor="w", padx=10, pady=(8, 0))
        
        self.lbl_usage = tk.Label(settings_card, text="💰 API Usage: 0 pts", fg="#d32f2f", font=("", 10, "bold"), bg="white")
        self.lbl_usage.pack(anchor="w", padx=10, pady=(5, 0))
        
        row1 = tk.Frame(settings_card, bg="white")
        row1.pack(fill="x", padx=10, pady=5)
        tk.Label(row1, text="YouTube API:", bg="white", width=12, anchor="w").pack(side="left")
        self.api_key_ent = tk.Entry(row1, show="*", width=33)
        self.api_key_ent.insert(0, self.settings.get("api_key", ""))
        self.api_key_ent.pack(side="left", fill="x", expand=True)
        
        row2 = tk.Frame(settings_card, bg="white")
        row2.pack(fill="x", padx=10, pady=3)
        tk.Label(row2, text="Handle (@):", bg="white", width=12, anchor="w").pack(side="left")
        self.ch_handle_ent = tk.Entry(row2, width=33)
        self.ch_handle_ent.insert(0, self.settings.get("handle", ""))
        self.ch_handle_ent.pack(side="left", fill="x", expand=True)
        
        row3 = tk.Frame(settings_card, bg="white")
        row3.pack(fill="x", padx=10, pady=3)
        tk.Label(row3, text="Gemini API:", bg="white", width=12, anchor="w").pack(side="left")
        self.gemini_key_ent = tk.Entry(row3, show="*", width=33)
        self.gemini_key_ent.insert(0, self.settings.get("gemini_api_key", ""))
        self.gemini_key_ent.pack(side="left", fill="x", expand=True)
        
        row4 = tk.Frame(settings_card, bg="white")
        row4.pack(fill="x", padx=10, pady=5)
        tk.Label(row4, text="Date Range:", bg="white", width=12, anchor="w").pack(side="left")
        self.start_date_ent = tk.Entry(row4, width=12)
        self.start_date_ent.insert(0, self.settings.get("start_date", (datetime.now() - timedelta(days=14)).strftime("%Y-%m-%d")))
        self.start_date_ent.pack(side="left", padx=(0, 5))
        tk.Label(row4, text="to", bg="white").pack(side="left")
        self.end_date_ent = tk.Entry(row4, width=12)
        self.end_date_ent.insert(0, self.settings.get("end_date", datetime.now().strftime("%Y-%m-%d")))
        self.end_date_ent.pack(side="left", padx=5)
        tk.Button(row4, text="💾Save", command=self.save_api_keys, bg="#4caf50", fg="white", font=("", 9), padx=10).pack(side="right")
        
        tk.Button(settings_card, text="🔍 Fetch Archive URLs", command=self.fetch_vids, bg="#ffc107", font=("", 10, "bold"), height=1).pack(fill="x", padx=10, pady=5)
        
        url_frame = tk.LabelFrame(left_frame, text="📺 Archive URL List", bg="white", fg="#555", font=("", 9))
        url_frame.pack(fill="both", expand=True, pady=5)
        self.url_list_text = scrolledtext.ScrolledText(url_frame, height=7, font=("Consolas", 9))
        self.url_list_text.pack(fill="both", expand=True, padx=3, pady=3)
        
        tk.Button(url_frame, text="🚀 Start Analysis", command=self.start_analysis, bg="#e53935", fg="white", font=("", 11, "bold"), height=1).pack(fill="x", padx=5, pady=5)
        
        log_frame = tk.LabelFrame(left_frame, text="📝 Activity Log", bg="white", fg="#555", font=("", 9))
        log_frame.pack(fill="x", pady=5)
        self.log_area = scrolledtext.ScrolledText(log_frame, height=10, bg="#111", fg="#0f0", font=("Consolas", 8))
        self.log_area.pack(fill="both", expand=True, padx=3, pady=3)
        
        # ========== Right Panel ==========
        right_frame = tk.Frame(main_paned, bg="#f0f2f5")
        main_paned.add(right_frame, width=780)
        
        search_frame = tk.Frame(right_frame, bg="white", relief="flat", bd=0)
        search_frame.pack(fill="x", pady=(0, 5))
        tk.Label(search_frame, text="🔎 Search Viewer:", bg="white", font=("", 9)).pack(side="left", padx=10)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *a: self.refresh_analytics())
        self.search_ent = tk.Entry(search_frame, textvariable=self.search_var, width=25)
        self.search_ent.pack(side="left", padx=5, pady=5)
        
        btn_frame = tk.Frame(right_frame, bg="#f0f2f5")
        btn_frame.pack(fill="x", pady=5)
        tk.Button(btn_frame, text="🔄 Refresh & Load Logs", command=self.manual_log_scan, bg="#ff5722", fg="white", font=("", 9)).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(btn_frame, text="🎙️ Ask Strategist", command=self.consult_strategist, bg="#673ab7", fg="white", font=("", 9)).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(btn_frame, text="📂 Export CSV", command=self.export_rank_csv, bg="#607d8b", fg="white", font=("", 9)).pack(side="left", fill="x", expand=True, padx=2)
        tk.Button(btn_frame, text="⚠️ Reset DB", command=self.reset_database, bg="#333", fg="white", font=("", 9)).pack(side="left", fill="x", expand=True, padx=2)
        
        tree_container = tk.Frame(right_frame, bg="white")
        tree_container.pack(fill="both", expand=True)
        
        tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical")
        tree_scroll_x = ttk.Scrollbar(tree_container, orient="horizontal")
        
        self.tree = ttk.Treeview(tree_container, 
                                  columns=("rank", "name", "rate", "count", "pts"), 
                                  show="headings",
                                  yscrollcommand=tree_scroll_y.set,
                                  xscrollcommand=tree_scroll_x.set)
        
        tree_scroll_y.config(command=self.tree.yview)
        tree_scroll_x.config(command=self.tree.xview)
        
        columns = [("rank", "Rank", 50), ("name", "Viewer Name", 180), ("rate", "Attendance", 130), ("count", "Comments", 100), ("pts", "Total Pts", 120)]
        for col, label, width in columns:
            self.tree.heading(col, text=label, command=lambda c=col: self.sort_column(c, c in ["rank", "rate", "count", "pts"]))
            self.tree.column(col, width=width, anchor="center")
        
        self.tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        self.tree.bind("<Double-1>", self.show_user_details)

    def save_api_keys(self):
        liver_name = self.ch_handle_ent.get().strip()
        if not liver_name:
            liver_name = "JemiNeko"
        
        save_settings({
            "api_key": self.api_key_ent.get().strip(),
            "handle": liver_name,
            "gemini_api_key": self.gemini_key_ent.get().strip(),
            "start_date": self.start_date_ent.get(),
            "end_date": self.end_date_ent.get()
        })
        
        self.brain.liver_name = liver_name
        self.brain.init_gemini_client()
        self.log("✅ API keys saved!")
        messagebox.showinfo("Settings Saved", "API keys have been saved!")

    def manual_log_scan(self):
        def run():
            self.log("🔍 Starting manual log scan...")
            handle = self.ch_handle_ent.get().strip()
            total = auto_scan_logs(logger=self.log, channel_handle=handle)
            self.root.after(0, self.refresh_analytics)
            self.log(f"✅ Log scan complete: {total} records added")
            messagebox.showinfo("Complete", f"Log scan complete!\n{total} records updated!")
        threading.Thread(target=run, daemon=True).start()

    def initial_load(self):
        handle = self.settings.get("handle", "").strip()
        auto_scan_logs(logger=self.log, channel_handle=handle)
        self.root.after(0, self.refresh_analytics)

    def refresh_analytics(self):
        query_word = self.search_var.get().strip().lower()
        for i in self.tree.get_children():
            self.tree.delete(i)
        
        with app.app_context():
            v_count = db.session.query(func.count(func.distinct(PointLog.reason))).filter(
                PointLog.reason.notlike('%(LOG)%')
            ).scalar() or 1
            
            query = db.session.query(
                PointLog.username,
                func.count(PointLog.id).label('chat_count'),
                func.count(func.distinct(
                    text("CASE WHEN reason NOT LIKE '%(LOG)%' THEN reason END")
                )).label('video_count'),
                func.sum(PointLog.point_value).label('total_pt')
            ).filter(
                PointLog.username.isnot(None),
                PointLog.username != ''
            ).group_by(PointLog.username)
            
            handle = self.ch_handle_ent.get().strip().lower()
            if handle:
                norm_handle = handle if handle.startswith('@') else f"@{handle}"
                query = query.filter(PointLog.username != norm_handle)
                query = query.filter(PointLog.username != handle)
            
            if query_word:
                query = query.filter(PointLog.username.ilike(f"%{query_word}%"))
            
            results = query.all()
            
            scored = []
            for r in results:
                uname = r[0]
                if not uname:
                    continue
                if not uname.startswith('@'):
                    uname = '@' + uname
                chat_count = r[1] or 0
                video_count = r[2] or 0
                total_pt = r[3] or 0
                attendance_rate = video_count / v_count if v_count > 0 else 0
                
                if attendance_rate == 0:
                    continue
                
                score = (attendance_rate * 0.6) + (min(chat_count / 1000, 1.0) * 0.4)
                scored.append({
                    "name": uname,
                    "chat_count": chat_count,
                    "video_count": video_count,
                    "total_pt": total_pt,
                    "attendance_rate": attendance_rate,
                    "score": score
                })
            
            scored.sort(key=lambda x: (x["score"], x["total_pt"]), reverse=True)
            
            for i, d in enumerate(scored[:500]):
                rate_str = f"{d['attendance_rate']*100:.1f}% ({d['video_count']}/{v_count})"
                self.tree.insert("", "end", values=(
                    i+1,
                    d["name"],
                    rate_str,
                    f"{d['chat_count']}",
                    f"{d['total_pt']:,}"
                ))
        
        self.lbl_usage.config(text=f"💰 API Usage: {self.api_usage} pts")

    def sort_column(self, col, is_num):
        data = [(self.tree.set(c, col), c) for c in self.tree.get_children('')]
        self.sort_reverse = not self.sort_reverse if self.sort_column_name == col else True
        self.sort_column_name = col
        if is_num:
            data.sort(key=lambda x: float(re.sub(r'[^\d.]', '', str(x[0]))) if re.sub(r'[^\d.]', '', str(x[0])) else 0, reverse=self.sort_reverse)
        else:
            data.sort(key=lambda x: str(x[0]).lower(), reverse=self.sort_reverse)
        for i, item in enumerate(data):
            self.tree.move(item[1], '', i)

    def show_user_details(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        u_name = self.tree.item(sel[0])['values'][1]
        
        win = tk.Toplevel(self.root)
        win.title(f"📊 {u_name}'s Details")
        win.geometry("1000x700")
        win.configure(bg="#f0f2f5")
        
        tk.Label(win, text=f"🐾 {u_name}'s Activity History", font=("Segoe UI", 12, "bold"), bg="#f0f2f5", fg="#333").pack(anchor="w", padx=10, pady=(10, 5))
        
        main_frame = tk.Frame(win, bg="#f0f2f5")
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)
        
        tree_frame = tk.LabelFrame(main_frame, text="📋 Stream/Log List", bg="#f0f2f5", fg="#555", font=("", 9))
        tree_frame.pack(fill="both", expand=True, pady=(0, 5))
        
        tree_container = tk.Frame(tree_frame, bg="white")
        tree_container.pack(fill="both", expand=True, padx=3, pady=3)
        
        v_tree_scroll_y = ttk.Scrollbar(tree_container, orient="vertical")
        v_tree_scroll_x = ttk.Scrollbar(tree_container, orient="horizontal")
        
        v_tree = ttk.Treeview(tree_container, columns=("type", "reason", "chat_count", "point_pt"), show="headings",
                               yscrollcommand=v_tree_scroll_y.set,
                               xscrollcommand=v_tree_scroll_x.set)
        v_tree.heading("type", text="Type")
        v_tree.column("type", width=80, anchor="center")
        v_tree.heading("reason", text="Title")
        v_tree.column("reason", width=470, anchor="w")
        v_tree.heading("chat_count", text="Comments")
        v_tree.column("chat_count", width=90, anchor="center")
        v_tree.heading("point_pt", text="Points")
        v_tree.column("point_pt", width=100, anchor="center")
        
        v_tree.grid(row=0, column=0, sticky="nsew")
        v_tree_scroll_y.grid(row=0, column=1, sticky="ns")
        v_tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)
        
        info_frame = tk.Frame(main_frame, bg="#f0f2f5")
        info_frame.pack(fill="x", pady=(0, 5))
        info = tk.Label(info_frame, text="", bg="#f0f2f5", fg="#d32f2f", font=("", 10, "bold"))
        info.pack(side="left")
        
        detail_frame = tk.LabelFrame(main_frame, text="📝 All Chats", bg="#f0f2f5", fg="#555", font=("", 9))
        detail_frame.pack(fill="both", expand=True)
        
        detail_text = scrolledtext.ScrolledText(detail_frame, font=("Consolas", 9), bg="#1e1e1e", fg="#e6edf3", height=10)
        detail_text.pack(fill="both", expand=True, padx=3, pady=3)
        
        detail_text.tag_config("time", foreground="#8b949e")
        detail_text.tag_config("point", foreground="#f472b6")
        detail_text.tag_config("message", foreground="#e6edf3")
        
        with app.app_context():
            res = db.session.query(
                PointLog.reason,
                func.sum(case((PointLog.event_type == "chat", 1), else_=0)).label('chat_cnt'),
                func.sum(PointLog.point_value).label('point_pt')
            ).filter(
                PointLog.username == u_name
            ).group_by(PointLog.reason).order_by(desc(PointLog.date)).all()
            
            total_point = sum(r[2] or 0 for r in res)
            info.config(text=f"💎 Total Points: {total_point:,} pts")
            
            for r in res:
                is_log = "(LOG)" in r[0] if r[0] else False
                type_label = "📁 Log" if is_log else "🎬 Stream"
                
                v_tree.insert("", "end", values=(
                    type_label,
                    r[0][:60] if r[0] else "Unknown",
                    f"{r[1] or 0}",
                    f"{r[2] or 0:,}"
                ), tags=(u_name, r[0]))
        
        def show_chat_detail(event):
            selected = v_tree.selection()
            if not selected:
                return
            item = selected[0]
            tags = v_tree.item(item, "tags")
            values = v_tree.item(item, "values")
            if len(tags) < 2:
                return
            target_user = tags[0]
            target_reason = tags[1]
            target_type = values[0] if values else "Unknown"
            
            detail_win = tk.Toplevel(win)
            detail_win.title(f"📝 {target_user}'s Chats - {target_reason[:50] if target_reason else 'Unknown'}")
            detail_win.geometry("900x600")
            detail_win.configure(bg="#1e1e1e")
            
            header_frame = tk.Frame(detail_win, bg="#1e1e1e", pady=10)
            header_frame.pack(fill="x", padx=10)
            tk.Label(header_frame, text=f"🐾 {target_user}'s Chat Details", 
                     font=("Segoe UI", 14, "bold"), fg="#f472b6", bg="#1e1e1e").pack(anchor="w")
            tk.Label(header_frame, text=f"📺 {target_type}: {target_reason[:80] if target_reason else 'Unknown'}", 
                     font=("Segoe UI", 9), fg="#aaaaaa", bg="#1e1e1e").pack(anchor="w", pady=(5, 0))
            
            text_frame = tk.Frame(detail_win, bg="#1e1e1e")
            text_frame.pack(fill="both", expand=True, padx=10, pady=10)
            
            scrollbar = tk.Scrollbar(text_frame)
            scrollbar.pack(side="right", fill="y")
            
            txt = tk.Text(text_frame, font=("Consolas", 10), bg="#0d1117", fg="#e6edf3",
                         yscrollcommand=scrollbar.set, wrap="word", padx=10, pady=10)
            txt.pack(side="left", fill="both", expand=True)
            scrollbar.config(command=txt.yview)
            
            txt.tag_config("time", foreground="#8b949e")
            txt.tag_config("point", foreground="#f472b6")
            txt.tag_config("message", foreground="#e6edf3")
            
            with app.app_context():
                logs = PointLog.query.filter_by(
                    username=target_user,
                    reason=target_reason
                ).order_by(PointLog.date).all()
                
                for log in logs:
                    time_str = log.date.strftime('%H:%M:%S') if log.date else "??:??:??"
                    point_str = f" +{log.point_value}pts" if log.point_value > 0 else ""
                    
                    txt.insert("end", f"[{time_str}] ", "time")
                    txt.insert("end", f"{log.message}", "message")
                    if point_str:
                        txt.insert("end", f" {point_str}", "point")
                    txt.insert("end", "\n")
            
            txt.configure(state="disabled")
            
            footer = tk.Frame(detail_win, bg="#1e1e1e", pady=10)
            footer.pack(fill="x")
            tk.Label(footer, text=f"💬 Total {len(logs)} chats", 
                     font=("Segoe UI", 9), fg="#888888", bg="#1e1e1e").pack()
            
            btn_frame = tk.Frame(detail_win, bg="#1e1e1e", pady=10)
            btn_frame.pack()
            tk.Button(btn_frame, text="Close", command=detail_win.destroy,
                     bg="#f472b6", fg="#0f172a", font=("", 10, "bold"),
                     relief="flat", padx=20, pady=5).pack()
        
        v_tree.bind("<Double-1>", show_chat_detail)
        
        btn_frame = tk.Frame(win, bg="#f0f2f5", pady=10)
        btn_frame.pack()
        tk.Button(btn_frame, text="Close", command=win.destroy, bg="#607d8b", fg="white", font=("", 10), padx=20, pady=5).pack()

    def consult_strategist(self):
        def _run():
            with app.app_context():
                handle = self.ch_handle_ent.get().strip().lower()
                norm_h = (handle if handle.startswith('@') else f"@{handle}") if handle else None
                
                try:
                    s_date = datetime.strptime(self.start_date_ent.get(), '%Y-%m-%d')
                    e_date = datetime.strptime(self.end_date_ent.get(), '%Y-%m-%d').replace(hour=23, minute=59, second=59)
                except ValueError:
                    self.log("⚠️ Invalid date format. Use YYYY-MM-DD.")
                    return
                
                total_chats = db.session.query(func.count(PointLog.id)).filter(
                    PointLog.date >= s_date,
                    PointLog.date <= e_date
                ).scalar() or 0
                
                top_stream = db.session.query(
                    PointLog.reason, 
                    func.count(PointLog.id)
                ).filter(
                    PointLog.reason.notlike('%(LOG)%'),
                    PointLog.date >= s_date,
                    PointLog.date <= e_date
                ).group_by(PointLog.reason).order_by(desc(func.count(PointLog.id))).first()
                
                fan_q = db.session.query(
                    PointLog.username, 
                    func.sum(PointLog.point_value)
                ).filter(
                    PointLog.date >= s_date,
                    PointLog.date <= e_date
                ).group_by(PointLog.username)
                
                if norm_h:
                    fan_q = fan_q.filter(
                        PointLog.username != norm_h,
                        PointLog.username != handle
                    )
                
                top_fan = fan_q.order_by(desc(func.sum(PointLog.point_value))).first()
                recent_logs = PointLog.query.filter(
                    PointLog.reason.notlike('%(LOG)%'),
                    PointLog.date >= s_date,
                    PointLog.date <= e_date
                ).order_by(desc(PointLog.date)).limit(500).all()
                
                self.log("🤖 Starting detailed analysis with Gemini API...")
                rep, sp = self.brain.generate_gemini_analysis(total_chats, top_stream, top_fan, recent_logs)
                self.root.after(0, lambda: self._show_report(rep))
                self.brain.speak(sp)
                self.log("✅ Analysis complete!")
        threading.Thread(target=_run, daemon=True).start()

    def _show_report(self, text):
        win = tk.Toplevel(self.root)
        win.title("🔮 Strategy Report (Gemini Analysis)")
        win.geometry("800x650")
        t = scrolledtext.ScrolledText(win, font=("Segoe UI", 11))
        t.pack(fill="both", expand=True, padx=10, pady=10)
        t.insert("end", text)
        t.configure(state="disabled")

    def fetch_vids(self):
        api_key = self.api_key_ent.get().strip()
        handle = self.ch_handle_ent.get().strip()
        start_date = self.start_date_ent.get()
        end_date = self.end_date_ent.get()
        
        if not api_key or not handle:
            messagebox.showwarning("Warning", "Please enter API key and handle!")
            return
        
        save_settings({"api_key": api_key, "handle": handle})
        
        def _task():
            try:
                self.log(f"📺 Searching for archives of {handle}...")
                yt = build('youtube', 'v3', developerKey=api_key)
                
                clean_handle = handle.lstrip('@')
                ch_res = yt.channels().list(forHandle=clean_handle, part="id,contentDetails").execute()
                if not ch_res.get('items'):
                    self.log(f"⚠️ forHandle not found. Retrying with search()...")
                    search_res = yt.search().list(q=handle, type='channel', part='id').execute()
                    if not search_res.get('items'):
                        self.log("❌ Channel not found...")
                        return
                    channel_id = search_res['items'][0]['id']['channelId']
                    ch_res2 = yt.channels().list(part="contentDetails", id=channel_id).execute()
                    uploads_id = ch_res2['items'][0]['contentDetails']['relatedPlaylists']['uploads']
                else:
                    channel_id = ch_res['items'][0]['id']
                    uploads_id = ch_res['items'][0]['contentDetails']['relatedPlaylists']['uploads']
                
                videos = []
                next_token = None
                while True:
                    pl_res = yt.playlistItems().list(
                        part="snippet",
                        playlistId=uploads_id,
                        maxResults=50,
                        pageToken=next_token
                    ).execute()
                    self.api_usage += 1
                    
                    for item in pl_res.get('items', []):
                        pub = item['snippet']['publishedAt'][:10]
                        if start_date <= pub <= end_date:
                            v_id = item['snippet']['resourceId']['videoId']
                            videos.append(f"https://www.youtube.com/watch?v={v_id}")
                        elif pub < start_date:
                            break
                    
                    next_token = pl_res.get('nextPageToken')
                    if not next_token:
                        break
                
                self.root.after(0, lambda: (
                    self.url_list_text.delete('1.0', tk.END),
                    self.url_list_text.insert('1.0', "\n".join(videos))
                ))
                self.log(f"✅ Retrieved {len(videos)} archive URLs!")
                
            except Exception as e:
                self.log(f"❌ Fetch error: {str(e)}")
        
        threading.Thread(target=_task, daemon=True).start()

    def start_analysis(self):
        urls = [u.strip() for u in self.url_list_text.get('1.0', tk.END).split('\n') if u.strip()]
        if not urls:
            messagebox.showwarning("Warning", "No URLs to analyze!")
            return
        
        api_key = self.api_key_ent.get().strip()
        if not api_key:
            messagebox.showwarning("Warning", "Please set YouTube API key!")
            return
        
        save_settings({
            "api_key": api_key,
            "handle": self.ch_handle_ent.get().strip(),
            "start_date": self.start_date_ent.get(),
            "end_date": self.end_date_ent.get(),
            "gemini_api_key": self.gemini_key_ent.get().strip()
        })
        
        threading.Thread(target=self._run_analysis, args=(urls,), daemon=True).start()

    def _run_analysis(self, urls):
        try:
            yt = build('youtube', 'v3', developerKey=self.api_key_ent.get().strip())
            jst_tz = timezone(timedelta(hours=9))
            s_date = datetime.strptime(self.start_date_ent.get(), '%Y-%m-%d').replace(tzinfo=jst_tz)
            e_date = datetime.strptime(self.end_date_ent.get(), '%Y-%m-%d').replace(hour=23, minute=59, second=59, tzinfo=jst_tz)
            
            total_videos = len(urls)
            for idx, u in enumerate(urls):
                self.root.after(0, lambda i=idx+1, t=total_videos: self._update_progress(i, t))
                
                vid_match = re.search(r"(v=|\/)([a-zA-Z0-9_-]{11})", u)
                if not vid_match:
                    continue
                vid = vid_match.group(2)
                
                v_res = yt.videos().list(part="liveStreamingDetails,snippet", id=vid).execute()
                self.api_usage += 1
                if not v_res.get('items'):
                    continue
                
                v_item = v_res['items'][0]
                label = v_item['snippet']['title']
                
                live_details = v_item.get('liveStreamingDetails', {})
                raw_start = (
                    live_details.get('actualStartTime')
                    or live_details.get('scheduledStartTime')
                    or v_item['snippet'].get('publishedAt')
                )
                if not raw_start:
                    self.log(f"⏩ Skip (no date): {label}")
                    continue
                
                start_dt = datetime.strptime(raw_start, '%Y-%m-%dT%H:%M:%SZ').replace(tzinfo=timezone.utc).astimezone(jst_tz)
                if start_dt > e_date or start_dt < (s_date - timedelta(days=1)):
                    self.log(f"⏩ Skip (out of range): {label}")
                    continue
                
                self.log(f"🚀 Analyzing: {label}")
                chat = pytchat.create(video_id=vid, interruptable=False)
                bulk = []
                processed_messages = set()
                
                while chat.is_alive():
                    for c in chat.get().items:
                        exact_date = start_dt + timedelta(seconds=self._parse_elapsed_time(c.elapsedTime))
                        if s_date <= exact_date <= e_date:
                            username = clean_name(c.author.name)
                            if username and username not in ["@Sakuneko", "@System", "@SakunekoBot"]:
                                msg_key = f"{username}_{c.message[:50]}"
                                if msg_key in processed_messages:
                                    continue
                                processed_messages.add(msg_key)
                                
                                bulk.append(PointLog(
                                    username=username,
                                    message=c.message,
                                    reason=label,
                                    video_id=vid,
                                    point_value=0,
                                    event_type="chat",
                                    date=exact_date.replace(tzinfo=None)
                                ))
                    time.sleep(0.01)
                
                with app.app_context():
                    if bulk:
                        db.session.bulk_save_objects(bulk)
                        db.session.commit()
                        self.log(f"✅ {label}: Recorded {len(bulk)} chats")
            
            self.root.after(0, lambda: self._on_analysis_complete())
            
        except Exception as e:
            error_msg = f"❌ Error:\n{traceback.format_exc()}"
            self.root.after(0, lambda: self.log(error_msg))
            self.root.after(0, lambda: messagebox.showerror("Analysis Error", f"An error occurred during analysis...\n{str(e)}"))

    def _parse_elapsed_time(self, elapsed_str):
        try:
            is_minus = elapsed_str.startswith("-")
            clean_str = elapsed_str.lstrip("-")
            parts = list(map(int, clean_str.split(':')))
            if len(parts) == 2:
                seconds = parts[0]*60 + parts[1]
            elif len(parts) == 3:
                seconds = parts[0]*3600 + parts[1]*60 + parts[2]
            else:
                seconds = 0
            return -seconds if is_minus else seconds
        except Exception:
            return 0

    def _update_progress(self, current, total):
        self.root.title(f"🐾 Analyzing... {current}/{total}")

    def _on_analysis_complete(self):
        messagebox.showinfo("🎉 Analysis Complete", "Data analysis complete!\nCheck the rankings!")
        self.log("=" * 50)
        self.log("🎉 Analysis complete!")
        self.log("=" * 50)
        self.refresh_analytics()
        handle = self.ch_handle_ent.get().strip()
        auto_scan_logs(logger=self.log, channel_handle=handle)
        self.refresh_analytics()

    def export_rank_csv(self):
        fp = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not fp:
            return
        try:
            with open(fp, 'w', newline='', encoding='utf-8-sig') as f:
                w = csv.writer(f)
                w.writerow(["Rank", "Viewer Name", "Attendance", "Comments", "Total Points"])
                for c in self.tree.get_children():
                    w.writerow(self.tree.item(c)['values'])
            self.log("✅ Rankings saved to CSV")
        except Exception as e:
            self.log(f"❌ Save error: {e}")
            messagebox.showerror("Error", f"Failed to save CSV: {e}")

    def reset_database(self):
        if messagebox.askyesno("⚠️ Warning", "Delete all analysis data?"):
            if messagebox.askyesno("❗ Final Confirmation", "Are you absolutely sure?"):
                try:
                    with app.app_context():
                        db.drop_all()
                        db.create_all()
                    self.log("🧨 Database reset.")
                    self.refresh_analytics()
                except Exception as e:
                    messagebox.showerror("Error", str(e))

    def log(self, m):
        self.root.after(0, lambda: (
            self.log_area.insert("end", f"[{datetime.now().strftime('%H:%M:%S')}] {m}\n"),
            self.log_area.see("end")
        ))

if __name__ == '__main__':
    # 🐾 Trial check (5 launches / 14 days limit)
    if not check_trial():
        sys.exit(0)
    root = tk.Tk()
    app_gui = ArchiveImporterGUI(root)
    root.mainloop()
