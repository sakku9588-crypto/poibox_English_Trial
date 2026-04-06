import tkinter as tk
from tkinter import messagebox, ttk
import threading
import re
import time
import json
from datetime import datetime, timedelta
from googleapiclient.discovery import build
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import sys
import logging

# ==========================================
# Logging Setup
# ==========================================
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# ==========================================
# Path Resolution Utility
# ==========================================
def get_resource_path(relative_path):
    """Resolve path correctly even after PyInstaller conversion"""
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# ==========================================
# Constants
# ==========================================
SETTINGS_FILE       = get_resource_path("settings_delta.json")
PROCESSED_IDS_FILE  = get_resource_path("processed_comment_ids.json")
LOG_DIR             = get_resource_path("log")
JSON_KEY_PATH       = get_resource_path("credentials.json")

# Create log folder if it doesn't exist
os.makedirs(LOG_DIR, exist_ok=True)

# ==========================================
# 🐾 Trial Management (5 launches / 14 days limit)
# ==========================================
TRIAL_FILE      = get_resource_path("shorts_trial.json")
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
    import tkinter as _tk
    from tkinter import messagebox as _mb
    _root = _tk.Tk()
    _root.withdraw()
    _mb.showwarning(
        "⏰ Trial Limit Reached",
        f"【PoiBox Video & Shorts - Trial】\n\n{msg}\n\nThank you for trying! 🐾"
    )
    _root.destroy()


# ==========================================
# Main Application Class
# ==========================================
class UltimateBatchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🐾 PoiBox Video & Shorts v12.0")
        self.root.geometry("600x920")
        self.root.configure(bg="#1a1a1a")

        self.comment_data   = []
        self.processed_ids  = self._load_processed_ids()
        self.video_titles   = {}
        self._scan_running  = False  # 🐾 Prevent duplicate scan
        self._sync_running  = False  # 🐾 Prevent duplicate sync

        self._setup_ui()

    # ------------------------------------------
    # UI Setup
    # ------------------------------------------
    def _setup_ui(self):
        main_frame = tk.Frame(self.root, bg="#1a1a1a", padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            main_frame, text="📦 PoiBox Video & Shorts",
            font=("Meiryo", 18, "bold"), fg="#ffcc00", bg="#1a1a1a"
        ).pack(pady=(0, 10))

        # 🛠️ Basic Settings
        config_frame = tk.LabelFrame(
            main_frame, text="🛠️ Basic Settings",
            fg="white", bg="#2c2c2c", font=("", 10, "bold"), padx=10, pady=10
        )
        config_frame.pack(fill=tk.X, pady=5)

        self.api_entry    = self._add_labeled_entry(config_frame, "YouTube API Key:", "")
        self.handle_entry = self._add_labeled_entry(config_frame, "Target Handle (without @):", "")
        self.sheet_entry  = self._add_labeled_entry(config_frame, "Spreadsheet URL:", "")

        saved = self._load_settings()
        self.api_entry.insert(0,    saved.get("api_key",   ""))
        self.handle_entry.insert(0, saved.get("handle",    ""))
        self.sheet_entry.insert(0,  saved.get("sheet_url", ""))

        # 🕒 Date Filter Settings
        filter_frame = tk.LabelFrame(
            main_frame, text="🕒 Date Range Filter",
            fg="white", bg="#2c2c2c", font=("", 10, "bold"), padx=10, pady=10
        )
        filter_frame.pack(fill=tk.X, pady=5)

        date_input_frame = tk.Frame(filter_frame, bg="#2c2c2c")
        date_input_frame.pack(fill=tk.X)

        now       = datetime.now()
        week_ago  = now - timedelta(days=7)

        tk.Label(date_input_frame, text="Start:", fg="white", bg="#2c2c2c").pack(side=tk.LEFT)
        self.start_date_entry = tk.Entry(date_input_frame, width=12, justify="center")
        self.start_date_entry.insert(0, week_ago.strftime("%Y-%m-%d"))
        self.start_date_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(date_input_frame, text="End:", fg="white", bg="#2c2c2c").pack(side=tk.LEFT, padx=(10, 0))
        self.end_date_entry = tk.Entry(date_input_frame, width=12, justify="center")
        self.end_date_entry.insert(0, now.strftime("%Y-%m-%d"))
        self.end_date_entry.pack(side=tk.LEFT, padx=5)

        tk.Label(
            filter_frame, text="※ Format: YYYY-MM-DD",
            fg="#aaaaaa", bg="#2c2c2c", font=("", 8)
        ).pack(anchor="w", pady=(5, 0))

        # 🚀 Buttons
        btn_frame = tk.Frame(main_frame, bg="#1a1a1a")
        btn_frame.pack(fill=tk.X, pady=15)

        self.btn_scan = tk.Button(
            btn_frame, text="🔍 Start Comment Analysis",
            command=self._start_scan_thread,
            bg="#007bff", fg="white", font=("", 11, "bold"), height=2
        )
        self.btn_scan.pack(fill=tk.X, pady=2)

        self.btn_sync = tk.Button(
            btn_frame, text="💎 Batch Add Points to Spreadsheet",
            command=self._start_sync_thread,
            bg="#28a745", fg="white", font=("", 11, "bold"), height=2,
            state=tk.DISABLED
        )
        self.btn_sync.pack(fill=tk.X, pady=2)

        self.status_var = tk.StringVar(value="Waiting... 🐾")
        tk.Label(
            main_frame, textvariable=self.status_var,
            fg="#00ffcc", bg="#1a1a1a", font=("", 10, "bold")
        ).pack(pady=5)

        # Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#2c2c2c", foreground="white",
            fieldbackground="#2c2c2c", rowheight=25
        )
        self.tree = ttk.Treeview(main_frame, columns=("ID", "Name", "Content"), show="headings")
        self.tree.heading("ID",      text="Video ID");    self.tree.column("ID",      width=100)
        self.tree.heading("Name",    text="Viewer");      self.tree.column("Name",    width=100)
        self.tree.heading("Content", text="Comment");     self.tree.column("Content", width=250)
        self.tree.pack(fill=tk.BOTH, expand=True)

    def _add_labeled_entry(self, parent, label, default):
        f = tk.Frame(parent, bg="#2c2c2c")
        f.pack(fill=tk.X, pady=2)
        tk.Label(f, text=label, fg="white", bg="#2c2c2c", width=25, anchor="w").pack(side=tk.LEFT)
        ent = tk.Entry(f, bg="#1a1a1a", fg="white", insertbackground="white")
        ent.insert(0, default)
        ent.pack(side=tk.LEFT, fill=tk.X, expand=True)
        return ent

    # ------------------------------------------
    # Settings Load/Save
    # ------------------------------------------
    def _load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception as e:
                logger.warning(f"Failed to load settings: {e}")
        return {}

    def _save_settings(self):
        data = {
            "api_key":   self.api_entry.get().strip(),
            "handle":    self.handle_entry.get().strip(),
            "sheet_url": self.sheet_entry.get().strip()
        }
        try:
            with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logger.warning(f"Failed to save settings: {e}")

    # ------------------------------------------
    # Processed Comment IDs Load/Save
    # ------------------------------------------
    def _load_processed_ids(self):
        if os.path.exists(PROCESSED_IDS_FILE):
            try:
                with open(PROCESSED_IDS_FILE, "r", encoding="utf-8") as f:
                    return set(json.load(f))
            except Exception as e:
                logger.warning(f"Failed to load processed IDs: {e}")
        return set()

    def _save_processed_ids(self):
        try:
            with open(PROCESSED_IDS_FILE, "w", encoding="utf-8") as f:
                json.dump(list(self.processed_ids), f)
        except Exception as e:
            logger.warning(f"Failed to save processed IDs: {e}")

    # ------------------------------------------
    # Scan Logic
    # ------------------------------------------
    def _start_scan_thread(self):
        if self._scan_running:
            messagebox.showwarning("Running", "Scan is already in progress!")
            return
        threading.Thread(target=self._scan_logic, daemon=True).start()

    def _scan_logic(self):
        self._scan_running = True
        self._save_settings()

        api_key   = self.api_entry.get().strip()
        handle    = self.handle_entry.get().strip().replace('@', '')
        start_str = self.start_date_entry.get().strip()
        end_str   = self.end_date_entry.get().strip()

        if not api_key or not handle:
            messagebox.showerror("Error", "Please enter both API Key and Handle!")
            self._scan_running = False
            return

        self.root.after(0, lambda: self.btn_scan.config(state=tk.DISABLED))
        self.comment_data = []
        self.root.after(0, lambda: [self.tree.delete(i) for i in self.tree.get_children()])

        try:
            # 🐾 Date validation
            try:
                start_dt = datetime.strptime(start_str, "%Y-%m-%d")
                end_dt   = datetime.strptime(end_str,   "%Y-%m-%d") + timedelta(days=1)
            except ValueError:
                messagebox.showerror("Date Error", "Invalid date format!\nPlease use YYYY-MM-DD.")
                return

            if start_dt >= end_dt:
                messagebox.showerror("Date Error", "Start date must be before end date!")
                return

            youtube = build("youtube", "v3", developerKey=api_key)

            # Get channel ID
            self.root.after(0, lambda: self.status_var.set("🔍 Looking up channel information..."))
            ch_res = youtube.channels().list(forHandle=handle, part="id").execute()
            items  = ch_res.get("items")
            if not items:
                raise ValueError(f"Channel not found: @{handle}")
            channel_id = items[0]["id"]

            published_after  = start_dt.isoformat() + "Z"
            published_before = end_dt.isoformat()   + "Z"

            # Get video list
            self.root.after(0, lambda: self.status_var.set("📺 Searching for videos in date range..."))
            videos    = []
            next_page = None
            while True:
                v_res = youtube.search().list(
                    channelId=channel_id, part="id,snippet", order="date",
                    maxResults=50, publishedAfter=published_after,
                    publishedBefore=published_before, pageToken=next_page
                ).execute()
                for item in v_res.get("items", []):
                    vid = item["id"].get("videoId")
                    if vid:
                        videos.append((vid, item["snippet"]["title"]))
                next_page = v_res.get("nextPageToken")
                if not next_page:
                    break

            if not videos:
                self.root.after(0, lambda: self.status_var.set("⚠️ No videos found in the specified date range"))
                messagebox.showinfo("Result", "No videos found in the specified date range.\nTry expanding the date range!")
                return

            # Get comments
            total_videos = len(videos)
            self.root.after(0, lambda: self.status_var.set(f"💬 Analyzing {total_videos} videos..."))

            for idx, (v_id, title) in enumerate(videos, 1):
                self.video_titles[v_id] = title
                short_title = title[:20] + "..." if len(title) > 20 else title
                self.root.after(0, lambda t=short_title, i=idx: self.status_var.set(
                    f"🎬 [{i}/{total_videos}] Analyzing: {t}"
                ))
                try:
                    c_next_page = None
                    while True:
                        c_res = youtube.commentThreads().list(
                            videoId=v_id, part="snippet",
                            maxResults=100, pageToken=c_next_page
                        ).execute()
                        for item in c_res.get("items", []):
                            top = item["snippet"]["topLevelComment"]
                            snippet = top["snippet"]
                            c_id    = top["id"]
                            author  = snippet["authorDisplayName"]
                            text    = snippet["textDisplay"]
                            if c_id not in self.processed_ids:
                                self.comment_data.append((v_id, author, text, c_id))
                                self.root.after(0, lambda vi=v_id, a=author, t=text: self.tree.insert(
                                    "", tk.END, values=(vi, a, t[:30])
                                ))
                        c_next_page = c_res.get("nextPageToken")
                        if not c_next_page:
                            break
                except Exception as e:
                    # Skip videos with disabled comments / private videos
                    logger.warning(f"Skipping comment fetch [{v_id}]: {e}")
                    continue

            new_count = len(self.comment_data)
            self.root.after(0, lambda: self.status_var.set(f"✅ Analysis complete! New comments: {new_count}"))
            if self.comment_data:
                self.root.after(0, lambda: self.btn_sync.config(state=tk.NORMAL))
            else:
                messagebox.showinfo("Result", "No new comments found.\nAll comments may have been processed already, or there are no comments.")

        except ValueError as e:
            messagebox.showerror("Error", str(e))
        except Exception as e:
            logger.error(f"Scan error: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred during scan:\n{e}")
        finally:
            self.root.after(0, lambda: self.btn_scan.config(state=tk.NORMAL))
            self._scan_running = False

    # ------------------------------------------
    # Spreadsheet Sync Logic
    # ------------------------------------------
    def _start_sync_thread(self):
        if self._sync_running:
            messagebox.showwarning("Running", "Sync is already in progress!")
            return
        threading.Thread(target=self._sync_logic, daemon=True).start()

    def _sync_logic(self):
        self._sync_running = True
        sheet_url = self.sheet_entry.get().strip()

        if not sheet_url:
            messagebox.showerror("Error", "Please enter the Spreadsheet URL!")
            self._sync_running = False
            return
        if not self.comment_data:
            messagebox.showwarning("No Data", "Please run comment analysis first!")
            self._sync_running = False
            return

        # Check if credentials.json exists
        if not os.path.exists(JSON_KEY_PATH):
            messagebox.showerror(
                "Auth Error",
                f"credentials.json not found!\nLocation: {JSON_KEY_PATH}"
            )
            self._sync_running = False
            return

        self.root.after(0, lambda: self.btn_sync.config(state=tk.DISABLED))

        try:
            self.root.after(0, lambda: self.status_var.set("📊 Connecting to spreadsheet..."))

            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEY_PATH, scope)
            gc    = gspread.authorize(creds)

            try:
                sheet = gc.open_by_url(sheet_url).get_worksheet(0)
            except gspread.exceptions.SpreadsheetNotFound:
                raise ValueError("Spreadsheet not found! Please check the URL.")
            except gspread.exceptions.APIError as e:
                raise ValueError(f"Failed to access spreadsheet: {e}")

            self.root.after(0, lambda: self.status_var.set("📊 Calculating analytics & updating..."))

            records    = sheet.get_all_records()
            user_cache = {
                str(r.get('Handle', '')).strip(): {
                    'Points': r.get('Points', 0),
                    'Total':  r.get('Total',  0),
                    'Row':    i + 2
                }
                for i, r in enumerate(records)
                if str(r.get('Handle', '')).strip()
            }

            # Aggregate comment data
            video_analytics = {}
            global_pts_map  = {}
            for v_id, author, content, c_id in self.comment_data:
                u = author.replace('@', '').strip()
                if not u:
                    continue
                if v_id not in video_analytics:
                    video_analytics[v_id] = {}
                video_analytics[v_id][u] = video_analytics[v_id].get(u, 0) + 1
                global_pts_map[u]        = global_pts_map.get(u, 0) + 10
                self.processed_ids.add(c_id)

            now_dt  = datetime.now()
            now_str = now_dt.strftime('%Y-%m-%d %H:%M:%S')

            # Generate report
            report = [f"📊 === PoiBox Analytics Report ===\n📅 Generated: {now_str}\n"]
            for v_id, users in video_analytics.items():
                title = self.video_titles.get(v_id, "Unknown")
                report.append(f"📺 【{title}】\n   💎 Points awarded: {sum(users.values()) * 10} pts")
                for i, (name, count) in enumerate(sorted(users.items(), key=lambda x: x[1], reverse=True)[:5]):
                    medals = ["🥇", "🥈", "🥉"]
                    medal  = medals[i] if i < 3 else " -"
                    report.append(f"      {medal} @{name}: {count * 10} pts")
                report.append("")

            # Write to spreadsheet
            total_u = len(global_pts_map)
            for i, (handle, pts) in enumerate(global_pts_map.items()):
                self.root.after(0, lambda idx=i+1, h=handle: self.status_var.set(
                    f"💎 Updating ({idx}/{total_u}): @{h[:10]}"
                ))
                try:
                    if handle in user_cache:
                        u_info  = user_cache[handle]
                        new_val = [[
                            int(u_info['Points']) + pts,
                            int(u_info['Total'])  + pts,
                            now_str
                        ]]
                        sheet.update(
                            values=new_val,
                            range_name=f'B{u_info["Row"]}:D{u_info["Row"]}'
                        )
                    else:
                        sheet.append_row([handle, pts, pts, now_str])
                    time.sleep(1.2)  # API rate limit protection
                except gspread.exceptions.APIError as e:
                    logger.warning(f"Skipped update for [{handle}]: {e}")
                    continue

            # Save log file
            log_file = os.path.join(LOG_DIR, f"ANALYTICS_{now_dt.strftime('%Y%m%d_%H%M%S')}.txt")
            try:
                with open(log_file, "w", encoding="utf-8") as f:
                    f.write("\n".join(report))
            except Exception as e:
                logger.warning(f"Failed to save log file: {e}")

            self._save_processed_ids()
            self.root.after(0, lambda: self.status_var.set(f"✅ Analytics complete! Check the log file."))
            messagebox.showinfo("Success", f"💎 Successfully updated {total_u} users!")
            self.root.after(0, self._clear_results)

        except ValueError as e:
            self.root.after(0, lambda: self.status_var.set("❌ Error occurred"))
            messagebox.showerror("Error", str(e))
        except Exception as e:
            logger.error(f"Sync error: {e}", exc_info=True)
            self.root.after(0, lambda: self.status_var.set("❌ Error occurred"))
            messagebox.showerror("Sync Error", f"An error occurred during sync:\n{e}")
        finally:
            self.root.after(0, lambda: self.btn_sync.config(state=tk.NORMAL))
            self._sync_running = False

    # ------------------------------------------
    # Clear Results
    # ------------------------------------------
    def _clear_results(self):
        self.comment_data  = []
        self.video_titles  = {}
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.btn_sync.config(state=tk.DISABLED)


# ==========================================
# Entry Point
# ==========================================
if __name__ == "__main__":
    # 🐾 Trial check (5 launches / 14 days limit)
    if not check_trial():
        sys.exit(0)
    root = tk.Tk()
    app  = UltimateBatchApp(root)
    root.mainloop()