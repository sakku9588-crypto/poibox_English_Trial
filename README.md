📦 Poibox Ecosystem (v15.0 PRO)
Developed by Sakku at Sakuramori Lab, Poibox has evolved from a simple reward automation tool into a full-scale AI-driven content and community management ecosystem. This suite leverages the latest Google Gemini AI, YouTube APIs, and OBS integration to automate the "boring" parts of being a creator.
🐱 The Sakuneko Philosophy
"Maximum rewards, minimum effort."
Whether you are farming points, managing a 24/7 live stream, or hunting for viral clips, Poibox is your digital command center.
🚀 Core Components
1. 🎯 Poibox Sniper (Sniper V5.0 PRO)
File: sniper_en_t 2.py
• What it does: Uses Gemini 2.0 Flash to scan hours of video content at 5x speed to find "viral" or "exciting" moments automatically.
• Key Feature: Outputs a JSON structure of the best scenes with high scores, ready for social media export.
2. 🎬 AI Director (AIdirector V14.0 PRO)
File: AIdirector_en_t.py
• What it does: Acts as a world-class video editor. It doesn't just find clips; it structures them following viral "Hook-Build-Peak" storytelling patterns.
• Key Feature: Automated FFmpeg-based cutting and concatenation of highlight reels.
3. 📡 Poibox Live (Live v5.2)
File: live_en.py
• What it does: Real-time monitoring of YouTube Live chats. It integrates with OBS WebSocket v5 to trigger scenes/alerts based on chat activity.
• Key Feature: Seamlessly syncs live viewer data to Google Sheets for immediate analytics.
4. 📊 Community Analysis Engine
Files: analysis.py, shorts_en_t.py
• What it does: Aggregates viewer data, calculates attendance, and assigns "Loyalty Points" to your community members.
• Key Feature: Database-backed rankings that help identify your most active fans and potential trolls (Smurf/Troll detection integration).
🏗️ Technical Architecture
• AI Engine: Google GenAI (Gemini 2.0 Flash / 1.5 Pro)
• GUI: Modern, dark-mode interfaces powered by customtkinter.
• Database: SQLAlchemy / SQLite for robust local data tracking.
• Integration: OBS WebSocket v5, Google Sheets API, YouTube Data API v3.
• Security: Fernet-based encryption for API keys and service credentials.
🛠️ Setup & Trial
Most modules in the Poibox suite come with a built-in trial system (HWID-bound) for secure distribution.
1.	Install requirements: pip install customtkinter google-genai pytchat gspread oauth2client yt-dlp
2.	Ensure ffmpeg is installed in your system PATH.
3.	Place your credentials.json (Google) and gemini_api_key.txt in the root folder.
4.	Run main.py or individual components as needed.
🐈 Developer's Note
"Poibox is the culmination of years of 'Poikatsu' and community management experience. We don't just build tools; we build digital assistants that think like humans but work like machines."
👉 Official Website: Sakuramori Lab
Empowering the next generation of creators.
