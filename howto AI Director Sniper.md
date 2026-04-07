# 🎬 AIdirector V14.0 PRO / 🎯 Sniper V5.0 PRO

AI-powered video editing tools with Gemini 2.0 Flash  
**Windows Only** (macOS/Linux support may work but not tested)

---

## ⚠️ REQUIREMENTS

### FFmpeg & FFprobe are REQUIRED!

These tools need FFmpeg to process videos.

#### 📥 Installation for Windows:

1. Go to [ffmpeg.org](https://ffmpeg.org/download.html)
2. Click "Windows" → "Windows builds by gyan.dev" or "BtbN"
3. Download the latest **release-full-build.7z** or **.zip**
4. Extract to `C:\ffmpeg` (recommended)
5. **Add to PATH:**
   - Right-click "This PC" → Properties → Advanced System Settings
   - Environment Variables → System Variables → Path → Edit
   - Add `C:\ffmpeg\bin`
   - Click OK
6. **Restart Command Prompt / IDE**

#### ✅ Verify Installation:

Open Command Prompt and run:
```cmd
ffmpeg -version
ffprobe -version