# 📊 How to Get Google Sheets API Credentials

## Step 1: Go to Google Cloud Console
👉 [https://console.cloud.google.com/](https://console.cloud.google.com/)

## Step 2: Select Your Project
- Use the same project as YouTube API
- Or create a new one

## Step 3: Enable Google Sheets API
1. Go to "APIs & Services" → "Library"
2. Search for "Google Sheets API"
3. Click "Enable"

## Step 4: Create Service Account (for desktop apps)
1. Go to "APIs & Services" → "Credentials"
2. Click "+ Create Credentials" → "Service Account"
3. Name it `poibox-sheets`
4. Click "Create and Continue"
5. Assign role: "Editor" (or "Viewer")
6. Click "Done"

## Step 5: Download Credentials
1. Click the service account email
2. Go to "Keys" tab
3. Click "Add Key" → "Create New Key"
4. Select "JSON"
5. Download the file → rename to `sheets_credentials.json`

## Step 6: Share Your Spreadsheet
1. Create a Google Spreadsheet
2. Click "Share" (top right)
3. Add the service account email as Editor
   (Find it in the downloaded JSON: `client_email`)
4. Click "Send"

## Step 7: Enter in PoiBox
1. Place `sheets_credentials.json` in the same folder as PoiBox
2. Enter your spreadsheet ID in settings
   (Get from URL: `https://docs.google.com/spreadsheets/d/【THIS_IS_YOUR_ID】/edit`)

---

⚠️ **Important!**
- Never share `sheets_credentials.json`
- The service account email needs access to your sheet
- Keep the JSON file safe! 🐾