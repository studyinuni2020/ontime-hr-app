# OnTime HR — Android + Google Apps Script

## 1. Google Sheet
Create a Google Sheet named `OnTime HR`. Copy its Spreadsheet ID.

## 2. Apps Script
Extensions → Apps Script.
Add `Code.gs` and `Index.html` from this package.
In `Code.gs`, replace:
`PASTE_YOUR_GOOGLE_SHEET_ID_HERE`
with your spreadsheet ID.

Run `setup()` once and authorize.

## 3. Deploy
Deploy → New deployment → Web app.
Execute as: Me.
Access: Anyone with the link (or your organization's policy).
Copy the Web App URL.

## 4. Android Studio
Open this folder as an Android Studio project.
Open `app/build.gradle`.
Replace:
`PASTE_YOUR_APPS_SCRIPT_WEB_APP_URL_HERE`
with the Web App URL.

Sync Gradle, then Build → Generate App Bundle / APK → Generate APK.

The generated debug APK is normally:
`app/build/outputs/apk/debug/app-debug.apk`

For a shareable release APK use:
Build → Generate Signed Bundle / APK → APK.

## 5. First login
User ID: `ADMIN001`
Password: `admin123`

Change the initial admin password before production use.

## 6. Architecture
Android APK (WebView) → Apps Script Web App → Google Sheet.

The APK itself contains no spreadsheet data. Employees use individual accounts.
