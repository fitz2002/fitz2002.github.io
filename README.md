# Work Schedule Calendar Exporter

Easily convert your work schedule Excel files into Google Calendar or Outlook/iCalendar events! This tool helps you upload a `.xlsx` schedule, configure event and reminder settings, and export a ready-to-import `.csv` or `.ics` file for your calendar.

---

## 🚀 Features
- **Excel Upload:** Drag and drop or select your `.xlsx` schedule file.
- **Automatic Parsing:** Extracts all relevant fields and fills in missing locations.
- **Custom Event Times:** Set default times for Prep Trip, Breakdown, and Prep Only activities.
- **Smart Reminders:** Enable reminders with customizable lead time and per-activity selection.
- **Preview:** See your events before exporting.
- **Export:** Download as Google Calendar CSV or Outlook/iCalendar ICS.
- **Area Inheritance:** Automatically fills in missing area information from previous rows.
- **Location Handling:** Processes travel locations and creates descriptive event details.

---

## 🛠️ Getting Started

### 1. **Clone the Repository**
```sh
git clone git@github.com:fitz2002/fitz2002.git
cd fitz2002
```

### 2. **Install Dependencies**
No installation needed! All dependencies are loaded via CDN.

### 3. **Run Locally**
You can use any static file server. For example, with Python:
```sh
python -m http.server 8000
```
Then open [http://localhost:8000](http://localhost:8000) in your browser.

---

## 📋 Usage
1. **Upload your schedule:** Click the file input and select your `.xlsx` file.
2. **Choose export format:** Google Calendar (CSV) or Outlook (ICS).
3. **Configure settings:**
   - **Activity Time Settings:** Set start and end times for Prep Trip, Breakdown, and Prep Only activities.
   - **Reminder Settings:** 
     - Check "Enable reminders" to activate reminder functionality
     - Set reminder time (hours before event)
     - Select which activities should have reminders using the activity checkboxes that appear after file upload
4. **Preview:** Click "Generate Preview" to see your events.
5. **Export:** Click "Export Calendar" to download your file.
6. **Import:**
   - For Google Calendar: Import the CSV file.
   - For Outlook/iCalendar: Import the ICS file.

---

## 📑 Excel File Format
Your Excel file should have columns like:
- `From`, `To`, `Version`, `Bkd`, `Activity` (or `Private Activity`), `Unit`, `From Loc`, `To Loc`, `Area`, `Notes`, `Departure Attributes`

The app is robust to missing or extra columns, and will fill in missing locations as needed.

### Required Columns:
- **From:** Start date (mm/dd/yy format or Excel serial date)
- **To:** End date (mm/dd/yy format or Excel serial date) 
- **Version:** Version information
- **Bkd:** Booking information
- **Activity:** Activity type (supports both "Activity" and "Private Activity" column headers)

### Optional Columns:
- **Unit:** Unit information
- **From Loc:** Starting location
- **To Loc:** Destination location
- **Area:** Geographic area (automatically inherited if missing)
- **Notes:** Additional notes
- **Departure Attributes:** Departure-related information

---

## 🔔 Reminder System
The reminder system allows you to:
- **Enable/Disable:** Toggle reminders on or off globally
- **Set Lead Time:** Configure how many hours before an event the reminder should trigger
- **Per-Activity Control:** After uploading a file, checkboxes appear for each unique activity type, allowing you to selectively enable reminders for specific activities
- **Automatic Integration:** Reminders are automatically included in both Google Calendar (CSV) and Outlook (ICS) exports

---

## 🕐 Time Configuration
Configure default times for different activity types:
- **Prep Trip:** Default 8:00 AM - 5:00 PM
- **Breakdown:** Default 8:00 AM - 5:00 PM  
- **Prep Only:** Default 9:00 AM - 3:00 PM

These times are applied to events based on the activity type in your Excel file.

---

## 🤝 Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

---

## 📄 License
MIT 