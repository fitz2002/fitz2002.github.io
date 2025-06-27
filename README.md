# Work Schedule Calendar Exporter

Easily convert your work schedule Excel files into Google Calendar or Outlook/iCalendar events! This tool helps you upload a `.xlsx` schedule, configure event and reminder settings, and export a ready-to-import `.csv` or `.ics` file for your calendar.

---

## 🚀 Features
- **Excel Upload:** Drag and drop or select your `.xlsx` schedule file.
- **Automatic Parsing:** Extracts all relevant fields and fills in missing locations.
- **Timezone Support:** Choose your timezone for accurate event and reminder times.
- **Custom Event Times:** Set default times for Prep Trip, Breakdown, and more.
- **Reminders:** Add reminders before events, with customizable lead time.
- **Preview:** See your events before exporting.
- **Export:** Download as Google Calendar CSV or Outlook/iCalendar ICS.

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
   - Set event times for each activity type.
   - Enable reminders and set how many hours in advance.
   - Select your timezone.
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

---

## 🤝 Contributing
Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

---

## 📄 License
MIT 