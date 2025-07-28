# C.L.A.R.A. – Company Legal & Accounting Reporting Assistant  
**by Hession Dynamics**

This is a simple but powerful desktop tool (built with Python + Tkinter) that helps you fill out and generate a company filing report in `.docx` format. It covers CS01 confirmation statements, expenses, annual accounts, and CT600 tax return basics – with the option to upload a logo and attach receipt images.

---

## 💡 What It Does

C.L.A.R.A. lets you:

- Fill in company details, CS01 confirmation data, accounting numbers, and CT600 tax return fields
- Add individual expenses, attach receipt images, and automatically calculate total expenses
- Import a metrics CSV and load values straight into the accounts tab
- Save all your inputs into a clean, professional `.docx` file
- Store your last used data in a `.json` file for easy reuse
- Add a company logo to the top of the generated document (optional)

---

## 🔧 How to Run It

Make sure you have Python 3.6+ and the required dependencies.

### Install dependencies:
```bash
pip install python-docx
```

### Run the app:
```bash
python your_script_name.py
```

(Replace `your_script_name.py` with whatever you named the file.)

---

## 📂 Features Breakdown

### Tabs
- **Company Info** – Basic stuff like name, number, and email.
- **CS01 Confirmation Statement** – Confirmation date, SIC code, shareholders, etc.
- **Expenses** – Add as many as you want. You can attach receipt images here too.
- **Annual Accounts** – Turnover, assets, liabilities, and auto-calculated profit/loss.
- **Corporation Tax Return (CT600)** – Tax year start/end and profit/tax numbers.

### Buttons
- **Submit & Download DOCX** – Pulls all your data into a Word file.
- **Import Metrics CSV** – Pulls values (like turnover) from a `.csv` file and auto-fills them.
- **Load Previous Data** – Loads data from your last saved session (`last_data.json`).
- **Choose Logo** – Pick a company logo to include in your final report.

---

## 🧠 How It Works (Briefly)

- Internally, all fields are stored in dictionaries for easy access.
- When you click “Submit & Download DOCX,” the app compiles your inputs, processes expenses, inserts any uploaded images, and builds a document using `python-docx`.
- The last submitted data is saved as `last_data.json` so you can reload it later.
- Receipt images and logos are scaled to fit inside the doc cleanly (2-inch width).

---

## 🛠 File Structure

- `last_data.json` – stores your previous session
- Output is saved as a `.docx` (user chooses name & location on submit)

---

## ⚠️ Notes

- Images must be `.jpg`, `.jpeg`, or `.png`.
- Expense amounts must be valid numbers – if not, the app will show an error popup.
- All fields are optional, but missing info will just leave that section blank in the report.

---

## ✅ To-Do (Optional Improvements)

- Add PDF export
- Add built-in receipt viewer
- Better validation and formatting for numbers and dates

---

## 📣 Author

Built by Hession Dynamics.  
Any feedback or feature requests, just shout.