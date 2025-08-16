# 📘 GCR Marks Extractor

This project automates the process of **extracting student marks from Google Classroom** and **mapping them to an Excel sheet**.  
It is designed to save teachers/students time by directly fetching marks from Google Classroom, parsing roll numbers, and updating them into class Excel files.

---

## ✨ Features
-  Authenticate with Google Classroom API  
-  Extract courses, assignments, and student submissions  
-  Parse roll numbers from student names or attachments  
-  Save marks to a JSON file (`gcr_marks.txt`)  
-  Load class Excel sheet (`section A.xlsx`) and map marks  
-  Automatically highlight unmatched rows in **red** for manual review  
-  Save final results in `output_with_marks.xlsx`  

---

## 📂 Project Structure
```bash
marks/
├── extract.py              # Fetches marks from Google Classroom → gcr_marks.txt
├── main.py                 # Maps marks into Excel and highlights missing matches
├── gcr_marks.txt           # Auto-generated JSON of marks
├── section A.xlsx          # Input Excel file with RollNo & Name
├── output_with_marks.xlsx  # Final output with mapped marks
├── credentials.json        # Google API credentials (not included)
└── README.md               # Documentation
```
---

## ⚙️ Installation

### Clone the Repository
```bash
git clone https://github.com/your-username/gcr-marks-extractor.git
cd gcr-marks-extractor
```
### Create Virtual Environment (optional but recommended)
```bash
python -m venv venv
source venv/bin/activate   # (Linux/Mac)
venv\Scripts\activate      # (Windows)
```

### Install Dependencies
```bash 
pip install google-auth google-auth-oauthlib google-api-python-client pandas openpyxl
```
---

## 🔑 Google API Setup
- Go to Google Cloud Console.
- Enable Google Classroom API.
- Create OAuth 2.0 Client ID credentials.
- Download credentials.json and place it in the project root.
- First run will open a browser for Google login and create a local token.

---
## 🚀 Usage
Step 1: Extract Marks
```bash 
python extract.py
```

- Select a course
- Select an assignment
- Script saves marks to gcr_marks.txt

Step 2: Update Excel
Make sure your Excel (section A.xlsx) has columns:
- RollNo
- Name

Then run:
```bash 
python main.py
```

This will:
- Match students by roll number (last 4 digits) or name
- Add a Marks column
- Highlight unmatched rows in red
- Save results to output_with_marks.xlsx

### 📝 Example Output (JSON)
```json
[
    {
        "name": "Ali Khan",
        "roll_number_last4": "1234",
        "marks": "85/100"
    },
    {
        "name": "Sara Ahmed",
        "roll_number_last4": "5678",
        "marks": "92/100"
    }
]
```

## ⚠️ Notes

Ensure your Excel sheet column names are exactly: RollNo and Name.
If roll numbers/names don’t match, rows will be highlighted red for manual review. **Google credentials (credentials.json) should not be shared publicly**