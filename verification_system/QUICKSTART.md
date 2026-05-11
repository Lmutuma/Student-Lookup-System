# Quick Start Guide

## 🚀 Getting Started in 3 Steps

### Step 1: Run the Application

**Option A: Using Batch File (Easiest)**
```
Double-click: run.bat
```
The system will automatically:
- Create a virtual environment
- Install dependencies
- Start the server
- Open the app in your browser

**Option B: Using PowerShell**
```powershell
powershell -ExecutionPolicy Bypass -File run.ps1
```

**Option C: Manual Setup**
```powershell
# Open PowerShell in the verification_system folder
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

### Step 2: Access the Application
Open your browser and go to:
```
http://127.0.0.1:5000
```

### Step 3: Start Verifying

1. **Enter your name** in the "Officer Name" field
2. **Enter a student roll number** (e.g., 0403AL231001)
3. **Click Search** to view student details
   - ✨ **Student's row is automatically highlighted RED in Excel!**
4. **Select an action**: Issued 📝 | Corrected ✎
5. **Add comments** (optional)
6. **Click Submit** to record the verification
   - ✨ **If "Issued" is selected, student's row is highlighted RED in Excel!**

---

## 📊 Automatic Excel Highlighting

**How it works:**
- When you search for a student, their details are displayed
- When you select **"Issued"** and submit, their row is **highlighted RED** in the Excel file
- If you select **"Corrected"**, no highlighting occurs
- The verification timestamp is recorded
- You can download the Excel file anytime to see which students have been issued certificates
- Go to **Downloads tab** → Click "⬇ Download Excel File"

---

## 💾 Where's My Data?

### CSV File (Student Records)
```
C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data.csv
```
Contains 67+ student records with all academic details.

### Audit Trail (Verification History)
```
C:\Users\lmutu\Downloads\LOOKUP APP\verification_system\audit_trail.json
```
Created automatically after first verification. Contains timestamped log of all actions.

### Excel Tracking File
```
C:\Users\lmutu\Downloads\LOOKUP APP\verification_system\AI_Data_Lookup_Tracker.xlsx
```
Created automatically when you first run the app. Contains highlighted rows for issued certificates.

---

## 📊 Dashboard Features

**Statistics Tab:**
- Total verifications
- Breakdown by action type (Issued, Corrected)

**Recent Activity Tab:**
- Last 20 verification records
- Timestamp, officer name, student name, action

---

## 🎯 Sample Roll Numbers to Try

From the CSV file:
- `0403AL231001` - AAKASH PIPALDE
- `0403AL231002` - AAYUSH PATEL
- `0403AL231004` - ABIR SAXENA
- `0403AL231005` - AMEY BHOKARIKAR

---

## ❓ Common Questions

**Q: Where is my officer name saved?**
A: In your browser's local storage. It will appear next time you use the system from the same computer.

**Q: Can I use this on multiple computers?**
A: Yes! Each computer stores its own officer name. The audit trail is on the server.

**Q: What happens if I close the browser?**
A: The server keeps running. Your verifications are saved. Just reload http://127.0.0.1:5000

**Q: How do I stop the server?**
A: Press `Ctrl+C` in the terminal/PowerShell window where it's running.

**Q: Can I restart and resume?**
A: Yes! Just run the application again. All previous verifications are preserved.

---

## 🛠 System Requirements

- Windows 7 or later
- Python 3.7+
- Modern web browser (Chrome, Firefox, Edge, Safari)
- ~50 MB disk space
- No internet required (works fully offline)

---

## 📞 Support

**Error: "Python not found"**
- Install Python from https://www.python.org/
- Make sure to check "Add Python to PATH" during installation

**Error: "Port 5000 already in use"**
- Another app is using port 5000
- Open Task Manager → End the process using port 5000
- Or edit app.py to use a different port (5001, 5002, etc.)

**Error: "CSV file not found"**
- Ensure AI_Data.csv is in: `C:\Users\lmutu\Downloads\LOOKUP APP\`
- Don't move or rename the CSV file

---

## 🔒 Security & Privacy

- Officer names are stored **locally only** (browser storage)
- Verification records are saved on the server machine
- No cloud upload or external transmission
- For production use, add authentication & encryption

---

**Version**: 1.0
**Last Updated**: May 7, 2026
**Status**: Ready to Use ✓
