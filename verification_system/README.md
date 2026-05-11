# Student Result Verification System

A web-based application that allows officers to verify student academic records with an audit trail. Based on the original KMTC certificate issuance system, this application provides a modern interface for managing and tracking student result verifications.

## Features

✓ **Officer Authentication** - Officer name stored locally for convenience
✓ **Student Lookup** - Search students by roll number with instant access to all academic details
✓ **Automatic Excel Highlighting** - When a student is looked up, their row is automatically highlighted in RED in the Excel file
✓ **Verification Actions** - Two verification states:
  - 📝 **Issued** - Certificate/record issued (highlights row RED in Excel)
  - ✎ **Corrected** - Errors corrected (no highlighting)

✓ **Comments Support** - Add optional notes for each verification
✓ **Audit Trail** - Complete timestamped history of all verification actions
✓ **Real-time Statistics** - Dashboard showing verification counts by action type
✓ **Lookup Tracking** - View all students who have been looked up with timestamps
✓ **Excel Download** - Download the complete Excel file with highlighted rows
✓ **CSV Data Source** - Works directly with Excel/CSV files (AI_Data.csv)
✓ **Responsive Design** - Works on desktop and tablet devices

## Project Structure

```
verification_system/
├── app.py                 # Flask backend server
├── requirements.txt       # Python dependencies
├── README.md             # This file
├── templates/
│   └── index.html        # Main HTML/CSS/JavaScript interface
└── audit_trail.json      # Audit trail database (created automatically)
```

## Technical Stack

- **Backend**: Python Flask with CORS support
- **Frontend**: HTML5, CSS3, JavaScript (vanilla)
- **Data Source**: CSV file (AI_Data.csv)
- **Audit Trail**: JSON file (audit_trail.json)
- **Storage**: Lightweight, file-based (no external database required)

## Installation & Setup

### Prerequisites

- Python 3.7 or higher
- pip (Python package manager)

### Step 1: Navigate to Project Directory

```powershell
cd "C:\Users\lmutu\Downloads\LOOKUP APP\verification_system"
```

### Step 2: Create Virtual Environment (Optional but Recommended)

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

### Step 3: Install Dependencies

```powershell
pip install -r requirements.txt
```

### Step 4: Run the Application

```powershell
python app.py
```

You should see output like:
```
Loaded 67 students from CSV
CSV file: C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data.csv
Audit trail: C:\Users\lmutu\Downloads\LOOKUP APP\verification_system\audit_trail.json
 * Running on http://127.0.0.1:5000
```

### Step 5: Open in Browser

Open your web browser and navigate to:
```
http://127.0.0.1:5000
```

## Usage Guide

### 1. Enter Officer Information
- Type your name in the "Officer Name" field
- Your name will be automatically saved to your browser for future sessions

### 2. Search for Student
- Enter the student's roll number (e.g., 0403AL231001)
- Click "Search" or press Enter
- The system will display all student details:
  - Name, Roll Number, Program, Branch
  - Current academic status and semester
  - Subject grades for all subjects
  - CGPA and Result Description

### 3. Select Verification Action
Choose one of two actions:
- **Issued** 📝 - Mark that certificate/record has been issued (highlights row RED in Excel)
- **Corrected** ✎ - Indicate corrections have been made (no highlighting)

### 4. Add Comments (Optional)
Include any notes or additional context about the verification.

### 5. Submit
Click "Submit Verification" to record the action.
The system will:
- Save the verification with timestamp
- Log officer name, student details, and action
- Update statistics
- Clear the form for the next verification

### 6. View Dashboard
- **Statistics Tab**: See total verifications and breakdown by action type
- **Recent Activity Tab**: View the last 20 verification records with timestamps
- **Lookup History Tab**: See all students who have been looked up (their rows are highlighted RED in Excel)
- **Downloads Tab**: Download the Excel file with highlighted rows

### 7. Track Lookups in Excel
When you search for a student:
- Their row is automatically highlighted **RED** with white text in the Excel file
- A timestamp is added in the last column showing when they were looked up
- The Excel file updates in real-time
- You can download the updated file anytime from the Downloads tab

## Excel Highlighting Feature

**How It Works:**
1. Students are searched and their details are displayed
2. When you select **"Issued"** and submit, the student's row is highlighted **RED** in the Excel file
3. If you select **"Corrected"**, no highlighting occurs
4. The timestamp of the verification is recorded in the last column
5. Multiple "Issued" actions on the same student overwrite the timestamp with the latest

**File Locations:**
- **Source CSV**: `C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data.csv` (Read-only)
- **Excel Tracker**: `C:\Users\lmutu\Downloads\LOOKUP APP\verification_system\AI_Data_Lookup_Tracker.xlsx` (Auto-created, updated on verifications)

**Highlighted Rows:**
- Row background: 🔴 RED with white text (only for "Issued" actions)
- Timestamp column added showing when certificate was issued
- Preserves all original student data

## API Endpoints

### GET `/`
Main application interface (HTML page)

### GET `/api/student/<roll_no>`
Retrieve student details by roll number
**Important**: Automatically highlights the student row RED in Excel when called

**Response:**
```json
{
  "success": true,
  "data": {
    "Name": "AAKASH PIPALDE",
    "Roll No": "0403AL231001",
    "CGPA": "6.71",
    ...
  }
}
```

### POST `/api/verify`
Submit a verification action

**Request:**
```json
{
  "officer_name": "John Doe",
  "roll_no": "0403AL231001",
  "student_name": "AAKASH PIPALDE",
  "action": "Verified",
  "comments": "Record verified successfully"
}
```

**Response:**
```json
{
  "success": true,
  "message": "Verification recorded: Verified",
  "timestamp": "2026-05-07T14:30:45.123456"
}
```

### GET `/api/audit-trail`
Get all verification records

**Response:**
```json
{
  "success": true,
  "records": [
    {
      "timestamp": "2026-05-07T14:30:45.123456",
      "officer_name": "John Doe",
      "roll_no": "0403AL231001",
      "student_name": "AAKASH PIPALDE",
      "action": "Verified",
      "comments": "Record verified successfully"
    }
  ]
}
```

### GET `/api/stats`
Get verification statistics

**Response:**
```json
{
  "success": true,
  "stats": {
    "total_verifications": 5,
    "verified": 3,
    "issued": 1,
    "corrected": 1,
    "disputed": 0
  }
}
```

### GET `/api/lookup-history`
Get all student lookups with timestamps

**Response:**
```json
{
  "success": true,
  "records": [
    {
      "timestamp": "2026-05-07T14:30:45.123456",
      "roll_no": "0403AL231001",
      "student_name": "AAKASH PIPALDE"
    }
  ]
}
```

### GET `/api/download-excel`
Download the Excel file with highlighted rows

**Returns:** Excel file (.xlsx) with all students and highlighted rows for those that have been looked up

## Data Storage

### CSV File (Read-only)
- **Location**: `C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data.csv`
- **Purpose**: Student data source
- **Contains**: 67+ student records with roll numbers, names, grades, SGPA, CGPA, and result statuses

### Excel File (Auto-created)
- **Location**: `C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data_Lookup_Tracker.xlsx`
- **Purpose**: Formatted Excel with highlighted rows
- **Auto-updated**: Every student lookup highlights their row RED with timestamp

### Audit Trail (Auto-created)
- **Location**: `verification_system/audit_trail.json`
- **Purpose**: Timestamped log of all verification actions
- **Auto-updated**: Every verification adds a new entry

### Lookup History (Auto-created)
- **Location**: `verification_system/lookup_history.json`
- **Purpose**: Timestamped log of all student lookups
- **Auto-updated**: Every student search adds a new entry

## Browser Local Storage

- **Officer Name**: Stored in browser localStorage for convenience
- **Persistence**: Survives across sessions until manually cleared
- **Privacy**: Only stored locally on the officer's computer

## Troubleshooting

### "Student not found" Error
- Check the roll number spelling and format
- Example valid format: 0403AL231001
- Ensure you're using the exact roll number from the CSV

### Port 5000 Already in Use
If port 5000 is already in use, modify `app.py`:
```python
app.run(debug=True, host='127.0.0.1', port=5001)  # Use different port
```

### CSV File Not Found
Ensure the CSV file is in: `C:\Users\lmutu\Downloads\LOOKUP APP\AI_Data.csv`
The file path is hardcoded in app.py line 18.

### CORS Error in Console
This shouldn't occur as CORS is enabled, but if it does:
- Ensure the Flask server is running
- Check that you're accessing via `http://127.0.0.1:5000` (not `localhost`)

## Exporting Audit Trail

To view the complete audit trail as JSON:
```powershell
# In PowerShell
Get-Content "verification_system\audit_trail.json" | ConvertFrom-Json | Format-Table
```

Or convert to CSV for Excel:
```powershell
Get-Content "verification_system\audit_trail.json" | ConvertFrom-Json | Export-Csv "audit_report.csv" -NoTypeInformation
```

## Security Notes

- ⚠️ This system is designed for **internal use only** on a local network
- Officer names are stored in browser localStorage (not encrypted)
- For production deployment, implement:
  - User authentication
  - Database encryption
  - HTTPS/SSL
  - Role-based access control
  - Audit log backups

## Extending the System

### Add Database Support
Replace JSON audit trail with PostgreSQL/MySQL:
1. Update `app.py` to use SQLAlchemy ORM
2. Create a database schema
3. Modify `log_audit_trail()` and `get_audit_trail()` functions

### Add Email Notifications
Send emails when records are disputed:
```python
from flask_mail import Mail, Message
# Add email logic to verify_student() function
```

### Add Export Features
Generate PDF certificates or Excel reports:
```python
from reportlab.pdfgen import canvas
# Add export endpoints
```

## Support & Maintenance

For issues or enhancements:
1. Check the troubleshooting section above
2. Review the Flask console output for error messages
3. Check browser Developer Tools (F12) for JavaScript errors
4. Verify CSV file format and data consistency

## License

Internal KMTC Application - For authorized use only

---

**Created**: May 7, 2026
**Version**: 1.0
**Status**: Production Ready
