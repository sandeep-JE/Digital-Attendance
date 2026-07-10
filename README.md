# Facial Attendance Web App

This project is a local web app for employee registration, facial attendance, secure login, and admin review. It uses the system camera through the browser, stores application data in an Excel workbook, registers face encodings for each employee, and supports user accounts with admin-only controls.

## Features

- Employee CRUD: create, edit, delete, and review employee records
- User master with admin and user roles
- Login, logout, and forgot-password reset flow
- Active and inactive status control
- Face registration from the connected camera
- Continuous attendance scanning from the connected camera
- Configurable cooldown between repeated face scans for the same person
- Attendance duration from first timestamp to last timestamp of the day
- Auto status rules: `>= 8h = present`, `>= 4h = halfday`, `< 4h = admin final decision`
- Overtime calculation for duration beyond 8 hours
- Clean validation and friendly error messages
- Excel workbook storage for employees, users, settings, attendance, and final decisions

## Setup

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

## Run

```bash
python app.py
```

Then open `http://127.0.0.1:8000`.

## Default Admin Login

- Username: `admin`
- Password: `admin123`

Change the default password immediately from User Master after first login.

## Data Storage

- Employee and attendance records are stored in `attendance_records.xlsx`
- User accounts, settings, and final attendance decisions are also stored in `attendance_records.xlsx`
- Face image files are stored in `data/employee_faces`

## Main Pages

- `/` dashboard with summary cards and recent attendance
- `/employees` employee list and status controls
- `/employees/new` new employee form
- `/employees/<id>/edit` employee edit page and face registration
- `/attendance` live attendance scanner
- `/users` user master for login accounts
- `/attendance/review` admin-only final attendance decision page
- `/settings` cooldown configuration page

## Notes

- The browser must be allowed to access the local camera.
- Registration and attendance both expect exactly one face in frame.
- Attendance timestamps are stored as scan events. Daily duration is calculated from the first timestamp to the last timestamp for that day.
- If only one timestamp exists for a day, the duration is `0` and the row goes to admin final decision.
- Short days below 4 hours are reviewed and finalized by admin from the Final Decision page.
- If `attendance_records.xlsx` is open in Excel while the app is trying to write data, the app will block the action and ask you to close the file first.
- `face_recognition_models` currently emits a `pkg_resources` deprecation warning upstream; the pinned `setuptools<81` dependency keeps it working on this environment.
