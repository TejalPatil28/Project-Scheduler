# Project Scheduler

A web-based project scheduling system that manages Excel-based project schedules across multiple departments (SW, HW, MFG, PM). Supports role-based access, real-time editing, and department-specific monitoring.

## Features

- **Multi-department support**: SW, HW, MFG, PM departments with separate schedule files and monitors
- **Role-based access**: Department heads, team leads, project managers, and admins
- **Real-time editing**: Editable cells with pending changes and background saving
- **JSON caching**: Fast loading with sidecar JSON files
- **Excel synchronization**: Background threads write to Excel without blocking UI
- **System memory flags**: Automatic calculation of PL/PM red activities, alert dates, and progress flags
- **Dark/Light theme**: User preference saved locally

## Folder Structure

```
project/
├── backend/
│ ├── app.py ← Flask server (routes, auth)
│ ├── excel_db.py ← Excel read/write/cache layer
│ ├── users.xlsx ← User credentials database
│ └── requirements.txt ← Python dependencies
├── frontend/
│ ├── index.html ← Main HTML entry
│ ├── style.css ← Global styles + themes
│ ├── api.js ← API wrapper
│ ├── app.js ← Main application logic
│ ├── monitor.js ← Monitor page (tabbed view)
│ └── users.js ← User management page
├── data/
│ ├── SW/ ← Software department
│ │ ├── SWESch/ ← SW schedule files
│ │ ├── SW_Monitor.xlsx ← SW project tracking
│ │ ├── PRS/ ← Project requirement sheets
│ │ ├── Template/ ← SW template file
│ │ └── cache/ ← JSON caches (auto-generated)
│ ├── HW/ ← Hardware department
│ │ ├── HWESch/ ← HW schedule files
│ │ ├── HW_Monitor.xlsx
│ │ ├── PRS/
│ │ ├── Template/
│ │ └── cache/
│ ├── MFG/ ← Manufacturing department
│ │ ├── MFGESch/
│ │ ├── MFG_Monitor.xlsx
│ │ ├── PRS/
│ │ ├── Template/
│ │ └── cache/
│ └── PM/ ← Project Management department
│ ├── PMESch/
│ ├── PM_Monitor.xlsx
│ ├── PRS/
│ ├── Template/
│ └── cache/
└── README.md
```

## Setup

### 1. Install Python dependencies
```bash
cd backend
pip install -r requirements.txt
```

Dependencies include:

openpyxl - Excel file handling
pycel - Excel formula evaluation
bcrypt - Password hashing
flask - Web server
flask-cors - CORS support
pyxlsb - Excel binary format support


### 2. Add your project Excel files
Copy your project schedule `.xlsx` files into `data/projects/`

### 3. Set TL assignments in each Excel file
Each project Excel needs the team assignments written in cells:
- `E1` → PM name
- `E2` → HW TL name
- `E3` → SW TL name
- `E4` → MFG TL name

These names must match exactly what is in `users.xlsx` (name column).

### 4. Edit users.xlsx
Open `backend/users.xlsx` and add your actual users:
| name | username | password_hash | role |
|------|----------|---------------|------|
| (full name matching Excel) | (login username) | (run hash script) | admin/head/pm/hw_tl/sw_tl/mfg_tl |

To generate a password hash:
```python
import bcrypt
print(bcrypt.hashpw(b"yourpassword", bcrypt.gensalt()).decode())
```

### 5. Run the app
```bash
cd backend
python app.py
```

Open http://localhost:5000 in your browser.

## Default Users (change these!)
| Username | Password | Role |
|----------|----------|------|
| admin    | admin123 | Admin |
| head     | head123  | Head |
| priya    | priya123 | PM |
| rahul    | rahul123 | HW TL |
| sara     | sara123  | SW TL |
| mohit    | mohit123 | MFG TL |

## Roles
- **Admin / Head** — see and edit all projects and all tasks
- **PM** — see and edit only their assigned projects (all tasks)
- **HW TL / SW TL / MFG TL** — see and edit only their assigned projects, only their own tasks
