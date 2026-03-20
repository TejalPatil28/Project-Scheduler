# Project Scheduler

A web app to view and update project schedule Excel files.

## Folder Structure

```
tool/
├── backend/
│   ├── app.py              ← Flask server
│   ├── excel_db.py         ← Excel read/write layer
│   ├── users.xlsx          ← User credentials
│   └── requirements.txt
├── frontend/
│   ├── index.html
│   ├── style.css
│   ├── api.js
│   └── app.js
└── data/
    └── projects/
            OR-001.xlsx     ← Your project schedule files go here
            OR-002.xlsx
```

## Setup

### 1. Install Python dependencies
```bash
cd backend
pip install -r requirements.txt
```

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
