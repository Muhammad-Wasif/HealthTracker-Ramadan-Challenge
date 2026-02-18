# Health Tracker – Ramadan Challenge

A PySide6 desktop app to track daily Ramadan health habits: water intake, sugar/sweet drinks, steps/walking, sleep, protein/fruit/vegetables, and weight/BMI. Data is stored in Excel files and you can view history and charts.

## Features
- User registration/login (passwords hashed with bcrypt)
- Daily logs (avoid + healthy habits)
- BMI calculation + category
- History table + export to CSV
- Charts for key metrics (Matplotlib)

## Tech stack
- Python + PySide6 (GUI)
- Pandas/OpenPyXL (Excel storage)
- Matplotlib (charts)

## Run from source
1. Install Python
2. Install dependencies:
   pip install -r requirements.txt
3. Run:
   python main.py

## Build EXE (Windows)
Example:
py -m PyInstaller --clean --onefile --windowed --name "Health Tracker- Ramadan Challenge" --icon "Untitled design.ico" main.py

## Notes
- App creates `data/` and `data_backup/` next to the script and writes multiple `.xlsx` files there.
- Do not commit personal `.xlsx` data to GitHub.
