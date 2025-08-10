# DUPR Club Crawler (Flet desktop app)

A small desktop UI to log into DUPR, fetch a club's members and match history, export a neat Excel report, and analyze/simulate rating changes with a simple, data‑fitted (Elo‑like) formula.

<img width="2500" height="1500" alt="image" src="https://github.com/user-attachments/assets/93527243-dfd4-45b8-9e9a-a90f843b104e" />

## What you get
- Login (email/password) or “Login as Guest”
- Crawl a club by Club ID with limits
- Export to Excel: Club Info, Members, Player Profiles, Match History
- Analysis & Simulation on your exported Excel (fit K & scale, then simulate)
- Help panel explaining each parameter + quick examples

---
## Fresh‑clone setup

### Prerequisites
- Python 3.11+
- Git
- Windows/macOS/Linux

### 1) Clone this repository
```bash
git clone https://github.com/Senju14/tool-crawl-dupr-data.git Tool_Crawl_SportPlus 
cd Tool_Crawl_SportPlus
```

### 2) Create & activate a virtual environment
Windows (PowerShell):
```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
```
macOS/Linux (bash/zsh):
```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3) Install Python dependencies
```bash
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 4) Get the DUPR client (duprly)
This app imports `dupr_client` from a local `duprly/` folder.
```bash
# from the project root
git clone https://github.com/pkshiu/duprly.git duprly
```

### 5) Run the app
```bash
python app_flet.py
```
A desktop window should open. If it opens in the browser instead, that’s fine—the UI is the same.

---
## Using the app (step by step)

### Login
- Enter your DUPR email/password or click “Login as Guest”.

### Crawl & Export
- Fields:
  - Club ID: the DUPR club ID (string of digits from club page/URL).
  - Max Members: maximum members to fetch from the club.
  - Players to fetch history: how many of those members to fetch match history for.
  - Matches per player: cap how many matches to fetch per player.
  - Filename prefix: prefix for the exported Excel file.
- Click “Start Crawl & Export”.
- Output file is saved as: `prefix_<ClubID>_YYYYMMDD_HHMMSS.xlsx` (e.g., `dupr_club_5986040853_20250810_110230.xlsx`).

### Analysis & Simulation
1) Click “Load Excel for Analysis” and pick your exported `.xlsx`.
2) The app estimates parameters for a simple Elo‑style model:
   - `rating_after = rating_before + K * (result - expected)`
   - `expected = 1 / (1 + 10^(-(diff/scale)))`, with `diff = rating_before - opponent_rating_before`
   - It displays the fitted `K`, `scale`, and `MAE`.
3) In “Chọn ví dụ từ Excel”, pick a sample row to auto‑fill the form, or enter values manually:
   - Player rating before
   - Opponent rating before
   - Result (Win/Loss)
4) Click “Compute Simulation” to see Expected, Delta, and New Rating.

### Help
- Click the “Help” button (top‑right) to toggle a short explanation of each parameter and examples.

---
## Troubleshooting
- ImportError: `dupr_client`
  - Ensure the `duprly/` folder exists in the project root (step 4).
- HTTP 403 during crawl
  - The app retries by re‑logging. If it persists, lower the crawl limits or run again later.
- Crawl is slow or too large
  - Reduce: Max Members / Players to fetch history / Matches per player.
- Excel analysis can’t find “Match History”
  - The app falls back to first sheet if needed. Make sure the export completed successfully.
- Windows PowerShell activation
  - If `.venv\Scripts\Activate.ps1` is blocked, adjust ExecutionPolicy for your current session.

---
## Notes & Disclaimer
- The analysis uses an inferred, Elo‑like model fitted to your data; it is NOT DUPR’s official algorithm.
- For best estimates, load exports that contain enough recent matches.
