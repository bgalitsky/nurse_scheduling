= Nurse Scheduling CSP System =
Automated hospital shift assignment using constraint solving

This project generates a nurse–room schedule for a hospital using an open-source CSP (Constraint Satisfaction Problem) solver.
It assigns nurses to rooms based on their qualifications, certifications, and shift-hour limits, ensuring each room is staffed with a qualified nurse.

The system reads nurse and facility data from CSV files, models the scheduling problem, and computes a valid assignment for a 6-hour shift.

Input Data Format
1. nurses.csv columns
Column	Meaning
nurse_id	Unique nurse identifier (N01 … N20)
full_name	Nurse name
degree	ADN/BSN/MSN
certification	Optional clinical certifications e.g. CEN, CCRN
max_daily_hours	Max shift hours allowed (6-hour limit enforced)
qualifications	Semi-colon list of clinical skills (RN; ICU; OR; Pediatrics)
2. facilities.csv columns
Column	Meaning
room_id	Unique room identifier
room_name	Name (ICU, Oncology, Pediatric Ward, etc.)
shift_start / shift_end	Shift timing (single daily shift here)
min_nurses	Minimum staff required (1 by default)
required_qualifications	Skills needed (RN; ICU, RN; OR)
▶ Run Scheduler
python scheduler_csp.py


If constraints are satisfiable, output looks like:

Room schedule for 08:00–14:00 shift (CSP solver):
R1 | Emergency         | 08:00-14:00 | N01 – Alice Smith (RN,ER,Procedural)
R2 | ICU               | 08:00-14:00 | N02 – Bob Johnson (RN,ICU,Charge)
R3 | Operating Room 1  | 08:00-14:00 | N03 – Carol Lee (RN,OR,Procedural)
...


Remaining nurses will be listed as float/backup staff.

🔍 How It Works

The solver models scheduling as a constraint search problem:

CSP Element	Project mapping
Variables	Each room = 1 variable
Domains	Nurses qualified for that room
Hard Constraints	Qualifications, max hours
Global Constraint	AllDifferent → nurse can't cover two rooms

The solver returns a complete feasible assignment of nurses to rooms.
If no valid solution exists, constraints must be adjusted (we can automate this later).
