import csv
from constraint import Problem, AllDifferentConstraint

NURSE_CSV = "../hospital/nurses.csv"
FACILITY_CSV = "../hospital/facilities.csv"

SHIFT_HOURS = 6.0  # 08:00–14:00

def load_nurses(path):
    nurses = {}
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            quals = row["qualifications"].split(";") if row["qualifications"] else []
            nurses[row["nurse_id"]] = {
                "full_name": row["full_name"],
                "degree": row["degree"],
                "certification": row["certification"],
                "max_daily_hours": float(row["max_daily_hours"]),
                "qualifications": set(q.strip() for q in quals if q.strip()),
            }
    return nurses

def load_rooms(path):
    rooms = []
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            reqs = row["required_qualifications"].split(";") if row["required_qualifications"] else []
            rooms.append({
                "room_id": row["room_id"],
                "room_name": row["room_name"],
                "shift_start": row["shift_start"],
                "shift_end": row["shift_end"],
                "min_nurses": int(row["min_nurses"]),
                "required_qualifications": set(q.strip() for q in reqs if q.strip()),
            })
    return rooms

def build_domains(nurses, rooms):
    """
    Build domains for each room:
      - only nurses whose qualifications cover room requirements
      - only nurses whose max_daily_hours >= SHIFT_HOURS
    """
    domains = {}
    for room in rooms:
        rid = room["room_id"]
        reqs = room["required_qualifications"]
        domain = []
        for nid, n in nurses.items():
            if n["max_daily_hours"] < SHIFT_HOURS:
                continue
            if not reqs.issubset(n["qualifications"]):
                continue
            domain.append(nid)
        domains[rid] = domain
    return domains

def solve_schedule(nurses, rooms):
    """
    CSP encoding:
      - Variable for each room (room_id).
      - Domain = list of nurses qualified for that room (and enough hours).
      - Global constraint: AllDifferent (each nurse can cover at most one room in this shift).
    """
    domains = build_domains(nurses, rooms)

    # Quick feasibility check (no domain must be empty)
    for rid, domain in domains.items():
        if not domain:
            raise RuntimeError(f"No qualified nurses available for room {rid}.")

    problem = Problem()

    # Add variables with their domains
    for rid, domain in domains.items():
        problem.addVariable(rid, domain)

    # Each nurse can only be assigned to at most one room (AllDifferent over nurse IDs)
    room_ids = [r["room_id"] for r in rooms]
    problem.addConstraint(AllDifferentConstraint(), room_ids)

    # Solve
    solutions = problem.getSolutions()
    if not solutions:
        raise RuntimeError("No feasible schedule found given current constraints.")

    # You can choose any solution; here we pick the first
    schedule = solutions[0]
    return schedule

def main():
    nurses = load_nurses(NURSE_CSV)
    rooms = load_rooms(FACILITY_CSV)
    schedule = solve_schedule(nurses, rooms)

    print("Room schedule for 08:00–14:00 shift (CSP solver):")
    print("-" * 80)
    for room in sorted(rooms, key=lambda r: r["room_id"]):
        rid = room["room_id"]
        rname = room["room_name"]
        nid = schedule.get(rid)
        if nid is None:
            nurse_str = "(unassigned)"
        else:
            nurse = nurses[nid]
            nurse_str = f"{nid} – {nurse['full_name']} ({', '.join(sorted(nurse['qualifications']))})"
        print(f"{rid:>3} | {rname:<20} | {room['shift_start']}-{room['shift_end']} | {nurse_str}")

    # Nurses without a room assignment (float/backup)
    assigned_nurses = set(schedule.values())
    unassigned = [nid for nid in nurses.keys() if nid not in assigned_nurses]

    print("\nUnassigned nurses (available as float/backup):")
    for nid in sorted(unassigned):
        n = nurses[nid]
        print(f"{nid} – {n['full_name']}")

if __name__ == "__main__":
    main()
