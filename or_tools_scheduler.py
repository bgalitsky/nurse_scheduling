from __future__ import annotations
from dataclasses import dataclass
from typing import Dict, List, Tuple, Set, Optional

import pandas as pd
from ortools.sat.python import cp_model


@dataclass
class Weights:
    w_pref: int = 5
    w_understaff: int = 200
    w_overstaff: int = 5
    w_fairness: int = 3
    w_weekend: int = 2  # penalty per weekend assignment (optional knob)


def parse_semicolon_set(s: str) -> Set[str]:
    if pd.isna(s) or not str(s).strip():
        return set()
    return {x.strip() for x in str(s).split(";") if x.strip()}


def pref_lookup(pref_df: pd.DataFrame) -> Dict[Tuple[str, str, str], int]:
    m: Dict[Tuple[str, str, str], int] = {}
    for _, r in pref_df.iterrows():
        m[(str(r["nurse_id"]), str(r["day"]), str(r["shift"]))] = int(r["preference"])
    return m


def locks_lookup(locks_df: pd.DataFrame) -> Dict[Tuple[str, str, str, str], int]:
    m: Dict[Tuple[str, str, str, str], int] = {}
    for _, r in locks_df.iterrows():
        m[(str(r["day"]), str(r["shift"]), str(r["room_id"]), str(r["nurse_id"]))] = int(r.get("locked", 0))
    return m


def solve_schedule_ortools(
    nurses_df: pd.DataFrame,
    rooms_df: pd.DataFrame,
    demand_df: pd.DataFrame,
    pref_df: pd.DataFrame,
    locks_df: pd.DataFrame,
    days: List[str],
    shifts: List[str],
    allow_overstaff: bool,
    weights: Weights,
    time_limit_seconds: int,
    target_shifts_per_nurse_week: Optional[int] = None,
    enforce_rest_night_to_day: bool = True,
    charge_rooms_tags: Set[str] = frozenset({"ICU", "ER"}),
    require_cnor_in_or: bool = True,
) -> Tuple[pd.DataFrame, Dict]:

    nurses_df = nurses_df.copy()
    rooms_df = rooms_df.copy()
    demand_df = demand_df.copy()
    pref_df = pref_df.copy()
    locks_df = locks_df.copy()

    nurses_df["qual_set"] = nurses_df["qualifications"].apply(parse_semicolon_set)
    rooms_df["req_set"] = rooms_df["required_qualifications"].apply(parse_semicolon_set)

    nurse_ids = nurses_df["nurse_id"].astype(str).tolist()
    room_ids = rooms_df["room_id"].astype(str).tolist()

    max_per_day = {r["nurse_id"]: int(r["max_shifts_per_day"]) for _, r in nurses_df.iterrows()}
    max_per_week = {r["nurse_id"]: int(r["max_shifts_per_week"]) for _, r in nurses_df.iterrows()}
    cert_by_nurse = {r["nurse_id"]: ("" if pd.isna(r["certification"]) else str(r["certification"])) for _, r in nurses_df.iterrows()}
    qual_by_nurse = {r["nurse_id"]: r["qual_set"] for _, r in nurses_df.iterrows()}

    req_by_room = {r["room_id"]: r["req_set"] for _, r in rooms_df.iterrows()}
    tag_by_room = {r["room_id"]: ("" if pd.isna(r.get("tags", "")) else str(r.get("tags", ""))) for _, r in rooms_df.iterrows()}

    # demand[(day, room, shift)]
    demand = {}
    for _, r in demand_df.iterrows():
        demand[(str(r["day"]), str(r["room_id"]), str(r["shift"]))] = int(r["required_nurses"])
    for d in days:
        for rid in room_ids:
            for sh in shifts:
                demand.setdefault((d, rid, sh), 0)

    pref = pref_lookup(pref_df)
    locks = locks_lookup(locks_df)

    def feasible(n: str, rid: str) -> bool:
        return req_by_room[rid].issubset(qual_by_nurse[n])

    model = cp_model.CpModel()

    # x[n, r, sh, d] ∈ {0,1} only if feasible
    x: Dict[Tuple[str, str, str, str], cp_model.IntVar] = {}
    for n in nurse_ids:
        for rid in room_ids:
            if not feasible(n, rid):
                continue
            for sh in shifts:
                for d in days:
                    x[(n, rid, sh, d)] = model.NewBoolVar(f"x_{n}_{rid}_{sh}_{d}")

    # Nurse: max shifts per day
    for n in nurse_ids:
        for d in days:
            vars_nd = [x[k] for k in x if k[0] == n and k[3] == d]
            if vars_nd:
                model.Add(sum(vars_nd) <= max_per_day[n])

    # Nurse: max shifts per week (Mon–Sun horizon)
    for n in nurse_ids:
        vars_n = [x[k] for k in x if k[0] == n]
        if vars_n:
            model.Add(sum(vars_n) <= max_per_week[n])

    # Coverage with slack
    understaff = {}
    overstaff = {}
    for d in days:
        for rid in room_ids:
            for sh in shifts:
                assigned = [x[(n, rid, sh, d)] for n in nurse_ids if (n, rid, sh, d) in x]
                assigned_sum = sum(assigned) if assigned else 0
                u = model.NewIntVar(0, 100, f"under_{rid}_{sh}_{d}")
                understaff[(d, rid, sh)] = u
                if allow_overstaff:
                    o = model.NewIntVar(0, 100, f"over_{rid}_{sh}_{d}")
                    overstaff[(d, rid, sh)] = o
                    model.Add(assigned_sum + u - o == demand[(d, rid, sh)])
                else:
                    model.Add(assigned_sum + u == demand[(d, rid, sh)])

    # Rest rule: no Night -> Day next day (per nurse)
    if enforce_rest_night_to_day and ("Night" in shifts) and ("Day" in shifts):
        day_index = {d: i for i, d in enumerate(days)}
        inv_day = {i: d for d, i in day_index.items()}

        for n in nurse_ids:
            for i in range(len(days) - 1):
                d = inv_day[i]
                d_next = inv_day[i + 1]

                night_vars = [x[k] for k in x if k[0] == n and k[2] == "Night" and k[3] == d]
                day_vars_next = [x[k] for k in x if k[0] == n and k[2] == "Day" and k[3] == d_next]

                if night_vars and day_vars_next:
                    model.Add(sum(night_vars) + sum(day_vars_next) <= 1)

    # Charge nurse requirement per shift/day in ICU + ER:
    # For any (day, shift), sum of assigned charge nurses in those rooms >= 1, IF total demand there > 0
    charge_nurses = [n for n in nurse_ids if "Charge" in qual_by_nurse[n]]
    charge_room_ids = [rid for rid in room_ids if tag_by_room.get(rid, "") in charge_rooms_tags]

    if charge_room_ids and charge_nurses:
        for d in days:
            for sh in shifts:
                total_demand = sum(demand[(d, rid, sh)] for rid in charge_room_ids)
                if total_demand <= 0:
                    continue
                charge_assigned = []
                for rid in charge_room_ids:
                    for n in charge_nurses:
                        key = (n, rid, sh, d)
                        if key in x:
                            charge_assigned.append(x[key])
                if charge_assigned:
                    model.Add(sum(charge_assigned) >= 1)

    # Skill mix: if OR room is staffed (demand>0), at least one assigned nurse must have CNOR certification
    if require_cnor_in_or:
        or_room_ids = [rid for rid in room_ids if tag_by_room.get(rid, "") == "OR"]
        for d in days:
            for rid in or_room_ids:
                for sh in shifts:
                    if demand[(d, rid, sh)] <= 0:
                        continue
                    cnor_vars = []
                    for n in nurse_ids:
                        if cert_by_nurse[n] == "CNOR":
                            key = (n, rid, sh, d)
                            if key in x:
                                cnor_vars.append(x[key])
                    # If OR is running and there is any feasible CNOR, require ≥1
                    if cnor_vars:
                        model.Add(sum(cnor_vars) >= 1)

    # Locks: enforce x=1 for locked rows (hard)
    for (d, sh, rid, n), locked in locks.items():
        if locked != 1:
            continue
        key = (n, rid, sh, d)
        if key not in x:
            raise RuntimeError(f"Locked assignment infeasible (qualification mismatch or unknown IDs): {d},{sh},{rid},{n}")
        model.Add(x[key] == 1)

    # Fairness: penalize deviation from target weekly shifts per nurse (optional)
    fairness_dev = {}
    if target_shifts_per_nurse_week is not None:
        for n in nurse_ids:
            total = model.NewIntVar(0, 1000, f"tot_{n}")
            vars_n = [x[k] for k in x if k[0] == n]
            model.Add(total == (sum(vars_n) if vars_n else 0))
            dev = model.NewIntVar(0, 1000, f"dev_{n}")
            model.Add(dev >= total - target_shifts_per_nurse_week)
            model.Add(dev >= target_shifts_per_nurse_week - total)
            fairness_dev[n] = dev

    # Objective: maximize preferences, minimize slack + fairness + optional weekend aversion
    obj = []

    # Preferences reward
    for (n, rid, sh, d), var in x.items():
        p = int(pref.get((n, d, sh), 0))
        if p:
            obj.append(weights.w_pref * p * var)

    # Under/over staffing penalties
    for (d, rid, sh), u in understaff.items():
        obj.append(-weights.w_understaff * u)
    if allow_overstaff:
        for (d, rid, sh), o in overstaff.items():
            obj.append(-weights.w_overstaff * o)

    # Fairness penalties
    for dev in fairness_dev.values():
        obj.append(-weights.w_fairness * dev)

    # Weekend penalty (soft): discourage Sat/Sun assignments slightly
    weekend_days = {d for d in days if d in ("Sat", "Sun")}
    if weekend_days and weights.w_weekend > 0:
        for (n, rid, sh, d), var in x.items():
            if d in weekend_days:
                obj.append(-weights.w_weekend * var)

    model.Maximize(sum(obj))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = float(time_limit_seconds)
    solver.parameters.num_search_workers = 8

    status = solver.Solve(model)

    meta = {
        "status": solver.StatusName(status),
        "objective": solver.ObjectiveValue() if status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None,
    }

    # Output schedule rows: each (day, room, shift) with assigned nurse list
    out = []
    for d in days:
        for rid in room_ids:
            for sh in shifts:
                assigned = []
                for n in nurse_ids:
                    key = (n, rid, sh, d)
                    if key in x and solver.Value(x[key]) == 1:
                        assigned.append(n)
                out.append({
                    "day": d,
                    "room_id": rid,
                    "room_name": rooms_df.loc[rooms_df["room_id"] == rid, "room_name"].iloc[0],
                    "shift": sh,
                    "required_nurses": demand[(d, rid, sh)],
                    "assigned_nurses": ";".join(assigned),
                    "understaff": int(solver.Value(understaff[(d, rid, sh)])) if status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None,
                    "overstaff": int(solver.Value(overstaff[(d, rid, sh)])) if allow_overstaff and status in (cp_model.OPTIMAL, cp_model.FEASIBLE) else None,
                })

    schedule_df = pd.DataFrame(out)
    return schedule_df, meta
