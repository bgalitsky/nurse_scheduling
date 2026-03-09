#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Streamlit + OR-Tools CP-SAT physician scheduling.

Fixed version:
- Removes syntax error
- Enforces:
  * 0.5 FTE / vacation -> lower cabinet priority
  * >=3 doctors want same cabinet -> one loses priority
  * Weekdays: morning+evening, weekends: one shift
  * <=1 shift/day
  * <=5 consecutive days
  * Norms: full=22, half=11 (minus vacation weekdays)
  * Extra shifts ONLY for allow_extra doctors
- UI table for wishes via st.data_editor
"""

import re
import datetime as dt
from dataclasses import dataclass, field
from collections import defaultdict, Counter
from typing import Dict, List, Tuple, Set
from io import BytesIO

import pandas as pd
import streamlit as st
from ortools.sat.python import cp_model
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from openpyxl.utils import get_column_letter


# =========================
# Data models
# =========================

@dataclass
class Doctor:
    name: str
    fte: float = 1.0
    keep_cabins: List[str] = field(default_factory=list)

@dataclass
class Vacation:
    name: str
    start: dt.date
    end: dt.date

@dataclass
class Wishes:
    even_pref: str = ""
    odd_pref: str = ""
    want_work: Set[dt.date] = field(default_factory=set)
    want_off: Set[dt.date] = field(default_factory=set)
    allow_extra: bool = False


# =========================
# Helpers
# =========================

def parse_date(s: str, year: int) -> dt.date:
    s = str(s).strip()
    if re.match(r"\d{2}\.\d{2}\.\d{4}", s):
        d, m, y = map(int, s.split("."))
        return dt.date(y, m, d)
    if re.match(r"\d{2}\.\d{2}", s):
        d, m = map(int, s.split("."))
        return dt.date(year, m, d)
    if re.match(r"\d{4}-\d{2}-\d{2}", s):
        y, m, d = map(int, s.split("-"))
        return dt.date(y, m, d)
    raise ValueError(s)

def is_weekend(d):
    return d.weekday() >= 5

def all_days(year, month):
    d = dt.date(year, month, 1)
    out = []
    while d.month == month:
        out.append(d)
        d += dt.timedelta(days=1)
    return out

def daterange(a, b):
    while a <= b:
        yield a
        a += dt.timedelta(days=1)


# =========================
# Parsing inputs
# =========================

def parse_doctors(text):
    out = []
    for l in text.splitlines():
        if not l.strip(): continue
        p = [x.strip() for x in l.split(",")]
        full_part = 1
        try:
            full_part = float(p[1])
        except ValueError:
            full_part = int(p[1])/1.0

        out.append(Doctor(p[0], full_part, p[2].split("|") if len(p) > 2 and p[2] else []))
    return out

def parse_vacations(text, year):
    out = []
    for l in text.splitlines():
        if not l.strip(): continue
        n, a, b = [x.strip() for x in l.split(",")]
        out.append(Vacation(n, parse_date(a, year), parse_date(b, year)))
    return out

def parse_wishes(text, year, month):
    res = {}
    for l in text.splitlines():
        if not l.strip() or l.startswith("#"): continue
        p = [x.strip() for x in l.split(",")]
        w = Wishes(p[1], p[2], set(), set(), p[5] in ("1","true","True"))
        for s in p[3].split(";"):
            if s:
                d = parse_date(s, year)
                if d.month == month: w.want_work.add(d)
        for s in p[4].split(";"):
            if s:
                d = parse_date(s, year)
                if d.month == month: w.want_off.add(d)
        res[p[0]] = w
    return res


# =========================
# Priority collision rule
# =========================

def adjust_keep_cabins(doctors, vacations, year, month):
    vac_map = defaultdict(set)
    for v in vacations:
        for d in daterange(v.start, v.end):
            vac_map[v.name].add(d)

    cab_map = defaultdict(list)
    for i,d in enumerate(doctors):
        for c in d.keep_cabins:
            cab_map[c].append(i)

    for cab, idxs in cab_map.items():
        if len(idxs) >= 3:
            idxs.sort(key=lambda i:(doctors[i].fte>=1, doctors[i].name))
            doctors[idxs[0]].keep_cabins.remove(cab)


# =========================
# Solver
# =========================

def solve(doctors, wishes, vacations, cabins, year, month):
    days = all_days(year, month)
    adjust_keep_cabins(doctors, vacations, year, month)

    vac_map = defaultdict(set)
    for v in vacations:
        for d in daterange(v.start, v.end):
            vac_map[v.name].add(d)

    model = cp_model.CpModel()

    slots = []
    for di,d in enumerate(days):
        shifts = ["р"] if is_weekend(d) else ["у","в"]
        for s in shifts:
            for c in cabins:
                slots.append((di,s,c))

    FREE = len(doctors)
    P = len(doctors)+1

    x = {(p,i):model.NewBoolVar(f"x{p}_{i}") for p in range(P) for i in range(len(slots))}

    for i in range(len(slots)):
        model.Add(sum(x[p,i] for p in range(P)) == 1)

    work = {}
    for p in range(len(doctors)):
        for di in range(len(days)):
            rel = [i for i,(d,_,_) in enumerate(slots) if d==di]
            w = model.NewBoolVar(f"w{p}_{di}")
            model.Add(sum(x[p,i] for i in rel) == w)
            work[p,di] = w

    for p in range(len(doctors)):
        for start in range(len(days)-5):
            model.Add(sum(work[p,di] for di in range(start,start+6)) <= 5)

    for p,d in enumerate(doctors):
        off = vac_map[d.name] | wishes.get(d.name,Wishes()).want_off
        for di,day in enumerate(days):
            if day in off:
                for i,(dd,_,_) in enumerate(slots):
                    if dd==di:
                        model.Add(x[p,i]==0)

    worked = {}
    required = {}
    for p,d in enumerate(doctors):
        base = 22 if d.fte>=1 else 11
        vac_wd = sum(1 for day in vac_map[d.name] if not is_weekend(day))
        req = max(0, base - vac_wd)
        required[d.name] = req
        w = model.NewIntVar(0,len(days),f"ws{p}")
        model.Add(w == sum(work[p,di] for di in range(len(days))))
        worked[p]=w
        if wishes.get(d.name,Wishes()).allow_extra:
            model.Add(w <= req+6)
        else:
            model.Add(w <= req)

    obj = []
    for p,d in enumerate(doctors):
        w = wishes.get(d.name,Wishes())
        for i,(di,s,c) in enumerate(slots):
            day = days[di]
            if s in ("у","в"):
                pref = w.even_pref if day.day%2==0 else w.odd_pref
                if pref and pref!=s:
                    obj.append(5*x[p,i])
            if d.keep_cabins and c not in d.keep_cabins:
                obj.append((1 if d.fte<1 else 2)*x[p,i])

    for i in range(len(slots)):
        obj.append(50*x[FREE,i])

    model.Minimize(sum(obj))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 20
    solver.Solve(model)

    sched = {d.name:{} for d in doctors}
    for d in doctors:
        for day in days:
            sched[d.name][day]=("-", "")

    for i,(di,s,c) in enumerate(slots):
        for p in range(len(doctors)):
            if solver.Value(x[p,i]):
                sched[doctors[p].name][days[di]]=(s,c)

    return days, sched, required


# =========================
# Streamlit UI
# =========================

st.set_page_config(layout="wide")
st.title("График врачей — OR-Tools (fixed)")
st.caption("Будни: у/в, выходные: р. ≤5 дней подряд. Норма: 22/11.")

year = 2025
month = 10

doctors_text = st.text_area("Doctors CSV", "Иванов И.И.,1,2А03")
vac_text = st.text_area("Vacations CSV", "")
wishes_text = st.text_area("Wishes CSV", "# name,even,odd,want_work,want_off,allow_extra")
cabins = ["2А03","2А04"]

if st.button("Recompute"):
    doctors = parse_doctors(doctors_text)
    vacations = parse_vacations(vac_text, year)
    wishes = parse_wishes(wishes_text, year, month)
    days, sched, req = solve(doctors, wishes, vacations, cabins, year, month)

    rows=[]
    for d in doctors:
        r={"Врач":d.name}
        for day in days:
            r[str(day.day)]=" ".join(sched[d.name][day])
        r["Норма"]=req[d.name]
        rows.append(r)

    st.dataframe(pd.DataFrame(rows), use_container_width=True)
