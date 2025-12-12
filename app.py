import pandas as pd
import streamlit as st

from or_tools_scheduler import solve_schedule_ortools, Weights

DATA_DIR = "data"

st.set_page_config(page_title="Nurse Scheduling (OR-Tools, Weekly)", layout="wide")
st.title("🏥 Nurse Scheduling — Weekly (Mon–Sun) + OR-Tools CP-SAT")

@st.cache_data
def load_csv(path: str) -> pd.DataFrame:
    return pd.read_csv(path)

def csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")

# Load
nurses = load_csv(f"{DATA_DIR}/nurses.csv")
rooms = load_csv(f"{DATA_DIR}/facilities.csv")
demand = load_csv(f"{DATA_DIR}/demand.csv")
prefs  = load_csv(f"{DATA_DIR}/preferences.csv")
locks  = load_csv(f"{DATA_DIR}/locks.csv")

st.sidebar.header("Horizon")
all_days = ["Mon","Tue","Wed","Thu","Fri","Sat","Sun"]
days = st.sidebar.multiselect("Days", options=all_days, default=all_days)

st.sidebar.header("Shifts")
all_shifts = ["Day","Evening","Night"]
shifts = st.sidebar.multiselect("Shifts", options=all_shifts, default=all_shifts)

st.sidebar.header("Hard constraints")
enforce_rest = st.sidebar.checkbox("Rest rule: no Night → Day next day", value=True)
charge_req = st.sidebar.checkbox("Require Charge nurse per shift/day in ICU+ER", value=True)
cnor_req = st.sidebar.checkbox("Require CNOR in each OR room when running", value=True)

st.sidebar.header("Soft constraints")
allow_overstaff = st.sidebar.checkbox("Allow overstaffing (soft penalty)", value=True)
time_limit = st.sidebar.slider("Time limit (seconds)", 1, 60, 20)

w_pref = st.sidebar.slider("Preference reward (w_pref)", 0, 20, 5)
w_under = st.sidebar.slider("Understaff penalty (w_understaff)", 50, 600, 200)
w_over = st.sidebar.slider("Overstaff penalty (w_overstaff)", 0, 50, 5)
w_fair = st.sidebar.slider("Fairness penalty (w_fairness)", 0, 30, 3)
w_weekend = st.sidebar.slider("Weekend penalty (w_weekend)", 0, 10, 2)

use_fairness = st.sidebar.checkbox("Use fairness target weekly shifts per nurse", value=False)
target_week = None
if use_fairness:
    target_week = st.sidebar.number_input("Target weekly shifts per nurse", min_value=0, max_value=14, value=4)

weights = Weights(
    w_pref=int(w_pref),
    w_understaff=int(w_under),
    w_overstaff=int(w_over),
    w_fairness=int(w_fair),
    w_weekend=int(w_weekend),
)

st.markdown("## Input data (editable)")

c1, c2 = st.columns(2)
with c1:
    st.caption("Nurses")
    nurses_edit = st.data_editor(nurses, use_container_width=True, num_rows="dynamic")
with c2:
    st.caption("Rooms")
    rooms_edit = st.data_editor(rooms, use_container_width=True, num_rows="dynamic")

c3, c4 = st.columns(2)
with c3:
    st.caption("Demand (day, room, shift → required_nurses)")
    demand_edit = st.data_editor(demand, use_container_width=True, num_rows="dynamic")
with c4:
    st.caption("Preferences (day, shift per nurse; missing rows default to 0)")
    prefs_edit = st.data_editor(prefs, use_container_width=True, num_rows="dynamic")

st.markdown("## Locks (pin assignments)")
st.caption("Set locked=1 to force that nurse to be assigned to that day/shift/room. Solver will error if it’s infeasible.")
locks_edit = st.data_editor(locks, use_container_width=True, num_rows="dynamic")

st.divider()

if st.button("🚀 Solve weekly schedule", type="primary"):
    if not days or not shifts:
        st.error("Select at least one day and one shift.")
    else:
        schedule_df, meta = solve_schedule_ortools(
            nurses_df=pd.DataFrame(nurses_edit),
            rooms_df=pd.DataFrame(rooms_edit),
            demand_df=pd.DataFrame(demand_edit),
            pref_df=pd.DataFrame(prefs_edit),
            locks_df=pd.DataFrame(locks_edit),
            days=days,
            shifts=shifts,
            allow_overstaff=allow_overstaff,
            weights=weights,
            time_limit_seconds=int(time_limit),
            target_shifts_per_nurse_week=(int(target_week) if target_week is not None else None),
            enforce_rest_night_to_day=enforce_rest,
            charge_rooms_tags=frozenset({"ICU","ER"}) if charge_req else frozenset(),
            require_cnor_in_or=cnor_req,
        )

        st.success(f"Solved: {meta['status']} | Objective: {meta['objective']}")
        st.markdown("## Result schedule (editable)")
        st.caption("You can manually edit assigned_nurses and export. (Edits are not re-validated unless you re-solve.)")

        schedule_edit = st.data_editor(
            schedule_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "assigned_nurses": st.column_config.TextColumn(
                    help="Semicolon-separated nurse IDs (e.g., N02;N23)"
                )
            }
        )

        # quick gap check
        tmp = pd.DataFrame(schedule_edit).copy()
        tmp["assigned_count"] = tmp["assigned_nurses"].fillna("").apply(
            lambda s: 0 if str(s).strip()=="" else len([x for x in str(s).split(";") if x.strip()])
        )
        tmp["gap"] = tmp["required_nurses"] - tmp["assigned_count"]

        st.markdown("## Quick coverage check")
        st.dataframe(tmp[["day","room_id","room_name","shift","required_nurses","assigned_count","gap","understaff"] + (["overstaff"] if "overstaff" in tmp.columns else [])],
                     use_container_width=True)

        st.markdown("## Downloads")
        d1, d2, d3, d4 = st.columns(4)
        with d1:
            st.download_button("⬇ schedule.csv", data=csv_bytes(pd.DataFrame(schedule_edit)), file_name="schedule.csv", mime="text/csv")
        with d2:
            st.download_button("⬇ nurses.csv", data=csv_bytes(pd.DataFrame(nurses_edit)), file_name="nurses.csv", mime="text/csv")
        with d3:
            st.download_button("⬇ demand.csv", data=csv_bytes(pd.DataFrame(demand_edit)), file_name="demand.csv", mime="text/csv")
        with d4:
            st.download_button("⬇ locks.csv", data=csv_bytes(pd.DataFrame(locks_edit)), file_name="locks.csv", mime="text/csv")
