import streamlit as st
import pandas as pd
import re, tempfile, shutil
from datetime import datetime, time, timedelta
from io import BytesIO
from openpyxl.utils import get_column_letter

# ============================================================
#                  UTILITY HELPER FUNCTIONS
# ============================================================

def clean_id(x):
    s = str(x).strip().upper()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def to_time(v):
    if pd.isna(v):
        return None
    s = re.sub(r"[^0-9:]", "", str(v))
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except:
            pass
    return None

def find_emp_col(df):
    for c in df.columns:
        if any(k in str(c).lower() for k in 
               ["pay code","emp code","empid","emp id","employee","code"]):
            return c
    return df.columns[0]

def dedupe(df):
    df.columns = [str(c) for c in df.columns]
    return df.loc[:, ~pd.Index(df.columns).duplicated()]

def get_punch_times(df, emp, idcol):
    sub = df[df[idcol] == emp]
    if sub.empty:
        return []
    r = sub.iloc[0]
    punches = []
    for c in df.columns:
        if "PUNCH" in c.upper() and pd.notna(r[c]):
            t = to_time(r[c])
            if t:
                punches.append(t)
    punches.sort()
    return punches

def fmt_hhmm(m):
    try:
        m = int(round(m))
        return f"{m//60:02d}:{m%60:02d}"
    except:
        return ""

# ============================================================
#                        SHIFT WINDOWS
# ============================================================

SHIFT = {
    "day": {
        "in_start": time(6,45), "in_end": time(7,30),
        "out_start": time(15,0), "out_end": time(15,45)
    },
    "hn": {
        "in_start": time(14,45), "in_end": time(15,30),
        "out_start": time(23,0), "out_end": time(23,45)
    },
    "fn": {
        "in_start": time(23,0), "in_end": time(23,30),
        "out_start": time(6,45), "out_end": time(7,30)
    },
    "general1": {
        "in_start": time(7,30), "in_end": time(8,15),
        "out_start": time(15,30), "out_end": time(16,15)
    },
    "general2": {
        "in_start": time(8,30), "in_end": time(9,15),
        "out_start": time(16,30), "out_end": time(17,45)
    },
}

# ============================================================
#                     MAIN COMPARISON LOGIC
# ============================================================

def compare_files(att_bytes, bio1_bytes, bio2_bytes=None):

    # copy fix
    def tmpcopy(b):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(b); f.flush()
            path = f.name
        new = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        shutil.copy(path, new)
        return new

    # ================= READ ATTENDANCE =================
    att_path = tmpcopy(att_bytes)
    xls = pd.ExcelFile(att_path)

    sheet = next((s for s in xls.sheet_names 
                  if "data" in s.lower() and "entry" in s.lower()),
                 xls.sheet_names[0])

    att = pd.read_excel(att_path, sheet_name=sheet)
    att.columns = att.columns.str.strip()

    date_col = next((c for c in att.columns if "date" in c.lower()), None)
    if date_col:
        att[date_col] = pd.to_datetime(att[date_col], errors="coerce")

    # ================= READ BIOMETRIC FILE 1 =================
    bio1 = dedupe(pd.read_excel(tmpcopy(bio1_bytes), skiprows=1))
    bio1.columns = bio1.columns.str.strip()

    # ================= READ BIOMETRIC FILE 2 =================
    bio2 = None
    if bio2_bytes:
        bio2 = dedupe(pd.read_excel(tmpcopy(bio2_bytes), skiprows=1))
        bio2.columns = bio2.columns.str.strip()

    # ================= NORMALIZE EMPLOYEE IDs =================
    emp_att = next((c for c in att.columns if "emp id" in c.lower()), "EMP ID")
    emp_b1 = find_emp_col(bio1)

    att[emp_att] = att[emp_att].astype(str).map(clean_id)
    bio1[emp_b1] = bio1[emp_b1].astype(str).map(clean_id)

    if bio2 is not None:
        emp_b2 = find_emp_col(bio2)
        bio2[emp_b2] = bio2[emp_b2].astype(str).map(clean_id)
    else:
        emp_b2 = None

    emps_b1 = set(bio1[emp_b1])
    emps_b2 = set(bio2[emp_b2]) if bio2 is not None else set()

    # ================= EXTRACT OT =================
    def get_ot(row):
        for c in att.columns:
            if "overtime" in c.lower():
                v = row.get(c)
                if pd.notna(v):
                    try:
                        return float(v)
                    except:
                        return 0.0
        return 0.0

    # ================= DURATION =================
    def duration(in_p, out_p):
        if in_p is None or out_p is None:
            return None
        # cross midnight
        out_adj = out_p
        if out_p < in_p:
            out_adj = out_p + timedelta(days=1)
        return (out_adj - in_p).total_seconds() / 60

    # ================= PICK SAME DAY =================
    def pick_same_day(punches, s):
        if not punches:
            return None, None
        # IN punch
        in_candidates = [p for p in punches if s["in_start"] <= p.time() <= s["in_end"]]
        if in_candidates:
            in_p = max(in_candidates)
        else:
            before = [p for p in punches if p.time() < s["out_start"]]
            in_p = max(before) if before else punches[0]
        # OUT punch = last punch
        out_p = punches[-1]
        return in_p, out_p

    # =====================================================
    #                PROCESS EMPLOYEES
    # =====================================================

    statuses = []
    in_list = []
    out_list = []
    hours_list = []
    remark_list = []

    shiftcol = next((c for c in att.columns if "shift" in c.lower()), "SHIFT")

    for _, row in att.iterrows():

        emp = clean_id(row[emp_att])
        shift = str(row.get(shiftcol, "")).lower()

        # Select shift
        if "day" in shift:
            s = SHIFT["day"]
        elif "general1" in shift:
            s = SHIFT["general1"]
        elif "general2" in shift:
            s = SHIFT["general2"]
        elif any(x in shift for x in ["hn","half","night"]):
            s = SHIFT["hn"]
        else:
            s = SHIFT["fn"]

        # gather punches
        punches = []
        if emp in emps_b1:
            punches += get_punch_times(bio1, emp, emp_b1)
        if bio2 is not None and emp in emps_b2:
            punches += get_punch_times(bio2, emp, emp_b2)

        punches = sorted(punches)
        ot_val = get_ot(row)

        # defaults
        status = "No Punch"
        remark = ""
        in_str = ""
        out_str = ""
        whr = ""

        # ========== No Punch ==========
        if not punches:
            statuses.append(status); in_list.append(""); out_list.append(""); hours_list.append(""); remark_list.append("")
            continue

        # ========== Single Punch ==========
        if len(punches) == 1:
            t = punches[0].time()
            if t < s["out_start"]:
                status = "Single In Punch"
                in_str = punches[0].strftime("%H:%M")
            else:
                status = "Single Out Punch"
                out_str = punches[0].strftime("%H:%M")
            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(""); remark_list.append("")
            continue

        # ========== Multiple Punches ==========
        in_p, out_p = pick_same_day(punches, s)

        in_str = in_p.strftime("%H:%M")
        out_str = out_p.strftime("%H:%M")
        dur = duration(in_p, out_p)
        whr = fmt_hhmm(dur)

        # =====================================================
        # 1) OUT PUNCH LOGIC (highest priority)
        # =====================================================
        if out_p.time() > s["out_end"]:
            status = "Late Out Punch"
            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        # =====================================================
        # 2) EARLY OUT
        # =====================================================
        if out_p.time() < s["out_start"]:
            status = "Early"
            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        # =====================================================
        # 3) LATE IN
        # =====================================================
        if in_p.time() > s["in_end"]:
            status = "Late In Punch"
            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        # =====================================================
        # 4) OVERTIME LOGIC
        # =====================================================
        if ot_val > 0:
            expected = 480 + (ot_val * 60)
            if dur >= expected:
                status = "Match"
            elif dur >= expected - 30:   # your 30 min rule kept
                status = "Match"
            else:
                status = "OT Deviation"

            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        # =====================================================
        # 5) BELOW 8 HOURS (ONLY OT = 0)
        # =====================================================
        if dur < 480:
            status = "Below 8 Hrs"
            remark = "BELOW 8HRS"
            statuses.append(status); in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append(remark)
            continue

        # =====================================================
        # 6) MATCH
        # =====================================================
        status = "Match"

        statuses.append(status)
        in_list.append(in_str)
        out_list.append(out_str)
        hours_list.append(whr)
        remark_list.append(remark)

    # =====================================================
    # OUTPUT DF
    # =====================================================

    att["Status"] = statuses
    att["In Time"] = in_list
    att["Out Time"] = out_list
    att["Work Hrs"] = hours_list
    att["Remark"] = remark_list

    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        att.to_excel(w, index=False, sheet_name="Result")
    out.seek(0)

    return out, att


# ============================================================
#                          STREAMLIT UI
# ============================================================

st.title("🕒 FINAL Attendance Comparator — OT + Late Out + Below 8 Hrs Logic")

att = st.file_uploader("📁 Attendance File", type=["xlsx"])
bio1 = st.file_uploader("📁 Biometric Day 1", type=["xlsx"])
bio2 = st.file_uploader("📁 Biometric Day 2 (Optional)", type=["xlsx"])

if st.button("🔍 Compare"):
    if not att or not bio1:
        st.error("⚠️ Upload attendance + biometric day 1")
    else:
        out, df = compare_files(att.read(), bio1.read(), bio2.read() if bio2 else None)
        st.success("Completed!")
        st.download_button("⬇ Download Result", out, "Attendance_with_Status.xlsx")
        st.dataframe(df)
