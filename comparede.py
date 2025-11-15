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

# FIXED TIME PARSER
def to_time(v):
    if pd.isna(v):
        return None

    s = str(v).strip()

    # Extract only HH:MM or HH:MM:SS
    match = re.search(r"(\d{1,2}:\d{2}(?::\d{2})?)", s)
    if not match:
        return None

    s = match.group(1)

    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt)
        except:
            pass
    return None


def find_emp_col(df):
    for c in df.columns:
        if any(k in str(c).lower() for k in
               ["pay code", "emp code", "empid", "emp id", "employee", "code"]):
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
        if "PUNCH" in c.upper() or "TIME" in c.upper():
            if pd.notna(r[c]):
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

# Fixed OUT cutoff times per shift
CUT = {
    "day": time(15,15),
    "hn": time(23,15),
    "fn": time(7,15),
    "general1": time(16,0),
    "general2": time(17,0)
}

# ============================================================
#                     MAIN COMPARISON LOGIC
# ============================================================

def compare_files(att_bytes, bio1_bytes, bio2_bytes=None):

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
        out_adj = out_p
        if out_p < in_p:
            out_adj = out_p + timedelta(days=1)
        return (out_adj - in_p).total_seconds() / 60


    # =====================================================
    # PROCESS EACH EMPLOYEE
    # =====================================================

    statuses, in_list, out_list, hours_list, remark_list = [], [], [], [], []
    shiftcol = next((c for c in att.columns if "shift" in c.lower()), "SHIFT")

    for _, row in att.iterrows():

        emp = clean_id(row[emp_att])
        shift = str(row.get(shiftcol, "")).lower()

        # FIX SHIFT DETECTION (Hyphens accepted)
        if "day" in shift:
            sk = "day"
        elif "general1" in shift or "general-1" in shift:
            sk = "general1"
        elif "general2" in shift or "general-2" in shift:
            sk = "general2"
        elif any(x in shift for x in ["hn", "half"]):
            sk = "hn"
        elif "night" in shift or "fn" in shift:
            sk = "fn"
        else:
            sk = "fn"

        cutoff = CUT[sk]

        # =========================================================
        # SPECIAL FULL NIGHT SHIFT LOGIC (FN)
        # =========================================================

        if sk == "fn":

            punches_1 = get_punch_times(bio1, emp, emp_b1) if emp in emps_b1 else []
            punches_2 = get_punch_times(bio2, emp, emp_b2) if (bio2 is not None and emp in emps_b2) else []

            if not punches_1 and not punches_2:
                statuses.append("No Punch"); in_list.append(""); out_list.append(""); hours_list.append(""); remark_list.append("")
                continue

            in_p = punches_1[-1] if punches_1 else None
            out_p = punches_2[0] if punches_2 else None

            in_str = in_p.strftime("%H:%M") if in_p else ""
            out_str = out_p.strftime("%H:%M") if out_p else ""

            # Single punches
            if in_p and not out_p:
                statuses.append("Single In Punch"); in_list.append(in_str); out_list.append(""); hours_list.append(""); remark_list.append("")
                continue

            if out_p and not in_p:
                statuses.append("Single Out Punch"); in_list.append(""); out_list.append(out_str); hours_list.append(""); remark_list.append("")
                continue

            # Early out
            if out_p.time() < cutoff:
                dur = duration(in_p, out_p)
                statuses.append("Early Out Punch")
                in_list.append(in_str); out_list.append(out_str); hours_list.append(fmt_hhmm(dur)); remark_list.append("")
                continue

            dur = duration(in_p, out_p)
            whr = fmt_hhmm(dur)
            ot_val = get_ot(row)

            # OT logic
            if ot_val > 0:
                expected = 480 + ot_val*60
                if dur >= expected or dur >= expected - 30:
                    status = "Match"
                else:
                    status = "OT Deviation"
            else:
                if dur < 480:
                    statuses.append("Below 8 Hrs")
                    in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("BELOW 8HRS")
                    continue
                status = "Match"

            statuses.append(status)
            in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue


        # =========================================================
        # OTHER SHIFTS
        # =========================================================

        punches = []
        if emp in emps_b1:
            punches += get_punch_times(bio1, emp, emp_b1)
        if bio2 is not None and emp in emps_b2:
            punches += get_punch_times(bio2, emp, emp_b2)

        punches = sorted(punches)
        ot_val = get_ot(row)

        if not punches:
            statuses.append("No Punch"); in_list.append(""); out_list.append(""); hours_list.append(""); remark_list.append("")
            continue

        # Single punch
        if len(punches) == 1:
            t = punches[0].time()
            if t >= cutoff:
                statuses.append("Single Out Punch")
                in_list.append(""); out_list.append(punches[0].strftime("%H:%M")); hours_list.append(""); remark_list.append("")
            else:
                statuses.append("Single In Punch")
                in_list.append(punches[0].strftime("%H:%M")); out_list.append(""); hours_list.append(""); remark_list.append("")
            continue

        # Multi-punch logic
        s = SHIFT[sk]
        out_p = punches[-1]

        # Correct IN punch selection
        in_candidates = [p for p in punches if s["in_start"] <= p.time() <= s["in_end"]]

        if in_candidates:
            in_p = min(in_candidates)  # Earliest within IN window
        else:
            before = [p for p in punches if p.time() < s["out_start"]]
            if before:
                in_p = min(before)
            else:
                in_p = punches[0]

        in_str = in_p.strftime("%H:%M")
        out_str = out_p.strftime("%H:%M")
        dur = duration(in_p, out_p)
        whr = fmt_hhmm(dur)

        # Early out
        if out_p.time() < cutoff:
            statuses.append("Early Out Punch")
            in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        # OT logic
        if ot_val > 0:
            expected = 480 + ot_val*60
            if dur >= expected or dur >= expected - 30:
               status = "Match"
            else:
               status = "OT Deviation"

            statuses.append(status)
            in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")
            continue

        if dur < 480:
            statuses.append("Below 8 Hrs")
            in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("BELOW 8HRS")
            continue

        statuses.append("Match")
        in_list.append(in_str); out_list.append(out_str); hours_list.append(whr); remark_list.append("")

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

st.title("🕒 FINAL Attendance Comparator — Updated Fixed Logic")

att = st.file_uploader("📁 Attendance File", type=["xlsx"])
bio1 = st.file_uploader("📁 Biometric Day 1", type=["xlsx"])
bio2 = st.file_uploader("📁 Biometric Day 2", type=["xlsx"])

if st.button("🔍 Compare"):
    if not att or not bio1 or not bio2:
        st.error("⚠️ Upload attendance + both biometric files")
    else:
        out, df = compare_files(att.read(), bio1.read(), bio2.read())
        st.success("Completed!")
        st.download_button("⬇ Download Result", out, "Attendance_with_Status.xlsx")
        st.dataframe(df)
