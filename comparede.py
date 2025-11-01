import streamlit as st
import pandas as pd
import numpy as np
import re
import tempfile
import shutil
from datetime import datetime, timedelta, time
from io import BytesIO
from openpyxl.utils import get_column_letter

# ======================================================
# Utility functions
# ======================================================
def clean_id(x):
    s = str(x).strip().upper()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def to_time(v):
    """Convert visible HH:MM or HH:MM:SS to datetime (ignore seconds)."""
    if pd.isna(v):
        return None
    s = str(v)
    s = re.sub(r"[^0-9:]", "", s).strip()
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            t = datetime.strptime(s, fmt)
            return t.replace(second=0, microsecond=0)
        except:
            pass
    return None

def find_emp_col(df):
    possible = ["pay code","emp code","employee code","empid","emp id","code"]
    for name in possible:
        for c in df.columns:
            if name in str(c).lower():
                return c
    return df.columns[0]

def dedupe_columns_inplace(df):
    df.columns = [str(c) for c in df.columns]
    mask = ~pd.Index(df.columns).duplicated(keep="first")
    return df.loc[:, mask]

def get_punch_times(df, emp, id_col):
    """Return sorted punch times for an employee."""
    sub = df[df[id_col] == emp]
    if sub.empty:
        return []
    r = sub.iloc[0]
    punches = []
    for c in df.columns:
        if "PUNCH" in str(c).upper() and pd.notna(r[c]):
            t = to_time(r[c])
            if t:
                punches.append(t)
    punches.sort()
    return punches

# ======================================================
# Shift timing dictionary (base start/end)
# ======================================================
SHIFT_WINDOWS = {
    "day":       {"start": time(7,0),  "end": time(15,15)},
    "hf":        {"start": time(15,15), "end": time(23,15)},
    "fn":        {"start": time(23,15), "end": time(7,15)},
    "general1":  {"start": time(8,0),  "end": time(16,0)},
    "general2":  {"start": time(9,0),  "end": time(17,0)},
}

# ======================================================
# Core comparison logic with ±15 min windows
# ======================================================
def compare_files(att_bytes, bio1_bytes, bio2_bytes):
    # ---------- SAFE COPY ----------
    def safe_copy(input_bytes):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as f:
            f.write(input_bytes)
            f.flush()
            temp_path = f.name
        tmp_copy = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx").name
        shutil.copy(temp_path, tmp_copy)
        return tmp_copy

    att_path = safe_copy(att_bytes)
    bio1_path = safe_copy(bio1_bytes)
    bio2_path = safe_copy(bio2_bytes)

    # ---------- Attendance ----------
    xls = pd.ExcelFile(att_path)
    sheet = next((s for s in xls.sheet_names if "data" in s.lower() and "entry" in s.lower()), xls.sheet_names[0])
    att_df = pd.read_excel(att_path, sheet_name=sheet)
    att_df.columns = att_df.columns.str.strip()
    date_col = next((c for c in att_df.columns if "date" in c.lower()), None)
    if date_col:
        att_df[date_col] = pd.to_datetime(att_df[date_col], errors="coerce")

    # ---------- Biometric ----------
    bio1 = pd.read_excel(bio1_path, skiprows=1)
    bio2 = pd.read_excel(bio2_path, skiprows=1)
    bio1, bio2 = dedupe_columns_inplace(bio1), dedupe_columns_inplace(bio2)
    bio1.columns, bio2.columns = bio1.columns.str.strip(), bio2.columns.str.strip()

    m_emp = next((c for c in att_df.columns if "emp id" in c.lower()), "EMP ID")
    b_emp1, b_emp2 = find_emp_col(bio1), find_emp_col(bio2)
    att_df[m_emp] = att_df[m_emp].astype(str).map(clean_id)
    bio1[b_emp1] = bio1[b_emp1].astype(str).map(clean_id)
    bio2[b_emp2] = bio2[b_emp2].astype(str).map(clean_id)

    statuses = []
    shift_col = next((c for c in att_df.columns if "shift" in c.lower()), "SHIFT")

    for _, row in att_df.iterrows():
        emp = row[m_emp]
        shift = str(row.get(shift_col, "")).lower().strip()
        p_day1 = get_punch_times(bio1, emp, b_emp1)
        p_day2 = get_punch_times(bio2, emp, b_emp2)
        status = "No Punch"

        # ========== Full Night ==========
        if "full" in shift or "fn" in shift:
            start = SHIFT_WINDOWS["fn"]["start"]
            end = SHIFT_WINDOWS["fn"]["end"]
            in_punch = p_day1[0] if len(p_day1) > 0 else None
            out_punch = p_day2[0] if len(p_day2) > 0 else None
            if in_punch and out_punch:
                if out_punch <= in_punch:
                    out_punch = out_punch + timedelta(days=1)
                # Apply ±15 min window rules
                in_valid = (datetime.combine(datetime.today(), start) - timedelta(minutes=15)).time() <= in_punch.time() <= (datetime.combine(datetime.today(), start) + timedelta(minutes=15)).time()
                out_valid = (datetime.combine(datetime.today(), end) - timedelta(minutes=15)).time() <= out_punch.time() <= (datetime.combine(datetime.today(), end) + timedelta(minutes=15)).time()
                if not in_valid or not out_valid:
                    if out_punch.time() < end:
                        status = "Early"
                    else:
                        status = "Match"
                else:
                    if out_punch.time() < end:
                        status = "Early"
                    else:
                        status = "Match"
            elif in_punch and not out_punch:
                status = "Single In Punch"
            elif not in_punch and out_punch:
                status = "Single Out Punch"
            else:
                status = "No Punch"

        # ========== Other Shifts ==========
        else:
            s_key = None
            for k in SHIFT_WINDOWS:
                if k in shift:
                    s_key = k
                    break

            if not s_key:
                if len(p_day1) == 0:
                    status = "No Punch"
                elif len(p_day1) == 1:
                    status = "Single In Punch"
                else:
                    status = "Match"
            else:
                base_start = SHIFT_WINDOWS[s_key]["start"]
                base_end = SHIFT_WINDOWS[s_key]["end"]

                # Compute ±15 min windows
                in_start = (datetime.combine(datetime.today(), base_start) - timedelta(minutes=15)).time()
                in_end = (datetime.combine(datetime.today(), base_start) + timedelta(minutes=15)).time()
                out_start = (datetime.combine(datetime.today(), base_end) - timedelta(minutes=15)).time()
                out_end = (datetime.combine(datetime.today(), base_end) + timedelta(minutes=15)).time()

                # Determine IN & OUT based on punch count
                if len(p_day1) >= 4:
                    in_p, out_p = p_day1[0], p_day1[3]
                elif len(p_day1) == 3:
                    in_p, out_p = p_day1[0], p_day1[2]
                elif len(p_day1) == 2:
                    in_p, out_p = p_day1[0], p_day1[1]
                elif len(p_day1) == 1:
                    t = p_day1[0].time()
                    if out_start <= t <= out_end:
                        status = "Single Out Punch"
                    else:
                        status = "Single In Punch"
                    statuses.append(status)
                    continue
                else:
                    status = "No Punch"
                    statuses.append(status)
                    continue

                # Evaluate based on windows
                if not in_p or not out_p:
                    status = "No Punch"
                else:
                    if in_start <= in_p.time() <= in_end:
                        # IN is valid; check OUT
                        if out_start <= out_p.time() <= out_end:
                            status = "Match"
                        elif out_p.time() < base_end:
                            status = "Early"
                        elif out_p.time() > out_end:
                            status = "Match"
                        else:
                            status = "Match"
                    else:
                        # IN outside valid window, still check OUT
                        if out_p.time() < base_end:
                            status = "Early"
                        else:
                            status = "Match"

        statuses.append(status)

    att_df["Status"] = statuses

    # ---------- Output ----------
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        att_df.to_excel(writer, index=False, sheet_name=sheet)
        ws = writer.sheets[sheet]
        if date_col:
            col_idx = att_df.columns.get_loc(date_col) + 1
            col_letter = get_column_letter(col_idx)
            for r in range(2, len(att_df) + 2):
                ws[f"{col_letter}{r}"].number_format = "yyyy-mm-dd"
            ws.column_dimensions[col_letter].width = 15
    output.seek(0)
    return output, att_df

# ======================================================
# Streamlit UI
# ======================================================
st.set_page_config(page_title="Attendance Comparator (All Shifts)", page_icon="🕒", layout="centered")
st.title("🕒 Attendance Comparator — ±15 min IN/OUT Logic for All Shifts")
st.markdown("""
✅ **IN Punch** valid if within *15 min before or after* shift start.  
✅ **OUT Punch** valid if within *15 min before or after* shift end.  
✅ **OUT before shift end** → *Early*.  
✅ Handles all shifts + Full-Night across two files.  
✅ Works even if Excel files are open.  
""")

att = st.file_uploader("📁 Upload Attendance File", type=["xlsx"])
bio1 = st.file_uploader("📁 Upload Biometric File – Day 1", type=["xlsx"])
bio2 = st.file_uploader("📁 Upload Biometric File – Day 2", type=["xlsx"])

if st.button("🔍 Compare Files"):
    if not att or not bio1 or not bio2:
        st.error("⚠️ Please upload all three files.")
    else:
        with st.spinner("Processing..."):
            try:
                output, df = compare_files(att.read(), bio1.read(), bio2.read())
            except Exception as e:
                st.error(f"❌ Error:\n{e}")
            else:
                st.success("✅ Comparison complete!")
                st.download_button(
                    "⬇️ Download Excel",
                    output,
                    "Attendance_with_Status.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.dataframe(df.head(20))
