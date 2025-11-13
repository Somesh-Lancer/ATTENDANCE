import streamlit as st
import pandas as pd
import re, tempfile, shutil
from datetime import datetime, time, timedelta
from io import BytesIO
from openpyxl.utils import get_column_letter

# ---------- Utility Helpers ----------
def clean_id(x):
    s = str(x).strip().upper()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s

def to_time(v):
    if pd.isna(v): return None
    s = re.sub(r"[^0-9:]", "", str(v))
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).replace(second=0, microsecond=0)
        except:
            pass
    return None

def find_emp_col(df):
    for c in df.columns:
        if any(k in str(c).lower() for k in ["pay code","emp code","empid","emp id","employee code","code"]):
            return c
    return df.columns[0]

def dedupe(df):
    df.columns = [str(c) for c in df.columns]
    return df.loc[:, ~pd.Index(df.columns).duplicated(keep="first")]

def get_punch_times(df, emp, id_col):
    sub = df[df[id_col]==emp]
    if sub.empty: return []
    r = sub.iloc[0]
    punches=[]
    for c in df.columns:
        if "PUNCH" in str(c).upper() and pd.notna(r[c]):
            t=to_time(r[c])
            if t: punches.append(t)
    punches.sort()
    return punches

def fmt_hhmm(minutes):
    try:
        minutes = int(round(minutes))
        h = minutes // 60
        m = minutes % 60
        return f"{h:02d}:{m:02d}"
    except:
        return ""

# ---------- Shift Windows ----------
SHIFT = {
    "day":       {"in_start": time(6,45), "in_end": time(7,30),
                  "out_start": time(15,0), "out_end": time(15,45)},
    "hn":        {"in_start": time(14,45), "in_end": time(15,30),
                  "out_start": time(23,0),  "out_end": time(23,45)},
    "fn":        {"in_start": time(23,0),  "in_end": time(23,30),
                  "out_start": time(6,45),  "out_end": time(7,30)},
    "general1":  {"in_start": time(7,30),  "in_end": time(8,15),
                  "out_start": time(15,30), "out_end": time(16,15)},
    "general2":  {"in_start": time(8,30),  "in_end": time(9,15),
                  "out_start": time(16,30), "out_end": time(17,45)},
}

# ---------- Core ----------
def compare_files(att_bytes,bio1_bytes,bio2_bytes=None):
    def tmpcopy(b):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx") as f:
            f.write(b); f.flush(); path=f.name
        tmp=tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx").name
        shutil.copy(path,tmp); return tmp

    # Attendance
    att_path = tmpcopy(att_bytes)
    xls=pd.ExcelFile(att_path)
    sheet=next((s for s in xls.sheet_names if "data" in s.lower() and "entry" in s.lower()), xls.sheet_names[0])
    att_df=pd.read_excel(att_path,sheet_name=sheet)
    att_df.columns=att_df.columns.str.strip()
    date_col=next((c for c in att_df.columns if "date" in c.lower()),None)
    if date_col: att_df[date_col]=pd.to_datetime(att_df[date_col],errors="coerce")

    # Biometric files
    bio1_path = tmpcopy(bio1_bytes)
    bio1 = dedupe(pd.read_excel(bio1_path, skiprows=1))
    bio1.columns = bio1.columns.str.strip()
    bio2=None
    if bio2_bytes:
        bio2_path = tmpcopy(bio2_bytes)
        bio2 = dedupe(pd.read_excel(bio2_path, skiprows=1))
        bio2.columns = bio2.columns.str.strip()

    # Normalize IDs
    m_emp=next((c for c in att_df.columns if "emp id" in c.lower()),"EMP ID")
    b_emp1=find_emp_col(bio1)
    if bio2 is not None: b_emp2=find_emp_col(bio2)
    att_df[m_emp]=att_df[m_emp].astype(str).map(clean_id)
    bio1[b_emp1]=bio1[b_emp1].astype(str).map(clean_id)
    if bio2 is not None: bio2[b_emp2]=bio2[b_emp2].astype(str).map(clean_id)

    emps_b1=set(bio1[b_emp1])
    emps_b2=set(bio2[b_emp2]) if bio2 is not None else set()
    all_emps = emps_b1 | emps_b2  # keep all attendance rows; we’ll mark No Punch if missing

    # ---- helpers ----
    def get_ot_hours(row):
        for c in att_df.columns:
            if "overtime" in c.lower():
                v=row.get(c)
                if pd.notna(v):
                    try: return float(v)
                    except: return None
        return None

    def pick_in_out_same_day(punches, s):
        """
        IN  = latest punch inside IN window; if none, latest punch BEFORE OUT-start; else last punch
        OUT = last punch at/after OUT-start (not restricted to OUT window, so we don't miss OT)
        Returns (in_p, out_p)
        """
        if not punches:
            return None, None
        in_win = [p for p in punches if s["in_start"] <= p.time() <= s["in_end"]]
        if in_win:
            in_p = max(in_win)  # latest IN inside window
        else:
            before_out = [p for p in punches if p.time() < s["out_start"]]
            in_p = max(before_out) if before_out else punches[-1]

        after_out = [p for p in punches if p.time() >= s["out_start"]]
        out_p = max(after_out) if after_out else None
        return in_p, out_p

    def compute_duration_minutes(in_p, out_p):
        if not in_p or not out_p:
            return None
        end = out_p
        if out_p < in_p:
            end = out_p + timedelta(days=1)
        return (end - in_p).total_seconds() / 60

    results=[]
    in_times=[]
    out_times=[]
    work_hrs=[]
    remark=[]
    shift_col=next((c for c in att_df.columns if "shift" in c.lower()),"SHIFT")

    for _,row in att_df.iterrows():
        emp=row[m_emp]
        shift=str(row.get(shift_col,"")).lower().replace("-","").replace(" ","")
        p1=get_punch_times(bio1,emp,b_emp1) if emp in emps_b1 else []
        p2=get_punch_times(bio2,emp,b_emp2) if bio2 is not None and emp in emps_b2 else []
        status="No Punch"
        in_time_str=""
        out_time_str=""
        whr_str=""
        rem_str=""
        ot_val=get_ot_hours(row)

        # If employee has no biometric punches at all
        if emp not in all_emps:
            results.append("No Punch"); in_times.append(""); out_times.append(""); work_hrs.append(""); remark.append("")
            continue

        # ---- DAY / GENERAL-1 / GENERAL-2 use same-day logic with robust OUT pick ----
        def handle_same_day(punches, s):
            nonlocal status, in_time_str, out_time_str, whr_str, rem_str
            if not punches:
                status="No Punch"; return
            # If only in-window cluster and no after-out punches → Single In Punch but still show last IN
            in_p, out_p = pick_in_out_same_day(punches, s)
            if not out_p and in_p:
                status="Single In Punch"; in_time_str=in_p.strftime("%H:%M"); out_time_str=""
                return
            if not in_p and out_p:
                status="Single Out Punch"; in_time_str=""; out_time_str=out_p.strftime("%H:%M")
                return
            if not in_p and not out_p:
                status="No Punch"; return

            # Compute duration & display
            in_time_str = in_p.strftime("%H:%M")
            out_time_str = out_p.strftime("%H:%M")
            dur_min = compute_duration_minutes(in_p, out_p)
            whr_str = fmt_hhmm(dur_min) if dur_min is not None else ""

            # If OT is provided, use OT rule (±30 min) and don't use Below 8 Hrs
            if ot_val and ot_val > 0:
                expected_min = ot_val*60
                status = "Match" if abs(dur_min-expected_min) <= 30 else "OT Deviation"
                return

            # Below 8 hours rule (independent of windows)
            if dur_min is not None and dur_min < 480:
                status = "Below 8 Hrs"
                rem_str = "BELOW 8HRS"
                return

            # If not OT and not <8h, then evaluate windows for status
            t_in, t_out = in_p.time(), out_p.time()
            if (s["in_start"]<=t_in<=s["in_end"]) and (s["out_start"]<=t_out<=s["out_end"]):
                status="Match"
            elif t_in > s["in_end"]:
                status="Late In Punch"
            elif t_out < s["out_start"]:
                status="Early"
            else:
                status="Mismatch"

        # ----- DAY -----
        if "day" in shift:
            handle_same_day(p1, SHIFT["day"])

        # ----- GENERAL-1 -----
        elif any(x in shift for x in ["general1","general 1"]):
            handle_same_day(p1, SHIFT["general1"])

        # ----- GENERAL-2 -----
        elif any(x in shift for x in ["general2","general 2"]):
            handle_same_day(p1, SHIFT["general2"])

        # ----- HALF-NIGHT -----
        elif any(x in shift for x in ["hf","half","hnight","hn"]):
            s=SHIFT["hn"]
            punches=(p1 or [])+(p2 or [])
            if not punches:
                status="No Punch"
            else:
                # For HN we’ll still compute duration when both exist
                # IN = latest inside IN window or last before out-start; OUT = last >= out-start
                in_p, out_p = pick_in_out_same_day(punches, s)
                if not out_p and in_p:
                    status="Single In Punch"; in_time_str=in_p.strftime("%H:%M")
                elif not in_p and out_p:
                    status="Single Out Punch"; out_time_str=out_p.strftime("%H:%M")
                elif not in_p and not out_p:
                    status="No Punch"
                else:
                    in_time_str=in_p.strftime("%H:%M"); out_time_str=out_p.strftime("%H:%M")
                    dur_min = compute_duration_minutes(in_p, out_p)
                    whr_str = fmt_hhmm(dur_min) if dur_min is not None else ""
                    if ot_val and ot_val>0:
                        expected_min=ot_val*60
                        status="Match" if abs(dur_min-expected_min)<=30 else "OT Deviation"
                    else:
                        if dur_min is not None and dur_min < 480:
                            status="Below 8 Hrs"; rem_str="BELOW 8HRS"
                        else:
                            t_in,t_out=in_p.time(),out_p.time()
                            if (s["in_start"]<=t_in<=s["in_end"]) and (s["out_start"]<=t_out<=s["out_end"]):
                                status="Match"
                            elif t_in>s["in_end"]:
                                status="Late In Punch"
                            elif t_out<s["out_start"]:
                                status="Early"
                            else:
                                status="Mismatch"

        # ----- FULL-NIGHT (cross-day) -----
        elif any(x in shift for x in ["fn","fullnight","night"]):
            s=SHIFT["fn"]
            # Prefer punches inside windows
            in_win=[p for p in p1 if s["in_start"]<=p.time()<=s["in_end"]]
            out_win=[p for p in p2 if s["out_start"]<=p.time()<=s["out_end"]]
            in_p = max(in_win) if in_win else (p1[-1] if p1 else None)
            out_p = max(out_win) if out_win else (p2[-1] if p2 else None)

            if not in_p and not out_p:
                status="No Punch"
            elif in_p and not out_p:
                status="Single In Punch"; in_time_str=in_p.strftime("%H:%M")
            elif not in_p and out_p:
                status="Single Out Punch"; out_time_str=out_p.strftime("%H:%M")
            else:
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                in_time_str=in_p.strftime("%H:%M"); out_time_str=out_p.strftime("%H:%M")
                dur_min = compute_duration_minutes(in_p, out_p)
                whr_str = fmt_hhmm(dur_min) if dur_min is not None else ""
                if ot_val and ot_val>0:
                    expected_min=ot_val*60
                    status="Match" if abs(dur_min-expected_min)<=30 else "OT Deviation"
                else:
                    t_in,t_out=in_p.time(),out_p.time()
                    if (s["in_start"]<=t_in<=s["in_end"]) and (s["out_start"]<=t_out<=s["out_end"]):
                        status="Match"
                    elif t_in>s["in_end"]:
                        status="Late In Punch"
                    elif t_out<s["out_start"]:
                        status="Early"
                    else:
                        status="Mismatch"

        else:
            status="No Punch"

        results.append(status)
        in_times.append(in_time_str)
        out_times.append(out_time_str)
        work_hrs.append(whr_str)
        remark.append(rem_str)

    # append columns
    att_df["Status"] = results
    att_df["In Time"] = in_times
    att_df["Out Time"] = out_times
    att_df["Work Hrs"] = work_hrs
    att_df["Remark"] = remark

    # ---------- Output ----------
    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl", datetime_format="yyyy-mm-dd") as w:
        att_df.to_excel(w,index=False,sheet_name=sheet)
        ws=w.sheets[sheet]
        # Date formatting and widen column to avoid ####
        if date_col:
            col_idx = att_df.columns.get_loc(date_col) + 1
            col_letter = get_column_letter(col_idx)
            for r in range(2, len(att_df) + 2):
                ws[f"{col_letter}{r}"].number_format = "dd-mm-yyyy"
            ws.column_dimensions[col_letter].width = 14
        # widen In/Out/Status/Work Hrs/Remark columns a bit
        for name, width in [("Status",14), ("In Time",10), ("Out Time",10), ("Work Hrs",10), ("Remark",14)]:
            if name in att_df.columns:
                idx = att_df.columns.get_loc(name)+1
                ws.column_dimensions[get_column_letter(idx)].width = width
    out.seek(0)
    return out,att_df

# ---------- Streamlit ----------
st.set_page_config(page_title="Attendance Comparator (Fixed Duration + In/Out)",page_icon="🕒",layout="centered")
st.title("🕒 Attendance Comparator — Fixed IN/OUT, Duration & Below-8Hrs")
st.markdown("""
**What’s fixed now**
- Picks **latest** valid IN inside window and **last** punch after OUT-start → Out Time never missed (e.g., 16:21).
- Computes **Work Hrs** = Out − In and shows it as **HH:MM**.
- If Work Hrs **< 8:00**, sets status to **Below 8 Hrs** (and Remark = BELOW 8HRS).
- OT still applies only if Attendance OT > 0 (±30 min tolerance).
- Date column is formatted and widened (no more #####).
""")

att=st.file_uploader("📁 Attendance File",type=["xlsx"])
bio1=st.file_uploader("📁 Biometric File – Day 1 (Required)",type=["xlsx"])
bio2=st.file_uploader("📁 Biometric File – Day 2 (Optional)",type=["xlsx"])

if st.button("🔍 Compare Files"):
    if not att or not bio1:
        st.error("⚠️ Please upload Attendance and at least one Biometric file.")
    else:
        with st.spinner("Processing …"):
            try:
                out,df=compare_files(att.read(),bio1.read(),bio2.read() if bio2 else None)
                st.success("✅ Comparison complete!")
                st.download_button("⬇️ Download Excel",data=out,
                    file_name="Attendance_with_Status.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(df.head(30))
            except Exception as e:
                st.error(f"❌ Error:\n{e}")
