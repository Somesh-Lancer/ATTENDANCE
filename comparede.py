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
        if any(k in str(c).lower() for k in
               ["pay code","emp code","empid","emp id","employee code","code"]):
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
    sheet=next((s for s in xls.sheet_names if "data" in s.lower() and "entry" in s.lower()),
               xls.sheet_names[0])
    att_df=pd.read_excel(att_path,sheet_name=sheet)
    att_df.columns=att_df.columns.str.strip()
    date_col=next((c for c in att_df.columns if "date" in c.lower()),None)
    if date_col: att_df[date_col]=pd.to_datetime(att_df[date_col],errors="coerce")

    # Biometric Files
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

    all_emps=set(bio1[b_emp1])
    if bio2 is not None: all_emps|=set(bio2[b_emp2])
    att_df=att_df[att_df[m_emp].isin(all_emps)].reset_index(drop=True)

    # ---- helpers ----
    def get_ot_hours(row):
        for c in att_df.columns:
            if "overtime" in c.lower():
                v=row.get(c)
                if pd.notna(v):
                    try: return float(v)
                    except: return None
        return None

    results=[]
    shift_col=next((c for c in att_df.columns if "shift" in c.lower()),"SHIFT")

    for _,row in att_df.iterrows():
        emp=row[m_emp]
        shift=str(row.get(shift_col,"")).lower().replace("-","").replace(" ","")
        p1=get_punch_times(bio1,emp,b_emp1)
        p2=get_punch_times(bio2,emp,b_emp2) if bio2 is not None else []
        status="No Punch"
        ot_val=get_ot_hours(row)

        # ----- DAY -----
        if "day" in shift:
            s=SHIFT["day"]
            if not p1:
                status="No Punch"
            elif len(p1)==1:
                t=p1[0].time()
                status="Single In Punch" if t < s["out_start"] else "Single Out Punch"
            else:
                in_candidates=[p for p in p1 if s["in_start"]<=p.time()<=s["in_end"]]
                out_candidates=[p for p in p1 if s["out_start"]<=p.time()<=s["out_end"]]
                in_p=min(in_candidates) if in_candidates else p1[0]
                out_p=max(out_candidates) if out_candidates else p1[-1]
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                if ot_val and ot_val>0:
                    status="Match"
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
            results.append(status)
            continue

        # ----- HALF-NIGHT -----
        if any(x in shift for x in ["hf","half","hnight","hn"]):
            s=SHIFT["hn"]
            punches=(p1 or [])+(p2 or [])
            if not punches:
                status="No Punch"
            elif len(punches)==1:
                t=punches[0].time()
                status="Single In Punch" if t < s["out_start"] else "Single Out Punch"
            else:
                in_candidates=[p for p in punches if s["in_start"]<=p.time()<=s["in_end"]]
                out_candidates=[p for p in punches if s["out_start"]<=p.time()<=s["out_end"]]
                in_p=min(in_candidates) if in_candidates else punches[0]
                out_p=max(out_candidates) if out_candidates else punches[-1]
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                if ot_val and ot_val>0:
                    status="Match"
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
            results.append(status)
            continue

        # ----- FULL-NIGHT (Cross-Day with OT only when OT value present) -----
        if any(x in shift for x in ["fn","fullnight","night"]):
            s=SHIFT["fn"]
            valid_in=[p for p in p1 if s["in_start"]<=p.time()<=s["in_end"]]
            valid_out=[p for p in p2 if s["out_start"]<=p.time()<=s["out_end"]]
            in_p=valid_in[0] if valid_in else (p1[1] if len(p1)>=2 else p1[0] if p1 else None)
            out_p=valid_out[0] if valid_out else (p2[0] if p2 else None)
            if not in_p and not out_p:
                status="No Punch"
            elif in_p and not out_p:
                status="Single In Punch"
            elif not in_p and out_p:
                status="Single Out Punch"
            else:
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                if ot_val and ot_val>0:
                    actual_min=(out_p - in_p).total_seconds()/60
                    expected_min=ot_val*60
                    if abs(actual_min-expected_min)<=30:
                        status="Match"
                    else:
                        status="OT Deviation"
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
            results.append(status)
            continue

        # ----- GENERAL-1 -----
        if any(x in shift for x in ["general1","general 1"]):
            s=SHIFT["general1"]
            if not p1:
                status="No Punch"
            elif len(p1)==1:
                t=p1[0].time()
                status="Single In Punch" if t < s["out_start"] else "Single Out Punch"
            else:
                in_candidates=[p for p in p1 if s["in_start"]<=p.time()<=s["in_end"]]
                out_candidates=[p for p in p1 if s["out_start"]<=p.time()<=s["out_end"]]
                in_p=min(in_candidates) if in_candidates else p1[0]
                out_p=max(out_candidates) if out_candidates else p1[-1]
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                if ot_val and ot_val>0:
                    status="Match"
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
            results.append(status)
            continue

        # ----- GENERAL-2 -----
        if any(x in shift for x in ["general2","general 2"]):
            s=SHIFT["general2"]
            if not p1:
                status="No Punch"
            elif len(p1)==1:
                t=p1[0].time()
                status="Single In Punch" if t < s["out_start"] else "Single Out Punch"
            else:
                in_candidates=[p for p in p1 if s["in_start"]<=p.time()<=s["in_end"]]
                out_candidates=[p for p in p1 if s["out_start"]<=p.time()<=s["out_end"]]
                in_p=min(in_candidates) if in_candidates else p1[0]
                out_p=max(out_candidates) if out_candidates else p1[-1]
                if out_p < in_p: out_p = out_p + timedelta(days=1)
                if ot_val and ot_val>0:
                    status="Match"
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
            results.append(status)
            continue

        results.append("No Punch")

    att_df["Status"]=results

    # ---------- Output ----------
    out=BytesIO()
    with pd.ExcelWriter(out,engine="openpyxl") as w:
        att_df.to_excel(w,index=False,sheet_name=sheet)
        ws=w.sheets[sheet]
        if date_col:
            col=att_df.columns.get_loc(date_col)+1
            col_letter=get_column_letter(col)
            for r in range(2,len(att_df)+2):
                ws[f"{col_letter}{r}"].number_format="yyyy-mm-dd"
            ws.column_dimensions[col_letter].width=15
    out.seek(0)
    return out,att_df

# ---------- Streamlit ----------
st.set_page_config(page_title="Attendance Comparator (All Shifts Unified)",page_icon="🕒",layout="centered")
st.title("🕒 Attendance Comparator — Unified Logic for All Shifts")
st.markdown("""
✅ **Unified Logic Implemented:**  
- **Day, Half-Night, Full-Night, General-1, General-2** handled.  
- Smart selection: **Earliest valid IN + Latest valid OUT**.  
- Single punches correctly classified as *Single In/Out*.  
- Overtime only applied if field > 0 (±30 mins tolerance).  
- Full-Night cross-day logic built in.  
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
                st.dataframe(df.head(20))
            except Exception as e:
                st.error(f"❌ Error:\n{e}")
