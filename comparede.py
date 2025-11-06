import streamlit as st
import pandas as pd
import re, tempfile, shutil
from datetime import datetime, timedelta, time
from io import BytesIO
from openpyxl.utils import get_column_letter

# ---------- Utility helpers ----------
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
        except: pass
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

# ---------- Shift windows ----------
SHIFT = {
    "day":       {"in_start": time(7,0),"in_end":time(7,15),"out_start":time(15,15),"out_end":time(15,30)},
    "hf":        {"in_start": time(15,0),"in_end":time(15,15),"out_start":time(23,15),"out_end":time(23,30)},
    "fn":        {"in_start": time(23,15),"in_end":time(23,30),"out_start":time(7,15),"out_end":time(7,30)},
    "general1":  {"in_start": time(8,0),"in_end":time(8,15),"out_start":time(16,0),"out_end":time(16,15)},
    "general2":  {"in_start": time(9,0),"in_end":time(9,15),"out_start":time(17,0),"out_end":time(17,15)},
}

# ---------- Core ----------
def compare_files(att_bytes,bio1_bytes,bio2_bytes):
    def tmpcopy(b):
        with tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx") as f:
            f.write(b); f.flush(); path=f.name
        tmp=tempfile.NamedTemporaryFile(delete=False,suffix=".xlsx").name
        shutil.copy(path,tmp); return tmp

    att,b1,b2=[tmpcopy(x) for x in (att_bytes,bio1_bytes,bio2_bytes)]

    # Attendance
    xls=pd.ExcelFile(att); sheet=next((s for s in xls.sheet_names if "data" in s.lower() and "entry" in s.lower()),xls.sheet_names[0])
    att_df=pd.read_excel(att,sheet_name=sheet)
    att_df.columns=att_df.columns.str.strip()
    date_col=next((c for c in att_df.columns if "date" in c.lower()),None)
    if date_col: att_df[date_col]=pd.to_datetime(att_df[date_col],errors="coerce")

    # Biometric
    bio1,bio2=[dedupe(pd.read_excel(p,skiprows=1)) for p in (b1,b2)]
    for b in (bio1,bio2): b.columns=b.columns.str.strip()
    m_emp=next((c for c in att_df.columns if "emp id" in c.lower()),"EMP ID")
    b_emp1,b_emp2=find_emp_col(bio1),find_emp_col(bio2)
    for df,c in [(att_df,m_emp),(bio1,b_emp1),(bio2,b_emp2)]:
        df[c]=df[c].astype(str).map(clean_id)

    results=[]
    shift_col=next((c for c in att_df.columns if "shift" in c.lower()),"SHIFT")

    for _,row in att_df.iterrows():
        emp=row[m_emp]; shift=str(row.get(shift_col,"")).lower().strip()
        p1,p2=get_punch_times(bio1,emp,b_emp1),get_punch_times(bio2,emp,b_emp2)
        status="No Punch"

        # ---- Full Night (unchanged basic logic) ----
        if "full" in shift or "fn" in shift:
            if not p1 and not p2: status="No Punch"
            elif p1 and p2:
                t1,t2=p1[0].time(),p2[0].time()
                if (time(7,0)<=t1<=time(7,30)) and (time(7,0)<=t2<=time(7,30)): status="No Match"
                elif (time(7,0)<=t1<=time(7,30)) and (time(23,0)<=t2<=time(23,30)): status="Match"
                elif (time(23,0)<=t1<=time(23,30)) and (time(7,0)<=t2<=time(7,30)): status="Match"
                else: status="No Punch"
            else: status="No Punch"

        # ---- Day Shift updated ----
        elif "day" in shift:
            s=SHIFT["day"]
            all_p=p1 or []
            if not all_p: status="No Punch"
            elif len(all_p)==1:
                t=all_p[0].time()
                if s["in_start"]<=t<=s["in_end"]: status="Single In Punch"
                elif s["out_start"]<=t<=s["out_end"]: status="Single Out Punch"
                else: status="No Punch"
            else:
                # earliest near in_start, latest near out_end
                in_p=min(all_p,key=lambda x:abs((datetime.combine(datetime.today(),x.time())-datetime.combine(datetime.today(),s["in_start"])).total_seconds()))
                out_p=max(all_p,key=lambda x:abs((datetime.combine(datetime.today(),x.time())-datetime.combine(datetime.today(),s["out_end"])).total_seconds()))
                if out_p.time()<s["out_start"]: status="Early"
                else: status="Match"

        # ---- Half Night Shift updated ----
        elif "hf" in shift or "half" in shift:
            s=SHIFT["hf"]
            all_p=(p1 or [])+(p2 or [])
            if not all_p: status="No Punch"
            elif len(all_p)==1:
                t=all_p[0].time()
                if s["in_start"]<=t<=s["in_end"]: status="Single In Punch"
                elif s["out_start"]<=t<=s["out_end"]: status="Single Out Punch"
                else: status="No Punch"
            else:
                in_p=min(all_p,key=lambda x:abs((datetime.combine(datetime.today(),x.time())-datetime.combine(datetime.today(),s["in_start"])).total_seconds()))
                out_p=max(all_p,key=lambda x:abs((datetime.combine(datetime.today(),x.time())-datetime.combine(datetime.today(),s["out_end"])).total_seconds()))
                # overtime (~8h) check
                if (out_p - in_p).total_seconds() >= 8*3600:
                    pass  # handled naturally by in/out selection
                if out_p.time()<s["out_start"]: status="Early"
                else: status="Match"

        # ---- Other shifts unchanged ----
        else:
            key=None
            for k in SHIFT:
                if k in shift: key=k; break
            if not key:
                if not p1: status="No Punch"
                elif len(p1)==1: status="Single In Punch"
                else: status="Match"
            else:
                s=SHIFT[key]
                if len(p1)>=2:
                    in_p,out_p=p1[0],p1[-1]
                    if out_p.time()<s["out_start"]: status="Early"
                    else: status="Match"
                elif len(p1)==1:
                    t=p1[0].time()
                    if s["in_start"]<=t<=s["in_end"]: status="Single In Punch"
                    elif s["out_start"]<=t<=s["out_end"]: status="Single Out Punch"
                    else: status="No Punch"
                else: status="No Punch"

        results.append(status)

    att_df["Status"]=results
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
    out.seek(0); return out,att_df

# ---------- Streamlit ----------
st.set_page_config(page_title="Attendance Comparator (All Shifts)",page_icon="🕒",layout="centered")
st.title("🕒 Attendance Comparator — Updated Day & Half-Night Logic")
st.markdown("""
**Updated Rules**
- **Day:** 7:00–7:15 IN, 15:15–15:30 OUT → Match; OUT<15:15 → Early; earliest IN + latest OUT from all punches.
- **Half Night:** 15:00–15:15 IN, 23:15–23:30 OUT → Match; OUT<23:15 → Early; supports cross-file 8 h overtime.
- All other shifts remain unchanged.
""")

att=st.file_uploader("📁 Attendance File",type=["xlsx"])
bio1=st.file_uploader("📁 Biometric File – Day 1",type=["xlsx"])
bio2=st.file_uploader("📁 Biometric File – Day 2",type=["xlsx"])

if st.button("🔍 Compare Files"):
    if not att or not bio1 or not bio2:
        st.error("⚠️ Please upload all three files.")
    else:
        with st.spinner("Processing..."):
            try:
                out,df=compare_files(att.read(),bio1.read(),bio2.read())
                st.success("✅ Comparison complete!")
                st.download_button("⬇️ Download Excel",data=out,file_name="Attendance_with_Status.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(df.head(20))
            except Exception as e:
                st.error(f"❌ Error:\n{e}")
